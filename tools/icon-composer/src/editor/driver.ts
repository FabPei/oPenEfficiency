/**
 * Headless driver for the whole SVG editor.
 *
 * Mirrors every document operation of the React editor (element CRUD,
 * transforms, styling, layer order, alignment, canvas config, SVG-code
 * apply, export) with the same semantics and the same shared model —
 * so a coding AI can fully steer the editor programmatically without
 * mouse control. Undo/redo history works like the editor's.
 */
import type { CanvasConfig, SvgElement } from '../types/svg';
import { sanitizeSvg, serializeSvgDocument } from '../utils/svg';
import { composeIcon, defaultParams, readComposerMetadata } from '../composer/compose';
import type { ComposerConfig, ComposerParams } from '../composer/types';
import { PROJECT_ICON_LIBRARY } from '../data/projectIconLibrary';
import {
  alignPosition,
  createElement,
  defaultConfig,
  importedElement,
  initialDocument,
  presetGridSize,
  type AlignType,
  type DrawToolType,
} from './model';

export interface ElementSummary {
  id: string;
  name: string;
  type: SvgElement['type'];
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  visible: boolean;
  locked: boolean;
}

export interface DriverSnapshot {
  config: CanvasConfig;
  elements: SvgElement[];
}

const cloneSnapshot = (snapshot: DriverSnapshot): DriverSnapshot => structuredClone(snapshot);

const positive = (value: number, label: string) => {
  if (!Number.isFinite(value) || value <= 0) throw new Error(`${label} must be finite and positive`);
};

// The source-repository bundle exposes only the owner's approved project kit.
export const ICON_LIBRARY = PROJECT_ICON_LIBRARY;

export class EditorDriver {
  private snapshot: DriverSnapshot;
  private history: DriverSnapshot[] = [];
  private historyIndex = -1;
  private seq = 0;
  private listeners = new Set<() => void>();
  private revision = 0;
  private batching = false;

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  };

  getRevision = (): number => this.revision;

  private notify(): void {
    if (this.batching) return;
    this.revision += 1;
    this.listeners.forEach((listener) => listener());
  }

  constructor(initialSvg?: string) {
    if (initialSvg?.trim()) sanitizeSvg(initialSvg);
    this.snapshot = initialDocument(initialSvg);
    this.commit();
  }

  // --- history (same shape as the editor's pushState/undo/redo) --------------

  private commit(): void {
    this.history = this.history.slice(0, this.historyIndex + 1);
    this.history.push(cloneSnapshot(this.snapshot));
    this.historyIndex = this.history.length - 1;
    this.notify();
  }

  undo(): boolean {
    if (this.historyIndex <= 0) return false;
    this.historyIndex -= 1;
    this.snapshot = cloneSnapshot(this.history[this.historyIndex]);
    this.notify();
    return true;
  }

  redo(): boolean {
    if (this.historyIndex >= this.history.length - 1) return false;
    this.historyIndex += 1;
    this.snapshot = cloneSnapshot(this.history[this.historyIndex]);
    this.notify();
    return true;
  }

  get canUndo(): boolean {
    return this.historyIndex > 0;
  }

  get canRedo(): boolean {
    return this.historyIndex < this.history.length - 1;
  }

  // --- canvas configuration ---------------------------------------------------

  get config(): CanvasConfig {
    return { ...this.snapshot.config, viewBox: { ...this.snapshot.config.viewBox } };
  }

  /** Set canvas + viewBox size (like the editor's dimension preset input). */
  setCanvasSize(width: number, height: number): void {
    positive(width, 'width');
    positive(height, 'height');
    this.snapshot.config = {
      ...this.snapshot.config,
      width,
      height,
      viewBox: { x: 0, y: 0, w: width, h: height },
      gridSize: presetGridSize(width),
    };
    this.commit();
  }

  setViewBox(x: number, y: number, w: number, h: number): void {
    positive(w, 'viewBox width');
    positive(h, 'viewBox height');
    if (!Number.isFinite(x) || !Number.isFinite(y)) throw new Error('viewBox origin must be finite');
    this.snapshot.config = { ...this.snapshot.config, viewBox: { x, y, w, h } };
    this.commit();
  }

  setGrid(options: { visible?: boolean; snap?: boolean; size?: number }): void {
    this.setConfig({
      ...(options.visible === undefined ? {} : { gridVisible: options.visible }),
      ...(options.snap === undefined ? {} : { snapToGrid: options.snap }),
      ...(options.size === undefined ? {} : { gridSize: options.size }),
    });
  }

  // --- element creation --------------------------------------------------------

  /** Create a shape via the same factory the canvas drawing tools use. */
  addShape(type: DrawToolType, props: Partial<SvgElement> = {}): string {
    if (!['rect', 'circle', 'ellipse', 'line', 'arrow', 'path', 'text', 'star', 'polygon'].includes(type)) throw new Error(`Unknown drawing tool: ${type}`);
    this.seq += 1;
    const element = createElement(type, props.x ?? 2, props.y ?? 2, this.seq, props.id);
    const next: SvgElement = { ...element, ...props, id: element.id, type };
    if (type === 'path' && !props.pathViewBox) next.pathViewBox = { x: 0, y: 0, w: next.width || 1, h: next.height || 1 };
    this.replaceElements([...this.snapshot.elements, next]);
    return next.id;
  }

  /** Import an SVG file/string as an editable raw element (editor import path). */
  addImportedSvg(svgSource: string, name = 'Imported SVG', x = 0, y = 0): string {
    const element = importedElement(svgSource, name, x, y);
    this.snapshot.elements = [...this.snapshot.elements, element];
    this.commit();
    return element.id;
  }

  /** Add an already composed Action Icon Composer SVG (undoable, like the panel). */
  addComposedSvg(composedSvg: string, name = 'Composer Icon'): string {
    return this.addImportedSvg(composedSvg, name);
  }

  // --- element mutation ---------------------------------------------------------

  private find(id: string): SvgElement {
    const element = this.snapshot.elements.find((candidate) => candidate.id === id);
    if (!element) {
      throw new Error(`Unknown element id "${id}". Known ids: ${this.ids().join(', ') || '(none)'}`);
    }
    return element;
  }

  ids(): string[] {
    return this.snapshot.elements.map((element) => element.id);
  }

  list(): ElementSummary[] {
    return this.snapshot.elements.map((element) => ({
      id: element.id,
      name: element.name,
      type: element.type,
      x: element.x,
      y: element.y,
      width: element.width,
      height: element.height,
      rotation: element.rotation,
      visible: element.visible,
      locked: element.locked,
    }));
  }

  get(id: string): SvgElement {
    return structuredClone(this.find(id));
  }

  /** Patch any element fields (position, size, paint, text, path data, …). */
  update(id: string, patch: Partial<SvgElement>, options: { commit?: boolean } = {}): void {
    const element = this.find(id);
    if (element.locked && patch.locked !== false && Object.keys(patch).some((key) => key !== 'locked' && key !== 'visible' && key !== 'name')) {
      throw new Error(`Element "${id}" is locked. Unlock it first (setLock), or include "locked": false in the patch.`);
    }
    const next = { ...element, ...patch, id: element.id };
    this.replaceElements(this.snapshot.elements.map((candidate) =>
      candidate.id === id ? next : candidate,
    ), options.commit !== false);
  }

  move(id: string, x: number, y: number): void {
    this.update(id, { x, y });
  }

  moveBy(id: string, dx: number, dy: number): void {
    const element = this.find(id);
    this.update(id, { x: element.x + dx, y: element.y + dy });
  }

  /** Resize keeping the top-left corner (editor resize semantics). */
  resize(id: string, width: number, height = width): void {
    this.update(id, { width: Math.max(1, width), height: Math.max(1, height) });
  }

  /** Resize keeping the center (typical "scale" feel). */
  scale(id: string, factor: number): void {
    positive(factor, 'scale factor');
    const element = this.find(id);
    const width = Math.max(1, element.width * factor);
    const height = Math.max(1, element.height * factor);
    this.update(id, {
      x: element.x + (element.width - width) / 2,
      y: element.y + (element.height - height) / 2,
      width,
      height,
    });
  }

  setRotation(id: string, rotation: number): void {
    this.update(id, { rotation });
  }

  setStyle(id: string, style: Partial<SvgElement>): void {
    this.update(id, style);
  }

  rename(id: string, name: string): void {
    this.update(id, { name });
  }

  setVisibility(id: string, visible: boolean): void {
    this.update(id, { visible });
  }

  setLock(id: string, locked: boolean): void {
    this.update(id, { locked });
  }

  duplicate(id: string, newId?: string): string {
    const element = this.find(id);
    this.seq += 1;
    const copy: SvgElement = {
      ...element,
      id: newId ?? `el-${Date.now()}-${this.seq}`,
      name: `${element.name} Copy`,
      x: element.x + 2,
      y: element.y + 2,
    };
    this.replaceElements([...this.snapshot.elements, copy]);
    return copy.id;
  }

  remove(id: string): void {
    this.find(id);
    this.snapshot.elements = this.snapshot.elements.filter((element) => element.id !== id);
    this.commit();
  }

  /** Move an element in the layer stack (front = last in the list = on top). */
  reorder(id: string, position: 'front' | 'back' | number): void {
    if (position !== 'front' && position !== 'back' && !Number.isInteger(position)) throw new Error('Layer position must be front, back or an integer');
    const element = this.find(id);
    const rest = this.snapshot.elements.filter((candidate) => candidate.id !== id);
    const index =
      position === 'front'
        ? rest.length
        : position === 'back'
          ? 0
          : Math.max(0, Math.min(rest.length, position));
    this.snapshot.elements = [...rest.slice(0, index), element, ...rest.slice(index)];
    this.commit();
  }

  /** Bring to front / send to back / one step up or down. */
  reorderRelative(id: string, direction: 'front' | 'back' | 'up' | 'down'): void {
    if (!['front', 'back', 'up', 'down'].includes(direction)) throw new Error(`Unknown layer direction: ${direction}`);
    this.find(id);
    const index = this.snapshot.elements.findIndex((element) => element.id === id);
    const length = this.snapshot.elements.length;
    if (direction === 'front') return this.reorder(id, 'front');
    if (direction === 'back') return this.reorder(id, 'back');
    const target = direction === 'up' ? index + 1 : index - 1;
    if (target < 0 || target >= length) return;
    this.reorder(id, target);
  }

  align(id: string, type: AlignType): void {
    if (!['left', 'centerH', 'right', 'top', 'centerV', 'bottom'].includes(type)) throw new Error(`Unknown alignment: ${type}`);
    const element = this.find(id);
    if (element.locked) return;
    const position = alignPosition(element, type, this.snapshot.config.viewBox);
    this.update(id, position);
  }

  // --- document-level operations -------------------------------------------------

  /**
   * Apply raw SVG code to the canvas (editor "Apply" in the code panel):
   * replaces the document content and resizes the canvas to the source.
   */
  applySvgCode(source: string): string {
    const safe = sanitizeSvg(source);
    const element = importedElement(safe.content, 'SVG Code');
    this.snapshot = {
      config: {
        ...this.snapshot.config,
        width: safe.width,
        height: safe.height,
        viewBox: { x: 0, y: 0, w: safe.width, h: safe.height },
        gridSize: safe.width <= 64 ? 1 : Math.max(2, Math.round(safe.width / 32)),
      },
      elements: [element],
    };
    this.commit();
    return element.id;
  }

  /** Serialize the whole document exactly like the editor's SVG output. */
  toSvg(): string {
    return serializeSvgDocument(this.snapshot.elements, this.snapshot.config);
  }

  /** Machine-readable full state (for inspection by the AI). */
  describe(): DriverSnapshot {
    return cloneSnapshot(this.snapshot);
  }

  /** Replace a native editable document, preserving all layer properties. */
  loadDocument(document: DriverSnapshot, commit = true): void {
    const next = cloneSnapshot(document);
    const c = next.config;
    positive(c.width, 'width'); positive(c.height, 'height'); positive(c.gridSize, 'grid size');
    positive(c.viewBox.w, 'viewBox width'); positive(c.viewBox.h, 'viewBox height');
    if (![c.viewBox.x, c.viewBox.y].every(Number.isFinite)) throw new Error('viewBox origin must be finite');
    if (typeof c.backgroundColor !== 'string' || typeof c.gridVisible !== 'boolean' || typeof c.snapToGrid !== 'boolean') throw new Error('Invalid canvas settings');
    const ids = new Set<string>();
    const validateNumbers = (value: unknown): void => {
      if (typeof value === 'number' && !Number.isFinite(value)) throw new Error('Shape numbers must be finite');
      if (value && typeof value === 'object') Object.values(value).forEach(validateNumbers);
    };
    const validate = (elements: SvgElement[]) => {
      for (const element of elements) {
        validateNumbers(element);
        if (!element.id || ids.has(element.id)) throw new Error(`Duplicate or empty element id: ${element.id}`);
        ids.add(element.id);
        if (!['rect', 'circle', 'ellipse', 'line', 'arrow', 'path', 'text', 'star', 'polygon', 'raw_svg'].includes(element.type)) throw new Error(`Unsupported shape: ${element.type}`);
        for (const [key, value] of Object.entries(element)) {
          if (typeof value === 'number' && !Number.isFinite(value)) throw new Error(`${key} must be finite`);
        }
        if (![element.x, element.y, element.width, element.height, element.rotation, element.fillOpacity, element.strokeWidth, element.strokeOpacity].every(Number.isFinite)) throw new Error('Missing numeric shape properties');
        if (element.width < 0 || element.height < 0 || element.strokeWidth < 0) throw new Error('Shape dimensions and stroke width cannot be negative');
        if (typeof element.visible !== 'boolean' || typeof element.locked !== 'boolean' || typeof element.name !== 'string' || typeof element.fill !== 'string' || typeof element.stroke !== 'string') throw new Error('Invalid shape properties');
        if (element.type === 'raw_svg' && !element.rawSvgContent) throw new Error('Imported layers require SVG content');
        if (element.rawSvgContent) element.rawSvgContent = sanitizeSvg(element.rawSvgContent).content;
        if (element.pathViewBox) {
          positive(element.pathViewBox.w, 'path viewBox width'); positive(element.pathViewBox.h, 'path viewBox height');
          if (![element.pathViewBox.x, element.pathViewBox.y].every(Number.isFinite)) throw new Error('path origin must be finite');
        }
        if (element.children) validate(element.children);
      }
    };
    validate(next.elements);
    this.snapshot = next;
    if (commit) this.commit(); else this.notify();
  }

  replaceElements(elements: SvgElement[], commit = true): void {
    this.loadDocument({ config: this.snapshot.config, elements }, commit);
  }

  setConfig(patch: Partial<CanvasConfig>): void {
    this.loadDocument({ config: { ...this.snapshot.config, ...patch }, elements: this.snapshot.elements });
  }

  commitInteraction(): void { this.commit(); }

  clear(): void { this.replaceElements([]); }

  reset(width = 24, height = width): void {
    positive(width, 'width'); positive(height, 'height');
    this.loadDocument({ config: { ...defaultConfig(), width, height, viewBox: { x: 0, y: 0, w: width, h: height }, gridSize: presetGridSize(width) }, elements: [] });
  }

  /** Atomic batch: one undo entry; a failed operation leaves state/history intact. */
  transaction<T>(work: () => T): T {
    if (this.batching) return work();
    const before = this.describe();
    const history = this.history.slice();
    const index = this.historyIndex;
    this.batching = true;
    try {
      const result = work();
      this.history = history;
      this.historyIndex = index;
      this.batching = false;
      this.commit();
      return result;
    } catch (error) {
      this.snapshot = before;
      this.history = history;
      this.historyIndex = index;
      this.batching = false;
      throw error;
    }
  }

  searchIcons(query = '') {
    const terms = query.toLowerCase().split(/\s+/).filter(Boolean);
    return structuredClone(ICON_LIBRARY.filter((icon) => terms.every((term) =>
      `${icon.id} ${icon.name} ${icon.category} ${icon.tags.join(' ')}`.toLowerCase().includes(term))));
  }

  addLibraryIcon(iconId: string, props: Partial<SvgElement> = {}): string {
    const icon = ICON_LIBRARY.find((entry) => entry.id === iconId);
    if (!icon) throw new Error(`Unknown library icon: ${iconId}`);
    const imported = importedElement(icon.svgContent, icon.name);
    const element: SvgElement = { ...imported, ...props, id: props.id ?? imported.id, type: 'raw_svg' };
    this.replaceElements([...this.snapshot.elements, element]);
    return element.id;
  }

  compose(config: Omit<ComposerConfig, 'params'> & { params?: Partial<ComposerParams> }, name = 'Composer Icon'): string {
    return this.addComposedSvg(composeIcon({ ...config, params: { ...defaultParams(), ...config.params } }).svg, name);
  }

  recompose(id: string, patch: Partial<Omit<ComposerConfig, 'params'>> & { params?: Partial<ComposerParams> }): void {
    const element = this.find(id);
    const stored = readComposerMetadata(element.rawSvgContent ?? '');
    if (!stored) throw new Error(`Element ${id} has no composer metadata`);
    const composed = composeIcon({ ...stored, ...patch, params: { ...stored.params, ...patch.params } });
    this.update(id, { rawSvgContent: composed.svg });
  }
}
