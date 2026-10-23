import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { resolve, dirname } from 'node:path';
import { parseArgs } from 'node:util';
import { JSDOM } from 'jsdom';
import { Resvg } from '@resvg/resvg-js';
import { exportForOpenEfficiency } from './openefficiency-export';
import { EditorDriver, executeBatch, OPERATIONS, type EditorCommand, type DriverSnapshot } from '../src/editor/index';
import { ACTION_POSITIONS } from '../src/composer/types';
import { getKitIcon, PROJECT_KIT_VERSION, PROJECT_ICON_LIBRARY } from '../src/data/projectIconLibrary';

const dom = new JSDOM('');
Object.assign(globalThis, { DOMParser: dom.window.DOMParser, XMLSerializer: dom.window.XMLSerializer });

try {
  const { values } = parseArgs({ options: {
    input: { type: 'string' }, commands: { type: 'string' }, output: { type: 'string' },
    document: { type: 'string' }, png: { type: 'string' }, scale: { type: 'string', default: '4' },
    capabilities: { type: 'boolean' }, help: { type: 'boolean' },
    profile: { type: 'string', default: 'editable' }, exporter: { type: 'string' }, recipe: { type: 'string' },
  } });
  if (values.help) {
    process.stdout.write('Usage: npm run editor -- [--recipe recipe.json | --input icon.svg|document.json [--commands commands.json|-]] [--output icon.svg] [--document editable.json] [--png preview.png] [--scale 4] [--profile editable|openefficiency] [--exporter Export-IconComposerSvg.ps1] [--capabilities]\n');
  } else if (values.capabilities) {
    process.stdout.write(JSON.stringify({ version: 1, operations: OPERATIONS, composer: { actionPositions: ACTION_POSITIONS }, kit: { version: PROJECT_KIT_VERSION, icons: PROJECT_ICON_LIBRARY.length, recipeFields: ['id', 'name', 'main', 'action', 'params'], recipeCommand: '--recipe' } }, null, 2) + '\n');
  } else {
    const driver = new EditorDriver();
    driver.reset();
    if (values.recipe && (values.input || values.commands)) throw new Error('--recipe cannot be combined with --input or --commands');
    if (values.input) {
      const source = await readFile(resolve(values.input), 'utf8');
      if (values.input.toLowerCase().endsWith('.json')) driver.loadDocument(JSON.parse(source) as DriverSnapshot);
      else driver.applySvgCode(source);
    }
    let results: unknown[] = [];
    if (values.recipe) {
      const recipe = JSON.parse(await readFile(resolve(values.recipe), 'utf8')) as {
        id?: string; name?: string; main: string; action: string; params?: Record<string, unknown>;
      };
      if (!recipe || typeof recipe.main !== 'string' || typeof recipe.action !== 'string') throw new Error('Recipe requires approved kit ids "main" and "action"');
      const main = getKitIcon(recipe.main);
      const action = getKitIcon(recipe.action);
      driver.clear();
      const uid = (recipe.id ?? `${recipe.main}-${recipe.action}`).replace(/[^a-zA-Z0-9_-]/g, '-');
      results = [driver.compose({ mainSvg: main.svgContent, actionSvg: action.svgContent, uid, params: recipe.params as never }, recipe.name ?? `${main.name} ${action.name}`)];
    } else if (values.commands) {
      let source: string;
      if (values.commands === '-') {
        const chunks: Buffer[] = [];
        for await (const chunk of process.stdin) chunks.push(Buffer.from(chunk));
        source = Buffer.concat(chunks).toString('utf8');
      } else source = await readFile(resolve(values.commands), 'utf8');
      const request = JSON.parse(source) as EditorCommand[] | { commands: EditorCommand[] };
      results = executeBatch(driver, Array.isArray(request) ? request : request.commands);
    }
    if (!['editable', 'openefficiency'].includes(values.profile)) throw new Error('Unknown export profile');
    if (values.exporter && values.profile !== 'openefficiency') throw new Error('--exporter requires --profile openefficiency');
    const editableSvg = driver.toSvg();
    const svg = values.profile === 'openefficiency'
      ? await exportForOpenEfficiency(editableSvg, values.exporter ?? '') : editableSvg;
    // Validate/rasterize before writing outputs, so command or render failures write no files.
    let png: Buffer | undefined;
    if (values.png) {
      const scale = Number(values.scale);
      if (!Number.isFinite(scale) || scale <= 0) throw new Error('scale must be finite and positive');
      png = new Resvg(svg, { fitTo: { mode: 'width', value: Math.max(1, Math.round(driver.config.width * scale)) } }).render().asPng();
    }
    const outputs: [string | undefined, string | Buffer][] = [
      [values.output, svg], [values.document, JSON.stringify(driver.describe(), null, 2)], [values.png, png ?? ''],
    ];
    const paths = outputs.filter(([file]) => file).map(([file]) => resolve(file!));
    if (new Set(paths.map((file) => process.platform === 'win32' ? file.toLowerCase() : file)).size !== paths.length) throw new Error('Output paths must be distinct');
    for (const [file, content] of outputs) {
      if (!file) continue;
      const target = resolve(file);
      await mkdir(dirname(target), { recursive: true });
      await writeFile(target, content);
    }
    process.stdout.write(JSON.stringify({ ok: true, results, document: driver.describe(), svg, files: paths }, null, 2) + '\n');
  }
} catch (error) {
  process.stderr.write(JSON.stringify({ ok: false, error: error instanceof Error ? error.message : String(error) }) + '\n');
  process.exitCode = 1;
} finally { dom.window.close(); }
