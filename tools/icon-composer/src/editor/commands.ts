import { EditorDriver } from './driver';

/** Explicit allowlist: never dispatch arbitrary object/prototype properties. */
export const OPERATIONS = {
  describe: [], list: [], ids: [], get: ['id'], toSvg: [],
  reset: ['width?', 'height?'], clear: [], loadDocument: ['document'],
  setCanvasSize: ['width', 'height'], setViewBox: ['x', 'y', 'w', 'h'],
  setConfig: ['patch'], setGrid: ['options'],
  addShape: ['type', 'props?'], addImportedSvg: ['svg', 'name?', 'x?', 'y?'],
  addComposedSvg: ['svg', 'name?'], addLibraryIcon: ['iconId', 'props?'],
  compose: ['config', 'name?'], recompose: ['id', 'patch'], searchIcons: ['query?'], applySvgCode: ['svg'],
  update: ['id', 'patch'], move: ['id', 'x', 'y'], moveBy: ['id', 'dx', 'dy'],
  resize: ['id', 'width', 'height?'], scale: ['id', 'factor'],
  setRotation: ['id', 'degrees'], setStyle: ['id', 'style'],
  rename: ['id', 'name'], setVisibility: ['id', 'visible'], setLock: ['id', 'locked'],
  duplicate: ['id', 'newId?'], remove: ['id'], reorder: ['id', 'position'],
  reorderRelative: ['id', 'direction'], align: ['id', 'alignment'], undo: [], redo: [],
} as const;

export type EditorOperation = keyof typeof OPERATIONS;
export type EditorCommand = {
  [K in EditorOperation]: { op: K; args?: Parameters<EditorDriver[K]> }
}[EditorOperation];

export function executeCommand(driver: EditorDriver, command: EditorCommand): unknown {
  if (!command || !Object.hasOwn(OPERATIONS, command.op)) throw new Error(`Unknown editor operation: ${command?.op}`);
  const args = command.args ?? [];
  if (!Array.isArray(args)) throw new Error('Command args must be an array');
  const signature = OPERATIONS[command.op] as readonly string[];
  const required = signature.filter((name) => !name.endsWith('?')).length;
  if (args.length < required || args.length > signature.length) throw new Error(`${command.op} expects (${signature.join(', ')})`);
  const operation = driver[command.op] as (...args: unknown[]) => unknown;
  return operation.apply(driver, args) ?? null;
}

export function executeBatch(driver: EditorDriver, commands: EditorCommand[]): unknown[] {
  if (!Array.isArray(commands)) throw new Error('commands must be an array');
  if (commands.some((command) => command.op === 'undo' || command.op === 'redo')) throw new Error('Use undo/redo individually, outside an atomic batch');
  const run = () => commands.map((command, index) => {
    try { return executeCommand(driver, command); }
    catch (error) { throw new Error(`Command ${index} (${command?.op}): ${error instanceof Error ? error.message : String(error)}`); }
  });
  return commands.every((command) => ['describe', 'list', 'ids', 'get', 'toSvg', 'searchIcons'].includes(command.op)) ? run() : driver.transaction(run);
}
