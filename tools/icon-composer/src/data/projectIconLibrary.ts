import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

interface KitEntry { id: string; file: string; semanticRole: string; }

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../kit');
const kit = JSON.parse(readFileSync(resolve(root, 'kit.json'), 'utf8')) as { version: number; entries: KitEntry[] };
export const PROJECT_KIT_VERSION = kit.version;

export const PROJECT_ICON_LIBRARY = kit.entries.map((entry) => ({
  id: entry.id,
  name: entry.id.split('-').map((word) => word[0].toUpperCase() + word.slice(1)).join(' '),
  category: 'general' as const,
  tags: [entry.semanticRole, entry.file.startsWith('actions/') ? 'action' : 'object'],
  svgContent: readFileSync(resolve(root, entry.file), 'utf8'),
  license: 'Approved icon source; original terms in THIRD_PARTY_NOTICES.md and kit provenance',
}));

export function getKitIcon(id: string): (typeof PROJECT_ICON_LIBRARY)[number] {
  const icon = PROJECT_ICON_LIBRARY.find((entry) => entry.id === id);
  if (!icon) throw new Error(`Unknown approved kit icon: ${id}`);
  return icon;
}
