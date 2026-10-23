import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, writeFile, unlink, rmdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

/** Windows authoring profile; native editor documents keep their editable masks. */
export async function exportForOpenEfficiency(svg: string, exporter: string): Promise<string> {
  if (process.platform !== 'win32') throw new Error('The openefficiency profile requires Windows WPF.');
  if (!exporter) throw new Error('The openefficiency profile requires --exporter pointing to Export-IconComposerSvg.ps1.');
  const directory = await mkdtemp(join(tmpdir(), 'oe-icon-export-'));
  const source = join(directory, 'source.svg'), output = join(directory, 'production.svg');
  try {
    await writeFile(source, svg);
    const host = join(process.env.WINDIR ?? 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
    const result = spawnSync(host, ['-NoProfile', '-STA', '-File', resolve(exporter), '-InputPath', source, '-OutputPath', output], {
      encoding: 'utf8', timeout: 120000, windowsHide: true,
    });
    if (result.error || result.status !== 0) throw new Error(`WPF icon export failed: ${result.error?.message ?? result.stderr.trim()}`);
    return await readFile(output, 'utf8');
  } finally {
    await Promise.all([source, output].map(file => unlink(file).catch(error => { if (error.code !== 'ENOENT') throw error; })));
    await rmdir(directory);
  }
}
