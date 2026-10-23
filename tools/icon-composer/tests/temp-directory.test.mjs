import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, realpath, rm, symlink, unlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, relative, isAbsolute, sep } from 'node:path';
import test from 'node:test';
import { ensureSafeDirectory } from '../scripts/output-safety.mjs';
import { createTemporaryTestDirectory } from './temp-directory.mjs';

test('canonical test fixtures work through a temporary-directory alias while output links remain rejected', async (t) => {
  const tempRoot = await realpath(tmpdir());
  const sandbox = await realpath(await mkdtemp(join(tempRoot, 'oe-temp-path-regression-')));
  const ownedPath = relative(tempRoot, sandbox);
  assert.ok(ownedPath && ownedPath !== '..' && !ownedPath.startsWith(`..${sep}`) && !isAbsolute(ownedPath));
  const aliases = [];
  t.after(async () => {
    for (const alias of aliases.reverse()) await unlink(alias);
    await rm(sandbox, { recursive: true, force: true });
  });
  const actualParent = join(sandbox, 'actual-temp');
  const aliasParent = join(sandbox, 'alias-temp');
  await mkdir(actualParent);
  await symlink(actualParent, aliasParent, 'junction');
  aliases.push(aliasParent);
  const fixture = await createTemporaryTestDirectory(aliasParent);
  assert.equal(fixture, await realpath(fixture), 'fixture must use canonical spelling before exercising strict output safety');
  const repo = join(fixture, 'repo');
  const foreign = join(sandbox, 'foreign');
  await mkdir(repo);
  await mkdir(foreign);
  const marker = join(foreign, 'keep.marker');
  await writeFile(marker, 'unchanged');
  await ensureSafeDirectory(join(repo, 'artifacts', 'production'), repo);

  const linkedRoot = join(sandbox, 'linked-repo');
  await symlink(repo, linkedRoot, 'junction');
  aliases.push(linkedRoot);
  await assert.rejects(ensureSafeDirectory(join(linkedRoot, 'artifacts', 'production'), linkedRoot), /Repository root cannot be a symlink or junction/);
  const linkedOutput = join(repo, 'artifacts', 'linked-output');
  await symlink(foreign, linkedOutput, 'junction');
  aliases.push(linkedOutput);
  await assert.rejects(ensureSafeDirectory(join(linkedOutput, 'nested'), repo), /Output path component must be a real directory/);
  assert.equal(await readFile(marker, 'utf8'), 'unchanged');
});
