import { lstat, mkdir, mkdtemp, realpath } from 'node:fs/promises';
import { dirname, isAbsolute, join, relative, resolve, sep } from 'node:path';

const normalize = (path) => process.platform === 'win32' ? path.toLowerCase() : path;

export function resolveApprovedOutputBase(requested, repositoryRoot, relativeBase) {
  if (!requested) return undefined;
  const repo = resolve(repositoryRoot);
  const expected = resolve(repo, relativeBase);
  if (!isAbsolute(requested) || normalize(resolve(requested)) !== normalize(expected)) {
    throw new Error(`Output base must be the approved repository artifact directory: ${expected}`);
  }
  const relativeBasePath = relative(repo, expected);
  if (!relativeBasePath || relativeBasePath === '.' || relativeBasePath === '..' || relativeBasePath.startsWith(`..${sep}`) || isAbsolute(relativeBasePath)) {
    throw new Error('Output base must be a child of the repository root');
  }
  return expected;
}

export function resolveApprovedFilePath(requested, repositoryRoot, relativeFile) {
  if (!requested) return undefined;
  const expected = resolve(resolve(repositoryRoot), relativeFile);
  if (!isAbsolute(requested) || normalize(resolve(requested)) !== normalize(expected)) {
    throw new Error(`Output file must be the approved repository artifact file: ${expected}`);
  }
  return expected;
}

export async function ensureSafeDirectory(base, repositoryRoot) {
  const repo = resolve(repositoryRoot);
  const target = resolve(base);
  const childPath = relative(repo, target);
  if (!childPath || childPath === '.' || childPath === '..' || childPath.startsWith(`..${sep}`) || isAbsolute(childPath)) {
    throw new Error('Output base must be a child of the repository root');
  }

  const realRepo = await realpath(repo);
  if (normalize(realRepo) !== normalize(repo)) throw new Error('Repository root cannot be a symlink or junction');

  let current = realRepo;
  for (const component of childPath.split(sep)) {
    current = join(current, component);
    try {
      await mkdir(current);
    } catch (error) {
      if (error?.code !== 'EEXIST') throw error;
    }
    const stat = await lstat(current);
    if (!stat.isDirectory() || stat.isSymbolicLink()) throw new Error(`Output path component must be a real directory: ${current}`);
    if (normalize(await realpath(current)) !== normalize(current)) throw new Error(`Output path component cannot be a symlink or junction: ${current}`);
  }
  return target;
}

export async function createUniqueOutputDirectory(base, repositoryRoot) {
  const target = await ensureSafeDirectory(base, repositoryRoot);
  // The actual output destination is always a fresh child. No previous run or
  // caller-controlled directory is recursively removed or overwritten.
  return mkdtemp(join(target, 'run-'));
}
