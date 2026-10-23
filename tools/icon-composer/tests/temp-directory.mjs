import { mkdtemp, realpath } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

export async function createTemporaryTestDirectory(parent = tmpdir()) {
  // Windows TEMP may use an 8.3 spelling (for example RUNNER~1). Canonicalize
  // this owned fixture before it is used as a strict fake repository root.
  return realpath(await mkdtemp(join(parent, 'oe-icon-composer-test-')));
}
