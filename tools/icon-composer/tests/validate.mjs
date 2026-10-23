import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { lstat, mkdir, mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, relative, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { createUniqueOutputDirectory, ensureSafeDirectory, resolveApprovedFilePath, resolveApprovedOutputBase } from '../scripts/output-safety.mjs';

const root = resolve(import.meta.dirname, '..');
const startedAt = new Date();
const started = Date.now();
const cli = [join(root, 'node_modules/tsx/dist/cli.mjs'), 'scripts/editor-cli.ts'];
const temp = await mkdtemp(join(tmpdir(), 'oe-icon-composer-test-'));
const run = (args) => spawnSync(process.execPath, [...cli, ...args], { cwd: root, encoding: 'utf8', windowsHide: true });
const hash = (value) => createHash('sha256').update(value).digest('hex');
const repositoryRoot = resolve(root, '../..');
const relativeRepoPath = (file) => relative(repositoryRoot, file).replaceAll('\\', '/');
const assertOutputSafety = async () => {
  const sandboxRepo = join(temp, 'repo');
  const foreign = join(temp, 'foreign');
  await mkdir(sandboxRepo);
  await mkdir(foreign);
  const rootMarker = join(sandboxRepo, 'keep-root.marker');
  const parentMarker = join(temp, 'keep-parent.marker');
  const foreignMarker = join(foreign, 'keep-foreign.marker');
  await writeFile(rootMarker, 'root marker');
  await writeFile(parentMarker, 'parent marker');
  await writeFile(foreignMarker, 'foreign marker');
  const base = resolve(sandboxRepo, 'artifacts/icon-composer-bundle/production');
  const baseMarker = join(base, 'keep-base.marker');
  await mkdir(base, { recursive: true });
  await writeFile(baseMarker, 'base marker');
  const priorRun = join(base, 'run-existing');
  await mkdir(priorRun);
  const priorMarker = join(priorRun, 'keep-existing.marker');
  await writeFile(priorMarker, 'existing run marker');

  assert.throws(() => resolveApprovedOutputBase(sandboxRepo, sandboxRepo, 'artifacts/icon-composer-bundle/production'), /approved repository artifact directory/);
  assert.throws(() => resolveApprovedOutputBase(foreign, sandboxRepo, 'artifacts/icon-composer-bundle/production'), /approved repository artifact directory/);
  assert.throws(() => resolveApprovedOutputBase(resolve(sandboxRepo, '../../../foreign'), sandboxRepo, 'artifacts/icon-composer-bundle/production'), /approved repository artifact directory/);
  assert.throws(() => resolveApprovedFilePath(rootMarker, sandboxRepo, 'artifacts/icon-composer-bundle/light-dark-size.json'), /approved repository artifact file/);
  assert.throws(() => resolveApprovedFilePath(foreignMarker, sandboxRepo, 'artifacts/icon-composer-bundle/light-dark-size.json'), /approved repository artifact file/);
  await assert.rejects(createUniqueOutputDirectory(sandboxRepo, sandboxRepo), /child of the repository root/);
  await assert.rejects(createUniqueOutputDirectory(resolve(sandboxRepo, '..'), sandboxRepo), /child of the repository root/);
  await assert.rejects(createUniqueOutputDirectory(foreign, sandboxRepo), /child of the repository root/);
  const approvedBase = resolveApprovedOutputBase(base, sandboxRepo, 'artifacts/icon-composer-bundle/production');
  const runDirectory = await createUniqueOutputDirectory(approvedBase, sandboxRepo);
  assert.notEqual(runDirectory, priorRun, 'new run output directory must not reuse a pre-existing destination');
  assert.equal((await stat(runDirectory)).isDirectory(), true);
  assert.equal(await readFile(rootMarker, 'utf8'), 'root marker');
  assert.equal(await readFile(parentMarker, 'utf8'), 'parent marker');
  assert.equal(await readFile(foreignMarker, 'utf8'), 'foreign marker');
  assert.equal(await readFile(baseMarker, 'utf8'), 'base marker');
  assert.equal(await readFile(priorMarker, 'utf8'), 'existing run marker');
  const reportPath = resolveApprovedFilePath(join(sandboxRepo, 'artifacts/icon-composer-bundle/light-dark-size.json'), sandboxRepo, 'artifacts/icon-composer-bundle/light-dark-size.json');
  await ensureSafeDirectory(resolve(reportPath, '..'), sandboxRepo);
  await writeFile(reportPath, 'test report');
  const reportStat = await lstat(reportPath);
  assert.equal(reportStat.isFile(), true);
  assert.equal(await readFile(rootMarker, 'utf8'), 'root marker');
};
const manifest = JSON.parse(await readFile(resolve(root, 'kit/kit.json'), 'utf8'));
const kitSource = async (id) => {
  const entry = manifest.entries.find((candidate) => candidate.id === id);
  assert.ok(entry, `approved kit entry exists: ${id}`);
  return readFile(resolve(root, 'kit', entry.file), 'utf8');
};
for (const entry of manifest.entries) {
  assert.equal(hash(await readFile(resolve(root, 'kit', entry.file))), entry.sha256, `approved kit source hash: ${entry.id}`);
}
for (const entry of manifest.entries) {
  const file = resolve(root, 'kit', entry.file);
  const relative = file.slice(resolve(root, '../..').length + 1).replaceAll('\\', '/');
  const attr = spawnSync('git', ['check-attr', 'eol', '--', relative], { cwd: resolve(root, '../..'), encoding: 'utf8' });
  assert.equal(attr.status, 0, attr.stderr);
  assert.match(attr.stdout, /: eol: lf$/m, `kit checkout EOL attribute: ${entry.id}`);
  const blob = spawnSync('git', ['hash-object', '-w', `--path=${relative}`, relative], { cwd: resolve(root, '../..'), encoding: 'utf8' });
  assert.equal(blob.status, 0, blob.stderr);
  const checkout = spawnSync('git', ['cat-file', '--filters', `--path=${relative}`, blob.stdout.trim()], { cwd: resolve(root, '../..') });
  assert.equal(checkout.status, 0, checkout.stderr?.toString());
  assert.deepEqual(checkout.stdout, await readFile(file), `Windows checkout round trip: ${entry.id}`);
}
try {
  await assertOutputSafety();
  const capabilities = run(['--capabilities']);
  assert.equal(capabilities.status, 0, capabilities.stderr);
  const contract = JSON.parse(capabilities.stdout);
  assert.equal(contract.kit.version, 19);
  assert.equal(contract.kit.icons, 74);
  assert.ok(contract.operations.compose);
  assert.ok(contract.kit.recipeFields.includes('main'));

  const recipe = resolve(root, 'examples/create-slide-add.recipe.json');
  const firstSvg = join(temp, 'first.svg');
  const secondSvg = join(temp, 'second.svg');
  const firstDoc = join(temp, 'editable.json');
  for (const [svg, document] of [[firstSvg, firstDoc], [secondSvg, undefined]]) {
    const result = run(['--recipe', recipe, '--profile', 'editable', '--output', svg, ...(document ? ['--document', document] : [])]);
    assert.equal(result.status, 0, result.stderr);
  }
  const a = await readFile(firstSvg);
  const b = await readFile(secondSvg);
  assert.equal(hash(a), hash(b), 'recipe SVG must be deterministic across runs');
  assert.match(a.toString(), /<mask|data-oe-knockout/);
  const editable = JSON.parse(await readFile(firstDoc, 'utf8'));
  assert.equal(editable.elements.length, 1);
  assert.match(editable.elements[0].rawSvgContent, /composer.*metadata|oe-composer/i);

  const commandSvg = await kitSource('slide');
  const actionSvg = await kitSource('add');
  const commandFile = join(temp, 'commands.json');
  const commandDocument = join(temp, 'commands.editor.json');
  const runCommands = async (commands, input) => {
    await writeFile(commandFile, JSON.stringify(commands));
    const args = [...(input ? ['--input', input] : []), '--commands', commandFile, '--document', commandDocument];
    const result = run(args);
    assert.equal(result.status, 0, result.stderr);
    return JSON.parse(await readFile(commandDocument, 'utf8'));
  };
  const assertCommandColors = (document, foreground, accent, label) => {
    const source = document.elements.find((element) => /data-oe-composer/.test(element.rawSvgContent ?? ''))?.rawSvgContent ?? document.elements[0]?.rawSvgContent ?? '';
    assert.match(source, new RegExp(`data-oe-layer="main-icon"[^>]*color="${foreground}"`), `${label} main foreground`);
    assert.match(source, new RegExp(`data-oe-accent-role="CreateGreen" color="${accent}"`), `${label} semantic accent`);
  };
  const recipeRoundTripPath = join(temp, 'recipe-roundtrip.recipe.json');
  const recipeRoundTripDocument = join(temp, 'recipe-roundtrip.editor.json');
  const recipeRoundTrip = JSON.parse(await readFile(recipe, 'utf8'));
  recipeRoundTrip.id = 'recipe-roundtrip-theme-change';
  recipeRoundTrip.params = { ...(recipeRoundTrip.params ?? {}), tokenTheme: 'light' };
  await writeFile(recipeRoundTripPath, JSON.stringify(recipeRoundTrip));
  const recipeCreated = run(['--recipe', recipeRoundTripPath, '--document', recipeRoundTripDocument]);
  assert.equal(recipeCreated.status, 0, recipeCreated.stderr);
  const recipeSavedDocument = JSON.parse(await readFile(recipeRoundTripDocument, 'utf8'));
  const recipeElementId = recipeSavedDocument.elements[0].id;
  assertCommandColors(recipeSavedDocument, '#424242', '#059669', 'recipe save Light');
  const recipeReopenedDark = await runCommands([{ op: 'recompose', args: [recipeElementId, { params: { tokenTheme: 'dark' } }] }], recipeRoundTripDocument);
  assertCommandColors(recipeReopenedDark, '#D6D6D6', '#10B981', 'recipe reopen/recompose Light to Dark');
  const recipeReopenedLight = await runCommands([{ op: 'recompose', args: [recipeElementId, { params: { tokenTheme: 'light' } }] }], commandDocument);
  assertCommandColors(recipeReopenedLight, '#424242', '#059669', 'recipe reopen/recompose Dark to Light');
  const commandDark = await runCommands([{ op: 'compose', args: [{ mainSvg: commandSvg, actionSvg, uid: 'command-dark-default', params: { tokenTheme: 'dark' } }] }]);
  assertCommandColors(commandDark, '#D6D6D6', '#10B981', 'command compose Dark default');
  const commandLight = await runCommands([{ op: 'compose', args: [{ mainSvg: commandSvg, actionSvg, uid: 'command-theme-change', params: { tokenTheme: 'light' } }] }]);
  const commandId = commandLight.elements[0].id;
  const commandDarkRecompose = await runCommands([{ op: 'recompose', args: [commandId, { params: { tokenTheme: 'dark' } }] }], commandDocument);
  assertCommandColors(commandDarkRecompose, '#D6D6D6', '#10B981', 'command recompose Light to Dark');
  const manualDark = await runCommands([{ op: 'compose', args: [{ mainSvg: commandSvg, actionSvg, uid: 'command-manual-ink', params: { tokenTheme: 'dark', baseColor: '#123456' } }] }]);
  assertCommandColors(manualDark, '#123456', '#10B981', 'command compose manual ink');
  const manualId = manualDark.elements[0].id;
  const manualLightRecompose = await runCommands([{ op: 'recompose', args: [manualId, { params: { tokenTheme: 'light' } }] }], commandDocument);
  assertCommandColors(manualLightRecompose, '#123456', '#059669', 'command recompose preserves manual ink');
  const manualDefault = await runCommands([{ op: 'compose', args: [{ mainSvg: commandSvg, actionSvg, uid: 'command-manual-default-valued-ink', params: { tokenTheme: 'light', baseColor: '#424242' } }] }]);
  const manualDefaultId = manualDefault.elements[0].id;
  const manualDefaultDark = await runCommands([{ op: 'recompose', args: [manualDefaultId, { params: { tokenTheme: 'dark' } }] }], commandDocument);
  assertCommandColors(manualDefaultDark, '#424242', '#10B981', 'command recompose preserves explicit manual color equal to Light default');

  const unknown = run(['--recipe', resolve(temp, 'missing.json')]);
  assert.equal(unknown.status, 1);
  assert.match(unknown.stderr, /"ok":false/);
  const rendered = [];
  const exporter = resolveApprovedFilePath(process.env.OE_ICON_EXPORTER, repositoryRoot, 'eng/Export-IconComposerSvg.ps1');
  if (exporter) {
    const productionBase = resolveApprovedOutputBase(process.env.OE_ICON_PRODUCTION_DIR, repositoryRoot, 'artifacts/icon-composer-bundle/production');
    const samplesBase = resolveApprovedOutputBase(process.env.OE_ICON_ARTIFACTS_DIR, repositoryRoot, 'artifacts/icon-composer-bundle/previews');
    const productionDirectory = productionBase ? await createUniqueOutputDirectory(productionBase, repositoryRoot) : undefined;
    const samplesDirectory = samplesBase ? await createUniqueOutputDirectory(samplesBase, repositoryRoot) : undefined;
    const samples = [];
    let productionThemeInvariant;
    for (const theme of ['light', 'dark']) {
      const recipeData = JSON.parse(await readFile(recipe, 'utf8'));
      recipeData.id = `example-slide-add-${theme}-v1`;
      recipeData.params.tokenTheme = theme;
      const themeRecipe = join(temp, `${theme}.recipe.json`);
      await writeFile(themeRecipe, JSON.stringify(recipeData));
      if (samplesDirectory) {
        await writeFile(join(samplesDirectory, `${theme}.recipe.json`), JSON.stringify(recipeData, null, 2) + '\n');
      }
      const editableSvg = join(temp, `${theme}-editable.svg`);
      const editableDocument = join(temp, `${theme}.editor.json`);
      const editableResult = run(['--recipe', themeRecipe, '--profile', 'editable', '--output', editableSvg, '--document', editableDocument]);
      assert.equal(editableResult.status, 0, editableResult.stderr);
      const editableSource = await readFile(editableSvg, 'utf8');
      assert.match(editableSource, /data-oe-accent-role=/);
      const expectedThemeColors = theme === 'light'
        ? { foreground: '#424242', accent: '#059669' }
        : { foreground: '#D6D6D6', accent: '#10B981' };
      assert.match(editableSource, new RegExp(`data-oe-layer="main-icon"[^>]*color="${expectedThemeColors.foreground}"`), `${theme} main foreground token`);
      assert.match(editableSource, new RegExp(`data-oe-accent-role="CreateGreen" color="${expectedThemeColors.accent}"`), `${theme} accent token`);
      if (samplesDirectory) {
        await writeFile(join(samplesDirectory, `${theme}.editable.svg`), editableSource);
        await writeFile(join(samplesDirectory, `${theme}.editor.json`), await readFile(editableDocument));
      }
      const productionSvg = join(temp, `${theme}-production.svg`);
      const productionResult = run(['--recipe', themeRecipe, '--profile', 'openefficiency', '--exporter', exporter, '--output', productionSvg]);
      assert.equal(productionResult.status, 0, productionResult.stderr);
      const productionSource = await readFile(productionSvg, 'utf8');
      assert.match(productionSource, /data-oe-accent-role="CreateGreen"/, `${theme} production SVG retains the semantic accent role`);
      if (productionThemeInvariant === undefined) productionThemeInvariant = hash(productionSource);
      else assert.equal(hash(productionSource), productionThemeInvariant, 'production SVG geometry remains theme-neutral; runtime semantic color resolution is separate from raster preview tokens');
      if (productionDirectory) {
        await writeFile(join(productionDirectory, `${theme}.svg`), productionSource);
      }
      for (const size of [16, 24, 48]) {
        const png = join(temp, `${theme}-${size}.png`);
        const renderedResult = run(['--recipe', themeRecipe, '--profile', 'editable', '--png', png, '--scale', String(size / 24)]);
        assert.equal(renderedResult.status, 0, renderedResult.stderr);
        if (samplesDirectory) await writeFile(join(samplesDirectory, `${theme}-${size}.png`), await readFile(png));
        rendered.push({ theme, size, editableSvgSha256: hash(editableSource), productionSvgSha256: hash(productionSource), previewPngSha256: hash(await readFile(png)), mainForeground: expectedThemeColors.foreground, semanticAccentRole: 'CreateGreen', resolvedAccentColor: expectedThemeColors.accent, previewPath: samplesDirectory ? relativeRepoPath(join(samplesDirectory, `${theme}-${size}.png`)) : undefined });
      }
      samples.push({ theme, recipe: samplesDirectory ? relativeRepoPath(join(samplesDirectory, `${theme}.recipe.json`)) : undefined, editableDocument: samplesDirectory ? relativeRepoPath(join(samplesDirectory, `${theme}.editor.json`)) : undefined, editableSvg: samplesDirectory ? relativeRepoPath(join(samplesDirectory, `${theme}.editable.svg`)) : undefined, productionSvg: productionDirectory ? relativeRepoPath(join(productionDirectory, `${theme}.svg`)) : undefined });
    }
    for (const size of [16, 24, 48]) {
      const light = rendered.find((entry) => entry.theme === 'light' && entry.size === size);
      const dark = rendered.find((entry) => entry.theme === 'dark' && entry.size === size);
      assert.notEqual(light.previewPngSha256, dark.previewPngSha256, `Light/Dark semantic preview differs at ${size}px`);
    }
    const manualRecipeData = JSON.parse(await readFile(recipe, 'utf8'));
    manualRecipeData.id = 'example-slide-add-manual-ink-v1';
    manualRecipeData.params = { ...(manualRecipeData.params ?? {}), tokenTheme: 'dark', baseColor: '#123456' };
    const manualRecipe = join(temp, 'manual-ink.recipe.json');
    const manualSvg = join(temp, 'manual-ink.svg');
    await writeFile(manualRecipe, JSON.stringify(manualRecipeData));
    const manualResult = run(['--recipe', manualRecipe, '--profile', 'editable', '--output', manualSvg]);
    assert.equal(manualResult.status, 0, manualResult.stderr);
    const manualSource = await readFile(manualSvg, 'utf8');
    assert.match(manualSource, /data-oe-layer="main-icon"[^>]*color="#123456"/, 'explicit manual baseColor overrides the theme default');
    assert.match(manualSource, /data-oe-accent-role="CreateGreen" color="#10B981"/, 'manual baseColor retains dark semantic accent resolution');
    const packageJson = await readFile(resolve(root, 'package.json'));
    const packageLock = await readFile(resolve(root, 'package-lock.json'));
    const sourceManifest = await readFile(resolve(root, 'source-manifest.json'));
    const evidencePath = resolveApprovedFilePath(process.env.OE_ICON_EVIDENCE_PATH, repositoryRoot, 'artifacts/icon-composer-bundle/light-dark-size.json');
    if (evidencePath) {
      await ensureSafeDirectory(resolve(evidencePath, '..'), repositoryRoot);
      try {
        const existing = await lstat(evidencePath);
        if (!existing.isFile() || existing.isSymbolicLink()) throw new Error('Evidence output must be a regular, non-symlink file');
      } catch (error) {
        if (error?.code !== 'ENOENT') throw error;
      }
      await writeFile(evidencePath, JSON.stringify({
        schemaVersion: 1,
        status: 'Executed',
        kind: 'icon-composer-bundle-cli-and-light-dark-size-validation',
        startedUtc: startedAt.toISOString(),
        durationMs: Date.now() - started,
        runtime: { node: process.version, os: process.platform, architecture: process.arch },
        source: { manifestSha256: hash(sourceManifest), kitVersion: manifest.version, kitEntries: manifest.entries.length },
        package: { packageJsonSha256: hash(packageJson), packageLockSha256: hash(packageLock) },
        testScriptSha256: hash(await readFile(resolve(root, 'tests/validate.mjs'))),
        exporter: { path: 'eng/Export-IconComposerSvg.ps1', sha256: hash(await readFile(exporter)) },
        outputDirectories: { production: productionDirectory ? relativeRepoPath(productionDirectory) : undefined, previews: samplesDirectory ? relativeRepoPath(samplesDirectory) : undefined },
        retainedSamples: samples,
        checks: { deterministicRecipeRuns: 2, machineReadableCapabilities: 'Passed', errorJson: 'Passed', safeOutputDirectoryRegressions: 'Passed', approvedKitSourceHashes: `${manifest.entries.length} Passed`, windowsCheckoutRoundTrip: `${manifest.entries.length} Passed`, productionWpfExports: 2, semanticColorPreserved: 'Passed', themeForegroundAndAccentAssertions: 'Passed', commandModeThemeDefaultAndRecompose: 'Passed', manualBaseColorOverride: 'Passed', explicitDefaultValuedBaseColorPreserved: 'Passed', themeNeutralProductionSvgGeometry: 'Passed', lightDarkPreviews: rendered.length, sizesPx: [16, 24, 48] },
        rendered,
      }, null, 2) + '\n');
    }
  }
  console.log(JSON.stringify({ ok: true, cases: 9, kitSources: manifest.entries.length, outputSha256: hash(a), rendered }));
} catch (error) {
  console.error(JSON.stringify({ ok: false, error: error.message }));
  process.exitCode = 1;
} finally {
  await rm(temp, { recursive: true, force: true });
}
