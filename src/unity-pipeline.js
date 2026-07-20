'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const { generateCSharp } = require('./codegen');
const { buildStageArtifacts, loadDelivery } = require('./delivery');
const { commitOutput, createArtifacts, loadConfigurations } = require('./exporter');
const { ConfigError } = require('./errors');

async function exists(target) {
  try { await fs.access(target); return true; } catch (error) {
    if (error.code === 'ENOENT') return false;
    throw error;
  }
}

async function replaceDirectory(target, files) {
  const token = `${process.pid}-${Date.now()}`;
  const staging = `${target}.staging-${token}`;
  const backup = `${target}.backup-${token}`;
  await fs.mkdir(staging, { recursive: true });
  try {
    for (const file of files) await fs.writeFile(path.join(staging, file.name), file.content);
    const hadTarget = await exists(target);
    if (hadTarget) await fs.rename(target, backup);
    try {
      await fs.rename(staging, target);
    } catch (error) {
      if (hadTarget && await exists(backup)) await fs.rename(backup, target);
      throw error;
    }
    await fs.rm(backup, { recursive: true, force: true });
  } finally {
    await fs.rm(staging, { recursive: true, force: true });
  }
}

async function writeAtomic(target, content) {
  const temporary = `${target}.tmp-${process.pid}-${Date.now()}`;
  await fs.writeFile(temporary, content);
  try {
    await fs.rename(temporary, target);
  } catch (error) {
    if (error.code !== 'EEXIST' && error.code !== 'EPERM') throw error;
    await fs.rm(target, { force: true });
    await fs.rename(temporary, target);
  } finally {
    await fs.rm(temporary, { force: true });
  }
}

async function syncUnityDirectory(target, files, extension) {
  await fs.mkdir(target, { recursive: true });
  const expected = new Set(files.map((file) => file.name));
  for (const entry of await fs.readdir(target, { withFileTypes: true })) {
    if (!entry.isFile() || !entry.name.endsWith(extension) || expected.has(entry.name)) continue;
    await fs.rm(path.join(target, entry.name), { force: true });
    await fs.rm(path.join(target, `${entry.name}.meta`), { force: true });
  }
  for (const file of files) await writeAtomic(path.join(target, file.name), file.content);
}

function resolveUnityRoot(configRoot, explicitRoot) {
  const candidate = explicitRoot || process.env.MATCHINGGO_UNITY_ROOT || path.resolve(configRoot, '..', 'MatchingGoUnityDong');
  return path.resolve(candidate);
}

async function syncUnity(unityRoot, configArtifacts, stageArtifacts, csharp, delivery) {
  if (!await exists(path.join(unityRoot, 'Assets')) || !await exists(path.join(unityRoot, 'ProjectSettings'))) {
    throw new ConfigError(`Unity 工程路径无效：${unityRoot}`);
  }
  const root = path.join(unityRoot, 'Assets', '_GameRes', 'Config');
  const jsonFiles = [...configArtifacts, ...stageArtifacts].map((artifact) => ({
    name: `${artifact.name}.json`, content: artifact.json
  }));
  jsonFiles.push({
    name: 'config-delivery.json',
    content: Buffer.from(`${JSON.stringify(delivery, null, 2)}\n`, 'utf8')
  });
  const byteFiles = [...configArtifacts, ...stageArtifacts].map((artifact) => ({
    name: `${artifact.name}.bytes`, content: artifact.bytes
  }));
  await syncUnityDirectory(path.join(root, 'Editor', 'json'), jsonFiles, '.json');
  await syncUnityDirectory(path.join(root, 'Runtime', 'bytes'), byteFiles, '.bytes');
  await syncUnityDirectory(path.join(unityRoot, 'Assets', 'Scripts', 'Config', 'Generated'), [
    { name: 'ConfigBindings.g.cs', content: Buffer.from(csharp, 'utf8') }
  ], '.cs');
}

async function exportClientPipeline({ rootDir, unityRoot, sync = true }) {
  const configurations = await loadConfigurations(rootDir);
  const delivery = await loadDelivery(rootDir, configurations);
  const configArtifacts = createArtifacts(configurations, 'client');
  const stageArtifacts = buildStageArtifacts(delivery, configurations, configArtifacts);
  await commitOutput(rootDir, 'client', [...configArtifacts, ...stageArtifacts]);
  const csharp = generateCSharp(configurations, delivery);
  await replaceDirectory(path.join(rootDir, 'Client', 'generated'), [
    { name: 'ConfigBindings.g.cs', content: Buffer.from(csharp, 'utf8') }
  ]);
  const resolvedUnityRoot = resolveUnityRoot(rootDir, unityRoot);
  if (sync) await syncUnity(resolvedUnityRoot, configArtifacts, stageArtifacts, csharp, delivery);
  return { configurations, delivery, unityRoot: resolvedUnityRoot };
}

module.exports = { exportClientPipeline, resolveUnityRoot, syncUnity };
