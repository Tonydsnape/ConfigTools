#!/usr/bin/env node
'use strict';

const path = require('node:path');
const { ConfigError } = require('./errors');
const { exportConfigurations } = require('./exporter');
const { exportClientPipeline } = require('./unity-pipeline');

function parseArguments(argv) {
  let target;
  let rootDir = process.cwd();
  let unityRoot;
  let syncUnity = true;
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === '--target') {
      target = argv[++index];
    } else if (argument === '--input') {
      rootDir = path.resolve(argv[++index] || '');
    } else if (argument === '--unity-root') {
      unityRoot = path.resolve(argv[++index] || '');
    } else if (argument === '--no-unity-sync') {
      syncUnity = false;
    } else {
      throw new ConfigError(`未知参数 "${argument}"`);
    }
  }
  if (target !== 'client' && target !== 'server') {
    throw new ConfigError('必须使用 --target client 或 --target server');
  }
  return { target, rootDir, unityRoot, syncUnity };
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  if (options.target === 'client') {
    const result = await exportClientPipeline({
      rootDir: options.rootDir,
      unityRoot: options.unityRoot,
      sync: options.syncUnity
    });
    console.log(`导出成功：${result.configurations.length} 个配置 -> Client/json, Client/bytes, Client/generated`);
    if (options.syncUnity) console.log(`Unity 同步完成：${result.unityRoot}`);
    for (const config of result.configurations) console.log(`  - ${config.name}`);
    return;
  }
  const results = await exportConfigurations(options);
  const outputFolder = options.target === 'client' ? 'Client' : 'Server';
  console.log(`导出成功：${results.length} 个配置 -> ${outputFolder}/json, ${outputFolder}/bytes`);
  for (const result of results) console.log(`  - ${result.name}`);
}

main().catch((error) => {
  if (error instanceof ConfigError) {
    console.error(`[配置错误] ${error.message}`);
  } else {
    console.error('[导出失败]', error);
  }
  process.exitCode = 1;
});
