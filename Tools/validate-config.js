#!/usr/bin/env node
'use strict';

const path = require('node:path');
const { ConfigError } = require('../src/errors');
const { validateAndWriteReport } = require('../src/validator');

function parseArguments(argv) {
  let rootDir = path.resolve(__dirname, '..');
  let reportPath = path.join(__dirname, 'ValidationReport.txt');
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === '--input') {
      const value = argv[++index];
      if (!value) throw new ConfigError('--input 缺少目录参数');
      rootDir = path.resolve(value);
    } else if (argument === '--report') {
      const value = argv[++index];
      if (!value) throw new ConfigError('--report 缺少文件参数');
      reportPath = path.resolve(value);
    } else {
      throw new ConfigError(`未知参数 "${argument}"`);
    }
  }
  return { rootDir, reportPath };
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  const result = await validateAndWriteReport(options);
  console.log(result.report);
  console.log(`报告文件：${options.reportPath}`);
  if (!result.valid) process.exitCode = 1;
}

main().catch((error) => {
  console.error(`[校验工具错误] ${error.message}`);
  process.exitCode = 1;
});
