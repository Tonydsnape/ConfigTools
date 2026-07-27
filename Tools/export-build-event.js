#!/usr/bin/env node
'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const { exportBuildEventWorkbook } = require('../src/build-event-workbook');

function parseArguments(argv) {
  const options = {};
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === '--input') options.inputPath = path.resolve(argv[++index] || '');
    else if (argument === '--output') options.outputPath = path.resolve(argv[++index] || '');
    else if (argument === '--report') options.reportPath = path.resolve(argv[++index] || '');
    else throw new Error(`Unknown argument: ${argument}`);
  }
  if (!options.inputPath || !options.outputPath || !options.reportPath) {
    throw new Error('Usage: export-build-event --input <json> --output <xlsx> --report <json>');
  }
  return options;
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  const document = JSON.parse(await fs.readFile(options.inputPath, 'utf8'));
  const report = await exportBuildEventWorkbook(document, options);
  console.log(`BuildEvent workbook exported: ${JSON.stringify(report.counts)}`);
}

main().catch((error) => {
  console.error(`[BuildEvent export failed] ${error.message}`);
  process.exitCode = 1;
});
