'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const ExcelJS = require('exceljs');

const FIELD_NAMES = ['tableName', 'textKey', 'smart', 'en', 'de', 'es', 'ja', 'ko', 'pt'];
const FIELD_TYPES = ['string', 'string', 'bool', 'string', 'string?', 'string?', 'string?', 'string?', 'string?'];
const TARGETS = ['sc', 'sc', 'c', 'c', 'c', 'c', 'c', 'c', 'c'];

function parseArgs(argv) {
  const result = { overwrite: false };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === '--input') result.input = argv[++index];
    else if (arg === '--output') result.output = argv[++index];
    else if (arg === '--overwrite') result.overwrite = true;
    else throw new Error(`Unknown argument: ${arg}`);
  }
  if (!result.input || !result.output) {
    throw new Error('Usage: node Tools/import-localization.js --input <json> --output <xlsx> [--overwrite]');
  }
  return result;
}

function validateDocument(document) {
  if (!document || document.schemaVersion !== 1 || !Array.isArray(document.rows)) {
    throw new Error('Localization migration document must use schemaVersion 1 and contain rows.');
  }
  const keys = new Set();
  for (const [index, row] of document.rows.entries()) {
    if (!row || typeof row.tableName !== 'string' || row.tableName.trim() === '' ||
        typeof row.textKey !== 'string' || row.textKey.trim() === '') {
      throw new Error(`Row ${index + 1} has an empty tableName or textKey.`);
    }
    if (typeof row.smart !== 'boolean' || typeof row.en !== 'string' || row.en.trim() === '') {
      throw new Error(`Row ${index + 1} must contain smart and a non-empty English translation.`);
    }
    for (const code of ['de', 'es', 'ja', 'ko', 'pt']) {
      if (row[code] !== null && row[code] !== undefined && typeof row[code] !== 'string') {
        throw new Error(`Row ${index + 1} field ${code} must be text or null.`);
      }
    }
    const signature = `${row.tableName}\u0000${row.textKey}`;
    if (keys.has(signature)) throw new Error(`Duplicate localization key: ${row.tableName}/${row.textKey}`);
    keys.add(signature);
  }
}

function createWorkbook(document) {
  validateDocument(document);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'MatchingGo Config Localization Migration';
  workbook.created = new Date(0);
  workbook.modified = new Date(0);
  const sheet = workbook.addWorksheet('Localization');

  sheet.getCell('A1').value = 'Name';
  sheet.getCell('B1').value = 'LocalizationTextConfig';
  sheet.getCell('C1').value = 'Type';
  sheet.getCell('D1').value = 'normal';
  sheet.getCell('A2').value = 'Key';
  sheet.getCell('B2').value = 2;
  sheet.getCell('C2').value = 'Required';
  sheet.getCell('D2').value = false;
  sheet.getRow(4).values = TARGETS;
  sheet.getRow(5).values = FIELD_NAMES;
  sheet.getRow(6).values = FIELD_TYPES;

  const rows = [...document.rows].sort((left, right) =>
    left.tableName.localeCompare(right.tableName, 'en') ||
    left.textKey.localeCompare(right.textKey, 'en'));
  rows.forEach((row, index) => {
    sheet.getRow(index + 7).values = FIELD_NAMES.map((field) => row[field] ?? null);
  });

  sheet.columns = [18, 36, 10, 48, 48, 48, 48, 48, 48].map((width) => ({ width }));
  sheet.views = [{ state: 'frozen', ySplit: 6 }];
  return workbook;
}

async function writeWorkbookAtomic(document, outputPath, overwrite) {
  const resolved = path.resolve(outputPath);
  await fs.mkdir(path.dirname(resolved), { recursive: true });
  try {
    await fs.access(resolved);
    if (!overwrite) throw new Error(`Output already exists: ${resolved}`);
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }

  const token = `${process.pid}-${Date.now()}`;
  const temporary = `${resolved}.tmp-${token}`;
  const backup = `${resolved}.bak-${token}`;
  let movedCurrent = false;
  try {
    await createWorkbook(document).xlsx.writeFile(temporary);
    if (overwrite) {
      try {
        await fs.rename(resolved, backup);
        movedCurrent = true;
      } catch (error) {
        if (error.code !== 'ENOENT') throw error;
      }
    }
    await fs.rename(temporary, resolved);
    if (movedCurrent) await fs.rm(backup, { force: true });
  } catch (error) {
    await fs.rm(temporary, { force: true });
    if (movedCurrent) {
      await fs.rm(resolved, { force: true });
      await fs.rename(backup, resolved);
    }
    throw error;
  }
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const document = JSON.parse(await fs.readFile(path.resolve(options.input), 'utf8'));
  await writeWorkbookAtomic(document, options.output, options.overwrite);
  process.stdout.write(`Generated ${path.resolve(options.output)} with ${document.rows.length} rows.\n`);
}

if (require.main === module) {
  main().catch((error) => {
    process.stderr.write(`${error.stack || error.message}\n`);
    process.exitCode = 1;
  });
}

module.exports = { createWorkbook, validateDocument, writeWorkbookAtomic };
