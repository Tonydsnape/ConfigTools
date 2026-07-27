'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const ExcelJS = require('exceljs');

const CONFIG_NAME = 'LocalizationTextConfig';
const SHEET_NAME = 'LocalizationText';
const EXCEL_TEXT_LIMIT = 32767;
const DEFAULT_COLLECTIONS = ['Text', 'AreaName', 'BuildEventText', 'Cards'];
const DEFAULT_LOCALES = ['en', 'de', 'es', 'pt', 'ja', 'ko'];
const FIELDS = [
  { target: 'sc', name: 'tableName', type: 'string' },
  { target: 'sc', name: 'key', type: 'string' },
  { target: 'c', name: 'keyId', type: 'long' },
  { target: 'c', name: 'isSmart', type: 'bool' },
  ...DEFAULT_LOCALES.map((name) => ({ target: 'c', name, type: 'string' }))
];

function isMissing(value) {
  return typeof value !== 'string' || value.trim() === '';
}

function compareOrdinal(left, right) {
  return left < right ? -1 : left > right ? 1 : 0;
}

function assertText(value, name) {
  if (typeof value !== 'string' || value.trim() === '') throw new Error(`${name} must be a non-empty string`);
  if (value.length > EXCEL_TEXT_LIMIT) throw new Error(`${name} exceeds Excel's ${EXCEL_TEXT_LIMIT} character limit`);
}

function normalizeDocument(document) {
  if (!document || document.schemaVersion !== 1) throw new Error('schemaVersion must be 1');
  if (!Array.isArray(document.rows)) throw new Error('rows must be an array');

  const collections = Array.isArray(document.collections) ? document.collections : DEFAULT_COLLECTIONS;
  const locales = Array.isArray(document.locales) ? document.locales : DEFAULT_LOCALES;
  if (JSON.stringify(collections) !== JSON.stringify(DEFAULT_COLLECTIONS)) {
    throw new Error(`collections must be ${DEFAULT_COLLECTIONS.join(',')}`);
  }
  if (JSON.stringify(locales) !== JSON.stringify(DEFAULT_LOCALES)) {
    throw new Error(`locales must be ${DEFAULT_LOCALES.join(',')}`);
  }

  const collectionOrder = new Map(collections.map((name, index) => [name, index]));
  const keys = new Set();
  const issues = Array.isArray(document.issues) ? document.issues.map((issue) => ({ ...issue })) : [];
  const rows = document.rows.map((source, index) => {
    const location = `rows[${index}]`;
    assertText(source.tableName, `${location}.tableName`);
    assertText(source.key, `${location}.key`);
    if (!collectionOrder.has(source.tableName)) throw new Error(`${location}.tableName is not supported: ${source.tableName}`);
    if (typeof source.keyId !== 'string' || !/^\d+$/.test(source.keyId)) {
      throw new Error(`${location}.keyId must be an unsigned decimal string`);
    }

    const signature = `${source.tableName}\u0000${source.key}`;
    if (keys.has(signature)) throw new Error(`duplicate localization key: ${source.tableName}/${source.key}`);
    keys.add(signature);

    const values = source.values && typeof source.values === 'object' ? source.values : {};
    const smartFlags = source.smartFlags && typeof source.smartFlags === 'object' ? source.smartFlags : {};
    const actualSmartFlags = [];
    for (const locale of locales) {
      if (!isMissing(values[locale])) {
        if (typeof smartFlags[locale] !== 'boolean') {
          throw new Error(`${location}.smartFlags.${locale} must be bool when text is present`);
        }
        actualSmartFlags.push(smartFlags[locale]);
      }
    }
    if (new Set(actualSmartFlags).size > 1) {
      throw new Error(`inconsistent SmartFormat flags: ${source.tableName}/${source.key}`);
    }
    const isSmart = actualSmartFlags[0] || false;
    const english = isMissing(values.en) ? source.key : values.en;
    const outputValues = {};
    const fallbackLocales = [];
    for (const locale of locales) {
      let value = values[locale];
      if (isMissing(value)) {
        const fallbackSource = isMissing(values.en) ? 'key' : 'en';
        value = english;
        fallbackLocales.push({ locale, source: fallbackSource });
        issues.push({
          severity: 'warning',
          type: 'missingTranslation',
          tableName: source.tableName,
          key: source.key,
          keyId: source.keyId,
          locale,
          fallbackSource
        });
      }
      assertText(value, `${source.tableName}/${source.key}.${locale}`);
      outputValues[locale] = value;
    }

    return {
      tableName: source.tableName,
      key: source.key,
      keyId: source.keyId,
      isSmart,
      values: outputValues,
      fallbackLocales
    };
  });

  rows.sort((left, right) => {
    const collectionDifference = collectionOrder.get(left.tableName) - collectionOrder.get(right.tableName);
    return collectionDifference || compareOrdinal(left.key, right.key);
  });

  return { collections, locales, rows, issues };
}

function buildWorkbook(normalized) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'MatchingGo Config Exporter';
  workbook.created = new Date(0);
  workbook.modified = new Date(0);
  const worksheet = workbook.addWorksheet(SHEET_NAME, {
    views: [{ state: 'frozen', ySplit: 6, activeCell: 'A7' }]
  });

  worksheet.getRow(1).values = ['Name', CONFIG_NAME, 'Type', 'normal'];
  worksheet.getRow(2).values = ['Key', 2, 'Required', false];
  worksheet.getRow(3).values = ['多语言文本配置表', 'tableName + key 复合索引'];
  worksheet.getRow(4).values = FIELDS.map((field) => field.target);
  worksheet.getRow(5).values = FIELDS.map((field) => field.name);
  worksheet.getRow(6).values = FIELDS.map((field) => field.type);

  const fallbackFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } };
  for (const row of normalized.rows) {
    const excelRow = worksheet.addRow([
      row.tableName,
      row.key,
      row.keyId,
      row.isSmart,
      ...DEFAULT_LOCALES.map((locale) => row.values[locale])
    ]);
    for (const fallback of row.fallbackLocales) {
      const column = 5 + DEFAULT_LOCALES.indexOf(fallback.locale);
      const cell = excelRow.getCell(column);
      cell.fill = fallbackFill;
      cell.note = `Missing ${fallback.locale}; fallback source: ${fallback.source}`;
    }
  }

  worksheet.autoFilter = { from: 'A5', to: 'J5' };
  worksheet.getColumn(1).width = 20;
  worksheet.getColumn(2).width = 36;
  worksheet.getColumn(3).width = 22;
  worksheet.getColumn(4).width = 12;
  for (let column = 5; column <= 10; column += 1) worksheet.getColumn(column).width = 42;
  for (let row = 1; row <= 6; row += 1) worksheet.getRow(row).font = { bold: true };
  worksheet.getRow(5).alignment = { horizontal: 'center' };
  worksheet.getRow(6).alignment = { horizontal: 'center' };
  return workbook;
}

async function replaceFileAtomic(target, content) {
  const directory = path.dirname(target);
  const token = `${process.pid}-${Date.now()}`;
  const temporary = `${target}.tmp-${token}`;
  const backup = `${target}.backup-${token}`;
  await fs.mkdir(directory, { recursive: true });
  await fs.writeFile(temporary, content);
  let hadTarget = false;
  try {
    try {
      await fs.rename(target, backup);
      hadTarget = true;
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
    try {
      await fs.rename(temporary, target);
    } catch (error) {
      if (hadTarget) await fs.rename(backup, target);
      throw error;
    }
    if (hadTarget) await fs.rm(backup, { force: true });
  } finally {
    await fs.rm(temporary, { force: true });
    await fs.rm(backup, { force: true });
  }
}

async function exportLocalizationWorkbook(document, options) {
  if (!options || !options.outputPath || !options.reportPath) {
    throw new Error('outputPath and reportPath are required');
  }
  const normalized = normalizeDocument(document);
  const workbook = buildWorkbook(normalized);
  const workbookBytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const report = {
    schemaVersion: 1,
    generatedAt: new Date().toISOString(),
    collections: normalized.collections,
    locales: normalized.locales,
    exportedRows: normalized.rows.length,
    fallbackCells: normalized.rows.reduce((total, row) => total + row.fallbackLocales.length, 0),
    orphanEntries: normalized.issues.filter((issue) => issue.type === 'orphanEntry').length,
    issues: normalized.issues
  };
  await replaceFileAtomic(path.resolve(options.outputPath), workbookBytes);
  await replaceFileAtomic(
    path.resolve(options.reportPath),
    Buffer.from(`${JSON.stringify(report, null, 2)}\n`, 'utf8')
  );
  return report;
}

module.exports = {
  CONFIG_NAME,
  DEFAULT_COLLECTIONS,
  DEFAULT_LOCALES,
  FIELDS,
  buildWorkbook,
  exportLocalizationWorkbook,
  normalizeDocument,
  replaceFileAtomic
};
