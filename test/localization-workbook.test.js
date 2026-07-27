'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const ExcelJS = require('exceljs');
const {
  DEFAULT_COLLECTIONS,
  DEFAULT_LOCALES,
  exportLocalizationWorkbook,
  normalizeDocument
} = require('../src/localization-workbook');
const { resolveUnityRoot } = require('../src/unity-pipeline');

function documentWith(rows, issues = []) {
  return { schemaVersion: 1, collections: DEFAULT_COLLECTIONS, locales: DEFAULT_LOCALES, rows, issues };
}

function row(overrides = {}) {
  return {
    tableName: 'Text',
    key: 'Greeting',
    keyId: '3278632355020800',
    values: {
      en: 'Hello\n<color=#fff>"Player"</color>',
      de: 'Hallo',
      es: 'Hola',
      pt: 'Olá',
      ja: 'こんにちは',
      ko: '안녕하세요'
    },
    smartFlags: { en: true, de: true, es: true, pt: true, ja: true, ko: true },
    ...overrides
  };
}

test('exports the normal workbook schema and preserves localization text', async (context) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'localization-workbook-'));
  context.after(() => fs.rm(directory, { recursive: true, force: true }));
  const outputPath = path.join(directory, 'LocalizationTextConfig.xlsx');
  const reportPath = path.join(directory, 'report.json');

  const report = await exportLocalizationWorkbook(documentWith([row()]), { outputPath, reportPath });
  assert.equal(report.exportedRows, 1);
  assert.equal(report.fallbackCells, 0);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputPath);
  const worksheet = workbook.getWorksheet('LocalizationText');
  assert.deepEqual(worksheet.getRow(1).values.slice(1), ['Name', 'LocalizationTextConfig', 'Type', 'normal']);
  assert.deepEqual(worksheet.getRow(2).values.slice(1), ['Key', 2, 'Required', false]);
  assert.deepEqual(worksheet.getRow(4).values.slice(1), ['sc', 'sc', 'c', 'c', 'c', 'c', 'c', 'c', 'c', 'c']);
  assert.deepEqual(worksheet.getRow(5).values.slice(1), ['tableName', 'key', 'keyId', 'isSmart', 'en', 'de', 'es', 'pt', 'ja', 'ko']);
  assert.deepEqual(worksheet.getRow(6).values.slice(1), ['string', 'string', 'long', 'bool', 'string', 'string', 'string', 'string', 'string', 'string']);
  assert.equal(worksheet.getCell('C7').value, '3278632355020800');
  assert.equal(worksheet.getCell('D7').value, true);
  assert.equal(worksheet.getCell('E7').value, 'Hello\n<color=#fff>"Player"</color>');
  assert.equal(worksheet.getCell('H7').value, 'Olá');
});

test('fills missing translations and reports and marks the fallback cells', async (context) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'localization-fallback-'));
  context.after(() => fs.rm(directory, { recursive: true, force: true }));
  const outputPath = path.join(directory, 'LocalizationTextConfig.xlsx');
  const reportPath = path.join(directory, 'report.json');
  const source = row({
    key: 'Missing',
    values: { en: 'English', de: '', es: null, pt: 'Português', ja: '日本語', ko: '한국어' },
    smartFlags: { en: false, pt: false, ja: false, ko: false }
  });

  const report = await exportLocalizationWorkbook(documentWith([source]), { outputPath, reportPath });
  assert.equal(report.fallbackCells, 2);
  assert.equal(report.issues.filter((issue) => issue.type === 'missingTranslation').length, 2);
  const savedReport = JSON.parse(await fs.readFile(reportPath, 'utf8'));
  assert.equal(savedReport.fallbackCells, 2);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputPath);
  const worksheet = workbook.getWorksheet('LocalizationText');
  assert.equal(worksheet.getCell('F7').value, 'English');
  assert.equal(worksheet.getCell('G7').value, 'English');
  assert.equal(worksheet.getCell('F7').fill.fgColor.argb, 'FFFFE699');
  assert.match(worksheet.getCell('F7').note, /fallback source: en/);
});

test('uses the key when English is missing', () => {
  const normalized = normalizeDocument(documentWith([row({
    key: 'Fallback_Key',
    values: { en: '', de: '', es: '', pt: '', ja: '', ko: '' },
    smartFlags: {}
  })]));
  assert.equal(normalized.rows[0].values.en, 'Fallback_Key');
  assert.equal(normalized.rows[0].values.ko, 'Fallback_Key');
  assert.equal(normalized.rows[0].fallbackLocales.length, 6);
});

test('rejects inconsistent SmartFormat flags', () => {
  assert.throws(() => normalizeDocument(documentWith([row({
    smartFlags: { en: true, de: false, es: true, pt: true, ja: true, ko: true }
  })])), /inconsistent SmartFormat flags/);
});

test('validation failure leaves an existing workbook untouched', async (context) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'localization-atomic-'));
  context.after(() => fs.rm(directory, { recursive: true, force: true }));
  const outputPath = path.join(directory, 'LocalizationTextConfig.xlsx');
  const reportPath = path.join(directory, 'report.json');
  await fs.writeFile(outputPath, 'existing workbook');

  await assert.rejects(
    exportLocalizationWorkbook(documentWith([row(), row()]), { outputPath, reportPath }),
    /duplicate localization key/
  );
  assert.equal(await fs.readFile(outputPath, 'utf8'), 'existing workbook');
});

test('Unity root resolution supports embedded and legacy Config layouts', async (context) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'localization-unity-root-'));
  context.after(() => fs.rm(directory, { recursive: true, force: true }));
  const embeddedConfig = path.join(directory, 'UnityProject', 'Config');
  await fs.mkdir(path.join(directory, 'UnityProject', 'Assets'), { recursive: true });
  await fs.mkdir(path.join(directory, 'UnityProject', 'ProjectSettings'), { recursive: true });
  await fs.mkdir(embeddedConfig, { recursive: true });
  assert.equal(resolveUnityRoot(embeddedConfig), path.join(directory, 'UnityProject'));

  const legacyConfig = path.join(directory, 'Config');
  await fs.mkdir(legacyConfig, { recursive: true });
  assert.equal(resolveUnityRoot(legacyConfig), path.join(directory, 'MatchingGoUnityDong'));
  assert.equal(resolveUnityRoot(legacyConfig, directory), directory);
});
