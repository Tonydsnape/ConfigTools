'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const ExcelJS = require('exceljs');
const { decode } = require('@msgpack/msgpack');
const {
  buildData,
  createArtifacts,
  exportConfigurations,
  loadConfigurations
} = require('../src/exporter');
const { generateCSharp } = require('../src/codegen');
const { exportClientPipeline, resolveUnityRoot, syncUnity } = require('../src/unity-pipeline');
const { validateAndWriteReport } = require('../src/validator');
const { writeWorkbookAtomic } = require('../Tools/import-localization');

const ROOT = path.resolve(__dirname, '..');

async function temporaryDirectory(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'config-exporter-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  return directory;
}

function setMetadata(sheet, name, type, keyCount) {
  sheet.getCell('A1').value = 'Name';
  sheet.getCell('B1').value = name;
  sheet.getCell('C1').value = 'Type';
  sheet.getCell('D1').value = type;
  sheet.getCell('A2').value = 'Key';
  sheet.getCell('B2').value = keyCount;
}

function setRequired(sheet, required) {
  sheet.getCell('C2').value = 'Required';
  sheet.getCell('D2').value = required;
}

async function writeSimpleBase(filePath, options = {}) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Base');
  setMetadata(sheet, options.name || 'SimpleBase', 'base', 1);
  sheet.getRow(3).values = ['KeyName', 'Target', 'ValueType', 'Value'];
  sheet.getRow(4).values = ['ConfigKey', 'sc', 'int', 1];
  sheet.getRow(5).values = ['Enabled', 'sc', 'bool', true];
  if (options.formula) sheet.getCell('D5').value = { formula: '1=1', result: true };
  await workbook.xlsx.writeFile(filePath);
}

async function writeTypedNormal(filePath, options = {}) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Typed');
  setMetadata(sheet, options.name || 'TypedConfig', 'normal', 2);
  if (options.required !== undefined) setRequired(sheet, options.required);
  sheet.getRow(4).values = ['sc', 'sc', 'c', 'sc', 'c', 's', 'c'];
  sheet.getRow(5).values = ['itemType', 'itemId', 'itemName', 'serial', 'payload', 'serverOnly', 'optionalText'];
  sheet.getRow(6).values = ['int', 'int', 'string', 'long', 'json', 'string', 'string?'];
  sheet.getRow(7).values = [1, 1001, 'Sword', '9223372036854775806', '{ power: 10 }', 'secret', options.optionalText ?? null];
  if (options.duplicate) {
    sheet.getRow(8).values = [1, 1001, 'Duplicate', '2', '{}', 'secret', null];
  }
  if (options.missingValue) sheet.getCell('B7').value = null;
  await workbook.xlsx.writeFile(filePath);
}

test('current example workbooks generate one descriptor per real config table', async () => {
  const configurations = await loadConfigurations(ROOT);
  assert.deepEqual(configurations.map((config) => config.name), [
    'BuildEventChapterConfig',
    'BuildEventStageConfig',
    'BuildEventStageDependencyConfig',
    'BuildEventStageSpineConfig',
    'BuildEventStageEffectConfig',
    'BuildEventSpineConfig',
    'BuildEventDialogueConfig',
    'BuildEventAudioConfig',
    'GlobalBaseConfig',
    'ItemNormalConfig',
    'LocalizationTextConfig'
  ]);
  const source = generateCSharp(configurations);
  assert.match(source, /public static class GlobalBaseConfig/);
  assert.match(source, /public static int StartGold/);
  assert.match(source, /TryGet\(int itemType, int itemId, out ItemNormalConfigRow row\)/);
  assert.match(source, /TryGet\(string tableName, string key, out LocalizationTextConfigRow row\)/);
  assert.match(source, /new ConfigDescriptor\("LocalizationTextConfig", false,/);
  assert.doesNotMatch(source, /ConfigStageDefinition|public const string StageId|EnsureLoadedAsync|ConfigLoadPolicy/);
});

test('JSON and MessagePack represent the same typed client data', async (t) => {
  const directory = await temporaryDirectory(t);
  await writeTypedNormal(path.join(directory, 'Typed.xlsx'));
  await exportConfigurations({ rootDir: directory, target: 'client' });
  const json = JSON.parse(await fs.readFile(
    path.join(directory, 'Client', 'json', 'TypedConfig.json'), 'utf8'));
  const bytes = await fs.readFile(path.join(directory, 'Client', 'bytes', 'TypedConfig.bytes'));
  assert.deepEqual(decode(bytes), json);
  assert.equal(json['1']['1001'].serial, '9223372036854775806');
  assert.deepEqual(json['1']['1001'].payload, { power: 10 });
  assert.equal(json['1']['1001'].optionalText, null);
  assert.equal(Object.hasOwn(json['1']['1001'], 'serverOnly'), false);
});

test('Required metadata and nullable strings are emitted into generated bindings', async (t) => {
  const directory = await temporaryDirectory(t);
  await writeTypedNormal(path.join(directory, 'Optional.xlsx'), { required: false });
  const configurations = await loadConfigurations(directory);
  assert.equal(configurations[0].required, false);
  const source = generateCSharp(configurations);
  assert.match(source, /ReadOptionalString/);
  assert.match(source, /new ConfigDescriptor\("TypedConfig", false,/);
});

test('localization migration writes the optional normal table contract', async (t) => {
  const directory = await temporaryDirectory(t);
  const output = path.join(directory, 'LocalizationTextConfig.xlsx');
  await writeWorkbookAtomic({
    schemaVersion: 1,
    rows: [{
      tableName: 'Text', textKey: 'Level', smart: false,
      en: 'Level {0}', de: null, es: 'Nivel {0}', ja: null, ko: null, pt: null
    }]
  }, output, false);

  const configurations = await loadConfigurations(directory);
  const config = configurations[0];
  assert.equal(config.name, 'LocalizationTextConfig');
  assert.equal(config.required, false);
  assert.deepEqual(config.records[0].keys, ['Text', 'Level']);
  assert.equal(config.records[0].values[4], null);
});

test('duplicate keys and partial rows are rejected', async (t) => {
  const duplicateDirectory = await temporaryDirectory(t);
  await writeTypedNormal(path.join(duplicateDirectory, 'Duplicate.xlsx'), { duplicate: true });
  await assert.rejects(loadConfigurations(duplicateDirectory));

  const partialDirectory = await temporaryDirectory(t);
  await writeTypedNormal(path.join(partialDirectory, 'Partial.xlsx'), { missingValue: true });
  await assert.rejects(loadConfigurations(partialDirectory));
});

test('failed validation preserves the previous exported files', async (t) => {
  const directory = await temporaryDirectory(t);
  await writeSimpleBase(path.join(directory, 'Good.xlsx'));
  await exportConfigurations({ rootDir: directory, target: 'client' });
  const outputPath = path.join(directory, 'Client', 'json', 'SimpleBase.json');
  const previous = await fs.readFile(outputPath);

  await writeSimpleBase(path.join(directory, 'Formula.xlsx'), {
    name: 'FormulaBase',
    formula: true
  });
  await assert.rejects(exportConfigurations({ rootDir: directory, target: 'client' }));
  assert.deepEqual(await fs.readFile(outputPath), previous);
});

test('client pipeline emits only table JSON, bytes, and generated C#', async (t) => {
  const directory = await temporaryDirectory(t);
  await writeSimpleBase(path.join(directory, 'OnlyTable.xlsx'));
  await exportClientPipeline({ rootDir: directory, sync: false });
  assert.deepEqual(await fs.readdir(path.join(directory, 'Client', 'json')), ['SimpleBase.json']);
  assert.deepEqual(await fs.readdir(path.join(directory, 'Client', 'bytes')), ['SimpleBase.bytes']);
  assert.deepEqual(await fs.readdir(path.join(directory, 'Client', 'generated')), ['ConfigBindings.g.cs']);
  assert.equal(resolveUnityRoot(directory), path.resolve(directory, '..', 'MatchingGoUnityDong'));

  const embeddedConfigRoot = path.join(directory, 'UnityProject', 'Config');
  await fs.mkdir(path.join(directory, 'UnityProject', 'Assets'), { recursive: true });
  await fs.mkdir(path.join(directory, 'UnityProject', 'ProjectSettings'), { recursive: true });
  assert.equal(resolveUnityRoot(embeddedConfigRoot), path.join(directory, 'UnityProject'));
});

test('Unity sync removes obsolete stage and delivery artifacts', async (t) => {
  const directory = await temporaryDirectory(t);
  const unityRoot = path.join(directory, 'UnityProject');
  const jsonRoot = path.join(unityRoot, 'Assets', '_GameRes', 'Config', 'Editor', 'json');
  const bytesRoot = path.join(unityRoot, 'Assets', '_GameRes', 'Config', 'Runtime', 'bytes');
  await fs.mkdir(jsonRoot, { recursive: true });
  await fs.mkdir(bytesRoot, { recursive: true });
  await fs.mkdir(path.join(unityRoot, 'ProjectSettings'), { recursive: true });
  await fs.writeFile(path.join(jsonRoot, 'ConfigStage_startup.json'), '{}');
  await fs.writeFile(path.join(jsonRoot, 'ConfigStage_startup.json.meta'), 'stale');
  await fs.writeFile(path.join(jsonRoot, 'config-delivery.json'), '{}');
  await fs.writeFile(path.join(bytesRoot, 'ConfigStage_startup.bytes'), 'stale');
  await fs.writeFile(path.join(bytesRoot, 'ConfigStage_startup.bytes.meta'), 'stale');

  const configurations = [{
    name: 'SimpleBase',
    type: 'base',
    keyCount: 1,
    entries: [{ name: 'Value', target: 'c', type: 'int', value: 7 }]
  }];
  await syncUnity(
    unityRoot,
    createArtifacts(configurations, 'client'),
    generateCSharp(configurations));

  assert.deepEqual(await fs.readdir(jsonRoot), ['SimpleBase.json']);
  assert.deepEqual(await fs.readdir(bytesRoot), ['SimpleBase.bytes']);
});

test('C# generator escapes reserved field names without stage metadata', () => {
  const configurations = [{
    name: 'KeywordConfig',
    type: 'base',
    keyCount: 1,
    entries: [{ name: 'class', target: 'c', type: 'int', value: 1 }]
  }];
  const source = generateCSharp(configurations);
  assert.match(source, /public int @class/);
  assert.doesNotMatch(source, /StageId/);
});

test('validation writes a UTF-8 BOM report without requiring a delivery manifest', async (t) => {
  const directory = await temporaryDirectory(t);
  await writeTypedNormal(path.join(directory, 'Invalid.xlsx'), { missingValue: true });
  const reportPath = path.join(directory, 'Tools', 'ValidationReport.txt');
  const result = await validateAndWriteReport({ rootDir: directory, reportPath });
  assert.equal(result.valid, false);
  assert.ok(result.errors.length > 0);
  const savedReport = await fs.readFile(reportPath, 'utf8');
  assert.equal(savedReport.codePointAt(0), 0xfeff);
});

test('buildData filters endpoint-specific values', () => {
  const config = {
    name: 'EndpointBase',
    type: 'base',
    entries: [
      { name: 'Shared', target: 'sc', type: 'int', value: 1 },
      { name: 'Client', target: 'c', type: 'int', value: 2 },
      { name: 'Server', target: 's', type: 'int', value: 3 }
    ]
  };
  assert.deepEqual(buildData(config, 'client'), { Shared: 1, Client: 2 });
  assert.deepEqual(buildData(config, 'server'), { Shared: 1, Server: 3 });
});
