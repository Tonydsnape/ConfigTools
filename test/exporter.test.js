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
  exportConfigurations,
  loadConfigurations
} = require('../src/exporter');
const { validateAndWriteReport } = require('../src/validator');
const { loadDelivery } = require('../src/delivery');
const { generateCSharp } = require('../src/codegen');

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

async function writeSimpleBase(filePath, options = {}) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(options.sheetName || 'Base');
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
  setMetadata(sheet, options.name || 'TypedConfig', 'normal', 1);
  sheet.getRow(4).values = [options.keyTarget || 'sc', 'sc', 'sc', 'c', 's', 'sc', 'sc'];
  sheet.getRow(5).values = ['id', 'count', 'ratio', 'clientOnly', 'serverOnly', 'enabled', 'payload'];
  sheet.getRow(6).values = ['long', 'int', 'double', 'string', 'string', 'bool', 'json'];
  sheet.getRow(7).values = [
    '9223372036854775807',
    42,
    1.25,
    'client',
    'server',
    true,
    "{nested:{value:2},list:[1,2]}"
  ];
  if (options.duplicate) sheet.getRow(8).values = sheet.getRow(7).values;
  if (options.partial) sheet.getCell('B7').value = null;
  if (options.multipleErrors) {
    sheet.getCell('B7').value = null;
    sheet.getRow(8).values = [2, 'not-an-int', 2.5, 'client-2', 'server-2', false, '{broken'];
  }
  await workbook.xlsx.writeFile(filePath);
}

test('current example workbooks follow target flags and nested key rules', async () => {
  const configurations = await loadConfigurations(ROOT);
  assert.deepEqual(configurations.map((config) => config.name), ['GlobalBaseConfig', 'ItemNormalConfig']);

  const base = configurations.find((config) => config.name === 'GlobalBaseConfig');
  const normal = configurations.find((config) => config.name === 'ItemNormalConfig');
  const clientBase = buildData(base, 'client');
  const serverBase = buildData(base, 'server');
  const clientNormal = buildData(normal, 'client');
  const serverNormal = buildData(normal, 'server');

  assert.deepEqual(Object.keys(clientBase), ['ConfigKey', 'StartGold2', 'MaxLevel', 'StartGold', 'BattleTimeLimit']);
  assert.deepEqual(Object.keys(serverBase), ['ConfigKey', 'MaxLevel', 'BattleTimeLimit']);
  assert.deepEqual(clientBase.MaxLevel, { key: 1, value: 1 });
  assert.equal(clientNormal['1']['1001'].itemName, '铁剑');
  assert.equal(clientNormal['1']['1001'].stackable, false);
  assert.equal(clientNormal['1']['1001'].dropWeight, undefined);
  assert.equal(clientNormal['1']['1001'].itemDesc, '新手基础近战武器');
  assert.equal(serverNormal['1']['1001'].itemDesc, undefined);
  assert.equal(serverNormal['1']['1001'].dropWeight, 500);
});

test('exports long, JSON5 and endpoint fields to equivalent JSON and MessagePack', async (t) => {
  const directory = await temporaryDirectory(t);
  await writeTypedNormal(path.join(directory, 'Typed.xlsx'));

  await exportConfigurations({ rootDir: directory, target: 'client' });
  await exportConfigurations({ rootDir: directory, target: 'server' });

  for (const folder of ['Client', 'Server']) {
    const json = JSON.parse(await fs.readFile(path.join(directory, folder, 'json', 'TypedConfig.json'), 'utf8'));
    const binary = decode(await fs.readFile(path.join(directory, folder, 'bytes', 'TypedConfig.bytes')));
    assert.deepEqual(binary, json);
    assert.equal(json['9223372036854775807'].id, '9223372036854775807');
    assert.deepEqual(json['9223372036854775807'].payload, { nested: { value: 2 }, list: [1, 2] });
  }

  const client = JSON.parse(await fs.readFile(path.join(directory, 'Client/json/TypedConfig.json'), 'utf8'));
  const server = JSON.parse(await fs.readFile(path.join(directory, 'Server/json/TypedConfig.json'), 'utf8'));
  assert.equal(client['9223372036854775807'].clientOnly, 'client');
  assert.equal(client['9223372036854775807'].serverOnly, undefined);
  assert.equal(server['9223372036854775807'].clientOnly, undefined);
  assert.equal(server['9223372036854775807'].serverOnly, 'server');
});

test('rejects key columns that are not exported to both targets', async (t) => {
  const directory = await temporaryDirectory(t);
  await writeTypedNormal(path.join(directory, 'Invalid.xlsx'), { keyTarget: 'c' });
  await assert.rejects(loadConfigurations(directory), /Invalid\.xlsx \[Typed\] 第 4 行 \(A4\): 键列的导出端标记必须为 sc/);
});

test('rejects duplicate key paths and partially empty rows with cell locations', async (t) => {
  const duplicateDirectory = await temporaryDirectory(t);
  await writeTypedNormal(path.join(duplicateDirectory, 'Duplicate.xlsx'), { duplicate: true });
  await assert.rejects(loadConfigurations(duplicateDirectory), /Duplicate\.xlsx \[Typed\] 第 8 行 \(A8\): 重复键路径/);

  const partialDirectory = await temporaryDirectory(t);
  await writeTypedNormal(path.join(partialDirectory, 'Partial.xlsx'), { partial: true });
  await assert.rejects(loadConfigurations(partialDirectory), /Partial\.xlsx \[Typed\] 第 7 行 \(B7\): int 字段不能为空/);
});

test('rejects formulas and preserves prior output when validation fails', async (t) => {
  const directory = await temporaryDirectory(t);
  await writeSimpleBase(path.join(directory, 'Good.xlsx'));
  await exportConfigurations({ rootDir: directory, target: 'client' });
  const outputPath = path.join(directory, 'Client/json/SimpleBase.json');
  const previous = await fs.readFile(outputPath);

  await writeSimpleBase(path.join(directory, 'Formula.xlsx'), { name: 'FormulaBase', formula: true });
  await assert.rejects(
    exportConfigurations({ rootDir: directory, target: 'client' }),
    /Formula\.xlsx \[Base\] 第 5 行 \(D5\): 不支持公式单元格/
  );
  assert.deepEqual(await fs.readFile(outputPath), previous);
});

test('delivery catalog covers every config and generated C# exposes typed accessors', async () => {
  const configurations = await loadConfigurations(ROOT);
  const delivery = await loadDelivery(ROOT, configurations);
  const source = generateCSharp(configurations, delivery);
  assert.match(source, /public static class GlobalBaseConfig/);
  assert.match(source, /public static int StartGold/);
  assert.match(source, /TryGet\(int itemType, int itemId, out ItemNormalConfigRow row\)/);
  assert.deepEqual(delivery.stages.map((stage) => stage.tag), ['config_startup', 'config_item']);
});

test('delivery rejects unassigned configs', async (t) => {
  const directory = await temporaryDirectory(t);
  await fs.writeFile(path.join(directory, 'config-delivery.json'), JSON.stringify({
    schemaVersion: 1,
    stages: [{
      stageId: 'startup', tag: 'config_startup', includeInBase: true,
      loadPolicy: 'startup', dataVersion: 1, configs: ['AssignedConfig']
    }]
  }));
  await assert.rejects(
    loadDelivery(directory, [{ name: 'AssignedConfig' }, { name: 'MissingConfig' }]),
    /配置 MissingConfig 未登记投放阶段/
  );
});

test('C# generator escapes reserved field names', () => {
  const configurations = [{
    name: 'KeywordConfig', type: 'base', keyCount: 1, entries: [
      { name: 'class', target: 'c', type: 'int', value: 1 }
    ]
  }];
  const delivery = { stages: [{
    stageId: 'startup', tag: 'config_startup', includeInBase: true,
    loadPolicy: 'startup', dataVersion: 1, configs: ['KeywordConfig']
  }] };
  assert.match(generateCSharp(configurations, delivery), /public int @class/);
});

test('validation tool aggregates row errors and writes a UTF-8 report', async (t) => {
  const directory = await temporaryDirectory(t);
  await writeTypedNormal(path.join(directory, 'Multiple.xlsx'), { multipleErrors: true });
  await fs.writeFile(path.join(directory, 'config-delivery.json'), JSON.stringify({
    schemaVersion: 1,
    stages: [{
      stageId: 'typed', tag: 'config_typed', includeInBase: false,
      loadPolicy: 'onDemand', dataVersion: 1, configs: ['TypedConfig']
    }]
  }));
  const reportPath = path.join(directory, 'Tools', 'ValidationReport.txt');

  const result = await validateAndWriteReport({ rootDir: directory, reportPath });
  assert.equal(result.valid, false);
  assert.equal(result.errors.length, 3);
  assert.match(result.report, /第 7 行 \(B7\): int 字段不能为空/);
  assert.match(result.report, /第 8 行 \(B8\): int 必须是 Excel 整数单元格/);
  assert.match(result.report, /第 8 行 \(G8\): JSON5 解析失败/);
  const savedReport = await fs.readFile(reportPath, 'utf8');
  assert.equal(savedReport.codePointAt(0), 0xfeff);
  assert.match(savedReport, /检测结果：失败，共发现 3 个错误/);
});
