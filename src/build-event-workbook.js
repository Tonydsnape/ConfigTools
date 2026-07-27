'use strict';

const path = require('node:path');
const ExcelJS = require('exceljs');
const { replaceFileAtomic } = require('./localization-workbook');

const OPTIONAL = '-';

const TABLES = [
  {
    property: 'chapters', name: 'BuildEventChapterConfig', keyCount: 1,
    fields: [
      ['sc', 'chapterId', 'int'], ['c', 'assetBundleName', 'string'], ['c', 'spineRootPath', 'string'],
      ['c', 'storyTextKey', 'string'], ['c', 'completeTextKey', 'string'],
      ['c', 'graySpritePath', 'string'], ['c', 'unlockSpritePath', 'string'],
      ['c', 'completeSpritePath', 'string'], ['c', 'popupSpritePath', 'string'],
      ['c', 'finishImagePath', 'string']
    ]
  },
  {
    property: 'stages', name: 'BuildEventStageConfig', keyCount: 2,
    fields: [
      ['sc', 'chapterId', 'int'], ['sc', 'stageId', 'int'], ['c', 'iconPath', 'string'],
      ['c', 'iconX', 'float'], ['c', 'iconY', 'float'], ['c', 'iconZ', 'float'],
      ['c', 'buildCost', 'int'], ['c', 'textId', 'int'], ['c', 'finishAudioName', 'string']
    ]
  },
  {
    property: 'dependencies', name: 'BuildEventStageDependencyConfig', keyCount: 3,
    fields: [['sc', 'chapterId', 'int'], ['sc', 'stageId', 'int'], ['sc', 'order', 'int'], ['c', 'requiredStageId', 'int']]
  },
  {
    property: 'stageSpines', name: 'BuildEventStageSpineConfig', keyCount: 3,
    fields: [['sc', 'chapterId', 'int'], ['sc', 'stageId', 'int'], ['sc', 'order', 'int'], ['c', 'spineName', 'string']]
  },
  {
    property: 'effects', name: 'BuildEventStageEffectConfig', keyCount: 3,
    fields: [
      ['sc', 'chapterId', 'int'], ['sc', 'stageId', 'int'], ['sc', 'order', 'int'],
      ['c', 'x', 'float'], ['c', 'y', 'float'], ['c', 'z', 'float']
    ]
  },
  {
    property: 'spines', name: 'BuildEventSpineConfig', keyCount: 2,
    fields: [
      ['sc', 'chapterId', 'int'], ['sc', 'spineName', 'string'], ['c', 'sortOrder', 'int'],
      ['c', 'spineType', 'string'], ['c', 'skeletonAssetPath', 'string'],
      ['c', 'animationName', 'string'], ['c', 'idleAnimationName', 'string'],
      ['c', 'finishAnimationName', 'string'], ['c', 'overridePrefabPath', 'string'],
      ['c', 'eventCheck', 'bool'], ['c', 'hideStage', 'int']
    ]
  },
  {
    property: 'dialogues', name: 'BuildEventDialogueConfig', keyCount: 4,
    fields: [
      ['sc', 'chapterId', 'int'], ['sc', 'triggerType', 'string'], ['sc', 'stageId', 'int'],
      ['sc', 'lineIndex', 'int'], ['c', 'textKey', 'string'], ['c', 'characterId', 'int']
    ]
  },
  {
    property: 'audios', name: 'BuildEventAudioConfig', keyCount: 2,
    fields: [
      ['sc', 'chapterId', 'int'], ['sc', 'order', 'int'], ['c', 'stageId', 'int'],
      ['c', 'audioName', 'string'], ['c', 'invert', 'bool']
    ]
  }
];

function assertRecord(record, table, index) {
  if (!record || typeof record !== 'object' || Array.isArray(record)) {
    throw new Error(`${table.property}[${index}] must be an object`);
  }
  for (const [, name, type] of table.fields) {
    const value = record[name];
    if (type === 'string' && (typeof value !== 'string' || value.length === 0)) {
      throw new Error(`${table.property}[${index}].${name} must be a non-empty string`);
    }
    if ((type === 'int' || type === 'float') && (typeof value !== 'number' || !Number.isFinite(value))) {
      throw new Error(`${table.property}[${index}].${name} must be a finite number`);
    }
    if (type === 'int' && !Number.isInteger(value)) {
      throw new Error(`${table.property}[${index}].${name} must be an integer`);
    }
    if (type === 'bool' && typeof value !== 'boolean') {
      throw new Error(`${table.property}[${index}].${name} must be a boolean`);
    }
  }
}

function normalizeDocument(document) {
  if (!document || document.schemaVersion !== 1) throw new Error('schemaVersion must be 1');
  const normalized = { schemaVersion: 1, issues: Array.isArray(document.issues) ? document.issues : [] };
  for (const table of TABLES) {
    if (!Array.isArray(document[table.property]) || document[table.property].length === 0) {
      throw new Error(`${table.property} must be a non-empty array`);
    }
    const keys = new Set();
    normalized[table.property] = document[table.property].map((record, index) => {
      assertRecord(record, table, index);
      const key = JSON.stringify(table.fields.slice(0, table.keyCount).map(([, name]) => record[name]));
      if (keys.has(key)) throw new Error(`${table.name} contains duplicate key ${key}`);
      keys.add(key);
      return { ...record };
    });
  }
  return normalized;
}

function buildWorkbook(document) {
  const normalized = normalizeDocument(document);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'MatchingGo BuildEvent Config Exporter';
  workbook.created = new Date(0);
  workbook.modified = new Date(0);

  for (const table of TABLES) {
    const sheet = workbook.addWorksheet(table.name.replace('BuildEvent', '').replace('Config', '') || table.name);
    sheet.getRow(1).values = ['Name', table.name, 'Type', 'normal'];
    sheet.getRow(2).values = ['Key', table.keyCount];
    sheet.getRow(3).values = ['BuildEvent configuration', `${table.keyCount}-column composite key`];
    sheet.getRow(4).values = table.fields.map(([target]) => target);
    sheet.getRow(5).values = table.fields.map(([, name]) => name);
    sheet.getRow(6).values = table.fields.map(([, , type]) => type);
    for (const record of normalized[table.property]) {
      sheet.addRow(table.fields.map(([, name]) => record[name]));
    }
    sheet.views = [{ state: 'frozen', ySplit: 6 }];
    sheet.autoFilter = { from: { row: 6, column: 1 }, to: { row: 6, column: table.fields.length } };
    for (let column = 1; column <= table.fields.length; column += 1) {
      const type = table.fields[column - 1][2];
      sheet.getColumn(column).width = type === 'string' ? 34 : 14;
    }
  }
  return workbook;
}

async function exportBuildEventWorkbook(document, options) {
  if (!options || !options.outputPath || !options.reportPath) {
    throw new Error('outputPath and reportPath are required');
  }
  const normalized = normalizeDocument(document);
  const workbook = buildWorkbook(normalized);
  const workbookBytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const counts = Object.fromEntries(TABLES.map((table) => [table.property, normalized[table.property].length]));
  const report = {
    schemaVersion: 1,
    generatedAt: new Date().toISOString(),
    counts,
    issues: normalized.issues
  };
  await replaceFileAtomic(path.resolve(options.outputPath), workbookBytes);
  await replaceFileAtomic(
    path.resolve(options.reportPath),
    Buffer.from(`${JSON.stringify(report, null, 2)}\n`, 'utf8')
  );
  return report;
}

module.exports = { OPTIONAL, TABLES, buildWorkbook, exportBuildEventWorkbook, normalizeDocument };
