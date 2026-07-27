'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const ExcelJS = require('exceljs');
const { TABLES, exportBuildEventWorkbook, normalizeDocument } = require('../src/build-event-workbook');

function fixture() {
  return {
    schemaVersion: 1,
    chapters: [{ chapterId: 1, assetBundleName: 'scene1', spineRootPath: 'Assets/Scene1/', storyTextKey: 'Story_1', completeTextKey: 'Complete_1', graySpritePath: 'Assets/gray.png', unlockSpritePath: 'Assets/unlock.png', completeSpritePath: 'Assets/complete.png', popupSpritePath: 'Assets/popup.png', finishImagePath: 'Assets/finish.png' }],
    stages: [{ chapterId: 1, stageId: 0, iconPath: 'Assets/icon.png', iconX: 1.25, iconY: -2.5, iconZ: 0, buildCost: 3, textId: 9, finishAudioName: '-' }],
    dependencies: [{ chapterId: 1, stageId: 0, order: 0, requiredStageId: -1 }],
    stageSpines: [{ chapterId: 1, stageId: 0, order: 0, spineName: 'home' }],
    effects: [{ chapterId: 1, stageId: 0, order: 0, x: 0, y: 1, z: 0 }],
    spines: [{ chapterId: 1, spineName: 'home', sortOrder: 0, spineType: 'new', skeletonAssetPath: 'Assets/home.asset', animationName: 'build', idleAnimationName: 'idle', finishAnimationName: '-', overridePrefabPath: '-', eventCheck: true, hideStage: 0 }],
    dialogues: [{ chapterId: 1, triggerType: 'enter', stageId: -1, lineIndex: 0, textKey: 'Hello_\u4e16\u754c', characterId: 0 }],
    audios: [{ chapterId: 1, order: 0, stageId: 0, audioName: 'ambience', invert: false }],
    issues: []
  };
}

test('builds eight normal sheets with declared composite keys', async () => {
  const workbook = require('../src/build-event-workbook').buildWorkbook(fixture());
  assert.equal(workbook.worksheets.length, 8);
  for (const table of TABLES) {
    const sheet = workbook.worksheets.find((item) => item.getCell('B1').value === table.name);
    assert.ok(sheet, table.name);
    assert.equal(sheet.getCell('D1').value, 'normal');
    assert.equal(sheet.getCell('B2').value, table.keyCount);
    assert.deepEqual(sheet.getRow(4).values.slice(1), table.fields.map(([target]) => target));
  }
});

test('preserves float, unicode, booleans and optional sentinel through xlsx', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'build-event-config-'));
  const outputPath = path.join(directory, 'BuildEventConfig.xlsx');
  const reportPath = path.join(directory, 'report.json');
  await exportBuildEventWorkbook(fixture(), { outputPath, reportPath });
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputPath);
  const stage = workbook.worksheets.find((sheet) => sheet.getCell('B1').value === 'BuildEventStageConfig');
  assert.equal(stage.getCell('D7').value, 1.25);
  assert.equal(stage.getCell('I7').value, '-');
  const dialogue = workbook.worksheets.find((sheet) => sheet.getCell('B1').value === 'BuildEventDialogueConfig');
  assert.equal(dialogue.getCell('E7').value, 'Hello_\u4e16\u754c');
});

test('rejects duplicate composite keys before replacing output', () => {
  const document = fixture();
  document.stages.push({ ...document.stages[0] });
  assert.throws(() => normalizeDocument(document), /duplicate key/);
});

test('keeps an existing workbook when validation fails', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'build-event-config-'));
  const outputPath = path.join(directory, 'BuildEventConfig.xlsx');
  const reportPath = path.join(directory, 'report.json');
  const original = Buffer.from('existing-workbook');
  await fs.writeFile(outputPath, original);
  const document = fixture();
  document.stages.push({ ...document.stages[0] });

  await assert.rejects(
    exportBuildEventWorkbook(document, { outputPath, reportPath }),
    /duplicate key/
  );
  assert.deepEqual(await fs.readFile(outputPath), original);
});
