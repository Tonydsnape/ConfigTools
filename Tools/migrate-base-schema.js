#!/usr/bin/env node
'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const ExcelJS = require('exceljs');

function inferType(value) {
  if (typeof value === 'boolean') return 'bool';
  if (typeof value === 'number') return Number.isInteger(value) ? 'int' : 'double';
  if (typeof value === 'string') {
    const text = value.trim();
    return text.startsWith('{') || text.startsWith('[') ? 'json' : 'string';
  }
  throw new Error(`无法迁移 Value 类型：${typeof value}`);
}

async function main() {
  const root = path.resolve(__dirname, '..');
  const names = (await fs.readdir(root)).filter((name) => /\.xlsx$/i.test(name) && !name.startsWith('~$'));
  for (const name of names) {
    const filePath = path.join(root, name);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    let changed = false;
    for (const sheet of workbook.worksheets) {
      if (String(sheet.getCell('D1').value).toLowerCase() !== 'base') continue;
      if (sheet.getCell('B3').value === 'Target' && sheet.getCell('C3').value === 'ValueType') continue;
      if (sheet.getCell('A3').value !== 'KeyName' || sheet.getCell('B3').value !== 'Type' || sheet.getCell('C3').value !== 'Value') {
        throw new Error(`${name} [${sheet.name}] 不是可识别的旧 base 协议`);
      }
      sheet.getRow(3).values = ['KeyName', 'Target', 'ValueType', 'Value'];
      for (let row = 4; row <= sheet.rowCount; row += 1) {
        const value = sheet.getCell(row, 3).value;
        if (value === null || value === undefined || value === '') continue;
        sheet.getCell(row, 4).value = value;
        sheet.getCell(row, 3).value = inferType(value);
      }
      changed = true;
    }
    if (changed) {
      await workbook.xlsx.writeFile(filePath);
      console.log(`已迁移：${name}`);
    }
  }
}

main().catch((error) => {
  console.error(error.message || error);
  process.exitCode = 1;
});
