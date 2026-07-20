'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const ExcelJS = require('exceljs');
const JSON5 = require('json5');
const { encode } = require('@msgpack/msgpack');
const { ConfigError, ConfigValidationError } = require('./errors');

const TARGET_FLAGS = new Set(['c', 's', 'sc']);
const FIELD_TYPES = new Set(['int', 'long', 'float', 'double', 'string', 'bool', 'json']);
const KEY_TYPES = new Set(['int', 'long', 'string']);
const INT_MIN = -2147483648;
const INT_MAX = 2147483647;
const LONG_MIN = -(2n ** 63n);
const LONG_MAX = 2n ** 63n - 1n;

function locationOf(filePath, worksheet, cellAddress) {
  const rowMatch = String(cellAddress).match(/(\d+)$/);
  const rowText = rowMatch ? ` 第 ${rowMatch[1]} 行 (${cellAddress})` : ` ${cellAddress}`;
  return `${path.basename(filePath)} [${worksheet.name}]${rowText}`;
}

function capture(errors, operation) {
  try {
    return operation();
  } catch (error) {
    if (error instanceof ConfigValidationError) errors.push(...error.errors);
    else if (error instanceof ConfigError) errors.push(error);
    else throw error;
    return undefined;
  }
}

function throwCollected(errors) {
  if (errors.length > 0) throw new ConfigValidationError(errors);
}

function isBlank(value) {
  return value === null || value === undefined || (typeof value === 'string' && value.trim() === '');
}

function getCellValue(cell, filePath, worksheet) {
  const value = cell.value;
  const location = locationOf(filePath, worksheet, cell.address);

  if (
    cell.type === ExcelJS.ValueType.Formula ||
    (value && typeof value === 'object' && ('formula' in value || 'sharedFormula' in value))
  ) {
    throw new ConfigError('不支持公式单元格，请粘贴为静态值', location);
  }
  if (value instanceof Date || cell.type === ExcelJS.ValueType.Date) {
    throw new ConfigError('不支持日期类型', location);
  }
  if (cell.type === ExcelJS.ValueType.Error) {
    throw new ConfigError('单元格包含 Excel 错误值', location);
  }
  if (value && typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      throw new ConfigError('不支持富文本单元格，请转换为普通文本', location);
    }
    if (typeof value.text === 'string' && typeof value.hyperlink === 'string') {
      throw new ConfigError('不支持超链接单元格，请转换为普通文本', location);
    }
  }
  return value;
}

function requiredString(cell, filePath, worksheet, description) {
  const value = getCellValue(cell, filePath, worksheet);
  const location = locationOf(filePath, worksheet, cell.address);
  if (typeof value !== 'string' || value.trim() === '') {
    throw new ConfigError(`${description}必须是非空字符串`, location);
  }
  return value.trim();
}

function expectText(cell, expected, filePath, worksheet) {
  const actual = requiredString(cell, filePath, worksheet, '协议表头');
  if (actual !== expected) {
    throw new ConfigError(`协议表头应为 "${expected}"，实际为 "${actual}"`, locationOf(filePath, worksheet, cell.address));
  }
}

function parsePositiveInteger(cell, filePath, worksheet, description) {
  const value = getCellValue(cell, filePath, worksheet);
  if (typeof value !== 'number' || !Number.isInteger(value) || value <= 0) {
    throw new ConfigError(`${description}必须是正整数`, locationOf(filePath, worksheet, cell.address));
  }
  return value;
}

function parseTargetFlag(cell, filePath, worksheet) {
  const flag = requiredString(cell, filePath, worksheet, '导出端标记').toLowerCase();
  if (!TARGET_FLAGS.has(flag)) {
    throw new ConfigError('导出端标记只允许 c、s 或 sc', locationOf(filePath, worksheet, cell.address));
  }
  return flag;
}

function isTargetIncluded(flag, target) {
  return flag === 'sc' || (target === 'client' ? flag === 'c' : flag === 's');
}

function validateConfigName(name, filePath, worksheet, cellAddress) {
  const invalidChars = /[<>:"/\\|?*\x00-\x1f]/;
  const reservedName = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i;
  if (
    name === '.' ||
    name === '..' ||
    invalidChars.test(name) ||
    /[. ]$/.test(name) ||
    reservedName.test(name)
  ) {
    throw new ConfigError(`Name "${name}" 不是合法的 Windows 文件名`, locationOf(filePath, worksheet, cellAddress));
  }
}

function worksheetHasData(worksheet) {
  let hasData = false;
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (!isBlank(cell.value)) hasData = true;
    });
  });
  return hasData;
}

function parseInteger(value, location) {
  if (typeof value !== 'number' || !Number.isInteger(value)) {
    throw new ConfigError('int 必须是 Excel 整数单元格', location);
  }
  if (value < INT_MIN || value > INT_MAX) {
    throw new ConfigError(`int 超出 ${INT_MIN} 到 ${INT_MAX} 范围`, location);
  }
  return value;
}

function parseLong(value, location) {
  let parsed;
  if (typeof value === 'number') {
    if (!Number.isSafeInteger(value)) {
      throw new ConfigError('数值型 long 必须在 JavaScript 安全整数范围内，超限值请使用文本单元格', location);
    }
    parsed = BigInt(value);
  } else if (typeof value === 'string' && /^[+-]?\d+$/.test(value.trim())) {
    parsed = BigInt(value.trim());
  } else {
    throw new ConfigError('long 必须是整数或十进制整数字符串', location);
  }
  if (parsed < LONG_MIN || parsed > LONG_MAX) {
    throw new ConfigError('long 超出 64 位有符号整数范围', location);
  }
  return parsed.toString();
}

function parseFloating(value, type, location) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new ConfigError(`${type} 必须是有限的 Excel 数字单元格`, location);
  }
  if (type === 'float' && !Number.isFinite(Math.fround(value))) {
    throw new ConfigError('float 超出 32 位浮点数范围', location);
  }
  return value;
}

function parseJson(value, location) {
  if (typeof value !== 'string') {
    throw new ConfigError('json 字段必须是文本单元格', location);
  }
  let parsed;
  try {
    parsed = JSON5.parse(value);
  } catch (error) {
    throw new ConfigError(`JSON5 解析失败：${error.message}`, location);
  }
  if (parsed === null || typeof parsed !== 'object') {
    throw new ConfigError('json 字段必须是对象或数组', location);
  }
  return parsed;
}

function parseTypedValue(cell, type, filePath, worksheet) {
  const value = getCellValue(cell, filePath, worksheet);
  const location = locationOf(filePath, worksheet, cell.address);
  if (isBlank(value)) throw new ConfigError(`${type} 字段不能为空`, location);

  switch (type) {
    case 'int': return parseInteger(value, location);
    case 'long': return parseLong(value, location);
    case 'float':
    case 'double': return parseFloating(value, type, location);
    case 'string':
      if (typeof value !== 'string') throw new ConfigError('string 必须是文本单元格', location);
      return value;
    case 'bool':
      if (typeof value !== 'boolean') throw new ConfigError('bool 必须是 Excel 布尔单元格', location);
      return value;
    case 'json': return parseJson(value, location);
    default: throw new ConfigError(`未知字段类型 "${type}"`, location);
  }
}

function rowIsBlank(worksheet, rowNumber, startColumn, endColumn, filePath) {
  for (let column = startColumn; column <= endColumn; column += 1) {
    if (!isBlank(worksheet.getCell(rowNumber, column).value)) return false;
  }
  return true;
}

function parseBaseSheet(worksheet, filePath, metadata) {
  const errors = [];
  capture(errors, () => expectText(worksheet.getCell('A3'), 'KeyName', filePath, worksheet));
  capture(errors, () => expectText(worksheet.getCell('B3'), 'Target', filePath, worksheet));
  capture(errors, () => expectText(worksheet.getCell('C3'), 'ValueType', filePath, worksheet));
  capture(errors, () => expectText(worksheet.getCell('D3'), 'Value', filePath, worksheet));
  if (metadata.keyCount !== 1) {
    errors.push(new ConfigError('base 表的 Key 必须为 1', locationOf(filePath, worksheet, 'B2')));
  }

  const entries = [];
  const names = new Set();
  let hasDataRow = false;
  for (let rowNumber = 4; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    if (rowIsBlank(worksheet, rowNumber, 1, 4, filePath)) continue;
    hasDataRow = true;
    const nameCell = worksheet.getCell(rowNumber, 1);
    const rowErrors = [];
    const name = capture(rowErrors, () => requiredString(nameCell, filePath, worksheet, 'KeyName'));
    const target = capture(rowErrors, () => parseTargetFlag(worksheet.getCell(rowNumber, 2), filePath, worksheet));
    const rawType = capture(rowErrors, () => requiredString(worksheet.getCell(rowNumber, 3), filePath, worksheet, 'ValueType'));
    const type = rawType && rawType.toLowerCase();
    if (type !== undefined && !FIELD_TYPES.has(type)) {
      rowErrors.push(new ConfigError(`ValueType 只允许 ${[...FIELD_TYPES].join('/')}`, locationOf(filePath, worksheet, worksheet.getCell(rowNumber, 3).address)));
    }
    const value = type && FIELD_TYPES.has(type)
      ? capture(rowErrors, () => parseTypedValue(worksheet.getCell(rowNumber, 4), type, filePath, worksheet))
      : undefined;
    if (name !== undefined && names.has(name)) {
      rowErrors.push(new ConfigError(`重复 KeyName "${name}"`, locationOf(filePath, worksheet, nameCell.address)));
    }
    if (name !== undefined) names.add(name);
    errors.push(...rowErrors);
    if (rowErrors.length === 0) entries.push({ name, target, type, value });
  }
  if (!hasDataRow) errors.push(new ConfigError('base 表没有数据行', locationOf(filePath, worksheet, 'A4')));
  throwCollected(errors);
  return { ...metadata, entries };
}

function lastDescriptorColumn(worksheet, filePath) {
  let last = 0;
  for (let row = 4; row <= 6; row += 1) {
    for (let column = 1; column <= worksheet.columnCount; column += 1) {
      if (!isBlank(getCellValue(worksheet.getCell(row, column), filePath, worksheet))) last = Math.max(last, column);
    }
  }
  return last;
}

function parseNormalSheet(worksheet, filePath, metadata) {
  const errors = [];
  const fieldCount = lastDescriptorColumn(worksheet, filePath);
  if (fieldCount === 0) throw new ConfigError('normal 表缺少第 4-6 行字段定义', locationOf(filePath, worksheet, 'A4'));
  if (metadata.keyCount > fieldCount) {
    throw new ConfigError(`Key=${metadata.keyCount} 超过字段数量 ${fieldCount}`, locationOf(filePath, worksheet, 'B2'));
  }

  const fields = [];
  const names = new Set();
  for (let column = 1; column <= fieldCount; column += 1) {
    const targetCell = worksheet.getCell(4, column);
    const nameCell = worksheet.getCell(5, column);
    const typeCell = worksheet.getCell(6, column);
    const fieldErrors = [];
    const target = capture(fieldErrors, () => parseTargetFlag(targetCell, filePath, worksheet));
    const name = capture(fieldErrors, () => requiredString(nameCell, filePath, worksheet, '字段名'));
    const rawType = capture(fieldErrors, () => requiredString(typeCell, filePath, worksheet, '字段类型'));
    const type = rawType && rawType.toLowerCase();
    if (type !== undefined && !FIELD_TYPES.has(type)) {
      fieldErrors.push(new ConfigError(`字段类型只允许 ${[...FIELD_TYPES].join('/')}`, locationOf(filePath, worksheet, typeCell.address)));
    }
    if (name !== undefined && names.has(name)) {
      fieldErrors.push(new ConfigError(`重复字段名 "${name}"`, locationOf(filePath, worksheet, nameCell.address)));
    }
    if (name !== undefined) names.add(name);
    if (column <= metadata.keyCount) {
      if (target !== undefined && target !== 'sc') {
        fieldErrors.push(new ConfigError('键列的导出端标记必须为 sc', locationOf(filePath, worksheet, targetCell.address)));
      }
      if (type !== undefined && FIELD_TYPES.has(type) && !KEY_TYPES.has(type)) {
        fieldErrors.push(new ConfigError('键列类型只允许 int、long 或 string', locationOf(filePath, worksheet, typeCell.address)));
      }
    }
    errors.push(...fieldErrors);
    if (fieldErrors.length === 0) fields.push({ column, target, name, type });
  }
  // Invalid descriptors make row types ambiguous, so continue with other sheets instead.
  throwCollected(errors);

  const records = [];
  const keyPaths = new Set();
  let hasDataRow = false;
  for (let rowNumber = 7; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    if (rowIsBlank(worksheet, rowNumber, 1, fieldCount, filePath)) continue;
    hasDataRow = true;
    const rowErrors = [];
    const values = fields.map((field) => capture(rowErrors, () => parseTypedValue(
      worksheet.getCell(rowNumber, field.column), field.type, filePath, worksheet
    )));
    if (rowErrors.length > 0) {
      errors.push(...rowErrors);
      continue;
    }
    const keys = values.slice(0, metadata.keyCount).map(String);
    const signature = JSON.stringify(keys);
    if (keyPaths.has(signature)) {
      errors.push(new ConfigError(
        `重复键路径 ${keys.join(' -> ')}`,
        locationOf(filePath, worksheet, worksheet.getCell(rowNumber, 1).address)
      ));
      continue;
    }
    keyPaths.add(signature);
    records.push({ keys, values });
  }
  if (!hasDataRow) errors.push(new ConfigError('normal 表没有数据行', locationOf(filePath, worksheet, 'A7')));
  throwCollected(errors);
  return { ...metadata, fields, records };
}

function parseWorksheet(worksheet, filePath) {
  expectText(worksheet.getCell('A1'), 'Name', filePath, worksheet);
  expectText(worksheet.getCell('C1'), 'Type', filePath, worksheet);
  expectText(worksheet.getCell('A2'), 'Key', filePath, worksheet);

  const name = requiredString(worksheet.getCell('B1'), filePath, worksheet, 'Name');
  validateConfigName(name, filePath, worksheet, 'B1');
  const type = requiredString(worksheet.getCell('D1'), filePath, worksheet, 'Type').toLowerCase();
  if (type !== 'base' && type !== 'normal') {
    throw new ConfigError('Type 只允许 base 或 normal', locationOf(filePath, worksheet, 'D1'));
  }
  const metadata = {
    name,
    type,
    keyCount: parsePositiveInteger(worksheet.getCell('B2'), filePath, worksheet, 'Key'),
    source: `${path.basename(filePath)} [${worksheet.name}]`
  };
  return type === 'base'
    ? parseBaseSheet(worksheet, filePath, metadata)
    : parseNormalSheet(worksheet, filePath, metadata);
}

async function findWorkbookFiles(rootDir) {
  const entries = await fs.readdir(rootDir, { withFileTypes: true });
  return entries
    .filter((entry) => entry.isFile() && /\.xlsx$/i.test(entry.name) && !entry.name.startsWith('~$'))
    .map((entry) => path.join(rootDir, entry.name))
    .sort((a, b) => path.basename(a).localeCompare(path.basename(b), 'en'));
}

async function loadConfigurations(rootDir) {
  const files = await findWorkbookFiles(rootDir);
  if (files.length === 0) throw new ConfigError(`目录中没有可导出的 .xlsx 文件：${rootDir}`);

  const configurations = [];
  const names = new Map();
  const errors = [];
  for (const filePath of files) {
    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.readFile(filePath);
    } catch (error) {
      errors.push(new ConfigError(`无法读取工作簿：${error.message}`, path.basename(filePath)));
      continue;
    }
    for (const worksheet of workbook.worksheets) {
      if (!worksheetHasData(worksheet)) continue;
      const config = capture(errors, () => parseWorksheet(worksheet, filePath));
      if (!config) continue;
      const normalizedName = config.name.toLocaleLowerCase('en-US');
      if (names.has(normalizedName)) {
        errors.push(new ConfigError(`Name "${config.name}" 与 ${names.get(normalizedName)} 重复`, locationOf(filePath, worksheet, 'B1')));
        continue;
      }
      names.set(normalizedName, config.source);
      configurations.push(config);
    }
  }
  if (configurations.length === 0 && errors.length === 0) {
    errors.push(new ConfigError('所有工作簿均为空，没有可导出的配置'));
  }
  throwCollected(errors);
  return configurations;
}

function defineDataProperty(object, key, value) {
  Object.defineProperty(object, key, {
    value,
    enumerable: true,
    configurable: true,
    writable: true
  });
}

function buildBaseData(config, target) {
  const output = {};
  for (const entry of config.entries) {
    if (isTargetIncluded(entry.target, target)) defineDataProperty(output, entry.name, entry.value);
  }
  return output;
}

function buildNormalData(config, target) {
  const output = {};
  for (const record of config.records) {
    let cursor = output;
    for (let index = 0; index < record.keys.length; index += 1) {
      const key = record.keys[index];
      if (index === record.keys.length - 1) {
        const leaf = {};
        config.fields.forEach((field, fieldIndex) => {
          if (isTargetIncluded(field.target, target)) {
            defineDataProperty(leaf, field.name, record.values[fieldIndex]);
          }
        });
        defineDataProperty(cursor, key, leaf);
      } else {
        if (!Object.hasOwn(cursor, key)) defineDataProperty(cursor, key, {});
        cursor = cursor[key];
      }
    }
  }
  return output;
}

function buildData(config, target) {
  if (target !== 'client' && target !== 'server') throw new ConfigError(`未知导出目标 "${target}"`);
  return config.type === 'base' ? buildBaseData(config, target) : buildNormalData(config, target);
}

async function pathExists(targetPath) {
  try {
    await fs.access(targetPath);
    return true;
  } catch (error) {
    if (error.code === 'ENOENT') return false;
    throw error;
  }
}

async function commitOutput(rootDir, target, artifacts) {
  const folderName = target === 'client' ? 'Client' : 'Server';
  const targetRoot = path.join(rootDir, folderName);
  const token = `${process.pid}-${Date.now()}`;
  const stagingRoot = path.join(rootDir, `.${folderName}.staging-${token}`);
  const backupRoot = path.join(rootDir, `.${folderName}.backup-${token}`);
  const kinds = ['json', 'bytes'];
  const committed = [];
  let swapComplete = false;

  try {
    await Promise.all(kinds.map((kind) => fs.mkdir(path.join(stagingRoot, kind), { recursive: true })));
    for (const artifact of artifacts) {
      await fs.writeFile(path.join(stagingRoot, 'json', `${artifact.name}.json`), artifact.json);
      await fs.writeFile(path.join(stagingRoot, 'bytes', `${artifact.name}.bytes`), artifact.bytes);
    }

    await fs.mkdir(targetRoot, { recursive: true });
    await fs.mkdir(backupRoot, { recursive: true });
    for (const kind of kinds) {
      const current = path.join(targetRoot, kind);
      const backup = path.join(backupRoot, kind);
      const staged = path.join(stagingRoot, kind);
      const hadCurrent = await pathExists(current);
      if (hadCurrent) await fs.rename(current, backup);
      try {
        await fs.rename(staged, current);
      } catch (error) {
        if (hadCurrent && await pathExists(backup)) await fs.rename(backup, current);
        throw error;
      }
      committed.push({ current, backup, hadCurrent });
    }
    swapComplete = true;
  } catch (error) {
    if (!swapComplete) {
      for (const item of committed.reverse()) {
        if (await pathExists(item.current)) await fs.rm(item.current, { recursive: true, force: true });
        if (item.hadCurrent && await pathExists(item.backup)) await fs.rename(item.backup, item.current);
      }
    }
    throw error;
  } finally {
    await fs.rm(stagingRoot, { recursive: true, force: true });
    if (swapComplete || committed.length === 0) {
      await fs.rm(backupRoot, { recursive: true, force: true });
    }
  }
}

function createArtifacts(configurations, target) {
  return configurations.map((config) => {
    const data = buildData(config, target);
    return {
      name: config.name,
      data,
      json: Buffer.from(`${JSON.stringify(data, null, 2)}\n`, 'utf8'),
      bytes: Buffer.from(encode(data))
    };
  });
}

async function exportConfigurations({ rootDir, target }) {
  const configurations = await loadConfigurations(rootDir);
  const artifacts = createArtifacts(configurations, target);
  await commitOutput(rootDir, target, artifacts);
  return artifacts.map(({ name, data }) => ({ name, data }));
}

module.exports = {
  buildData,
  commitOutput,
  createArtifacts,
  exportConfigurations,
  loadConfigurations,
  parseWorksheet
};
