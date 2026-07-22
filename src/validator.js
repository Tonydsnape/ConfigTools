'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const { ConfigError, ConfigValidationError } = require('./errors');
const { loadConfigurations } = require('./exporter');

async function countWorkbookFiles(rootDir) {
  const entries = await fs.readdir(rootDir, { withFileTypes: true });
  return entries.filter((entry) => (
    entry.isFile() && /\.xlsx$/i.test(entry.name) && !entry.name.startsWith('~$')
  )).length;
}

function createReport({ checkedAt, workbookCount, configurations, errors }) {
  const lines = [
    '配置合理性检测报告',
    '==================',
    `检测时间：${checkedAt.toLocaleString('zh-CN', { hour12: false })}`,
    `Excel 文件：${workbookCount}`,
    ''
  ];

  if (errors.length === 0) {
    lines.push('检测结果：通过');
    lines.push(`有效配置：${configurations.length}`);
    lines.push('未发现结构、类型、端标记、重复键或数据内容错误。');
  } else {
    lines.push(`检测结果：失败，共发现 ${errors.length} 个错误`);
    lines.push('');
    errors.forEach((error, index) => lines.push(`${index + 1}. ${error.message}`));
  }
  lines.push('');
  return lines.join('\r\n');
}

async function validateAndWriteReport({ rootDir, reportPath }) {
  const workbookCount = await countWorkbookFiles(rootDir);
  let configurations = [];
  let errors = [];

  try {
    configurations = await loadConfigurations(rootDir);
  } catch (error) {
    if (error instanceof ConfigValidationError) errors = error.errors;
    else if (error instanceof ConfigError) errors = [error];
    else throw error;
  }

  const report = createReport({
    checkedAt: new Date(),
    workbookCount,
    configurations,
    errors
  });
  await fs.mkdir(path.dirname(reportPath), { recursive: true });
  // BOM lets Windows Notepad recognize UTF-8 reports reliably.
  await fs.writeFile(reportPath, `\uFEFF${report}`, 'utf8');
  return { valid: errors.length === 0, configurations, errors, report };
}

module.exports = { createReport, validateAndWriteReport };
