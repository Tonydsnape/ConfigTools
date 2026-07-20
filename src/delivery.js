'use strict';

const crypto = require('node:crypto');
const fs = require('node:fs/promises');
const path = require('node:path');
const { encode } = require('@msgpack/msgpack');
const { ConfigError, ConfigValidationError } = require('./errors');

const LOAD_POLICIES = new Set(['startup', 'background', 'onDemand']);
const ID_PATTERN = /^[a-z][a-z0-9_]*$/;

function stable(value) {
  if (Array.isArray(value)) return value.map(stable);
  if (!value || typeof value !== 'object') return value;
  const result = {};
  for (const key of Object.keys(value).sort()) result[key] = stable(value[key]);
  return result;
}

function sha256(value) {
  const bytes = Buffer.isBuffer(value) ? value : Buffer.from(JSON.stringify(stable(value)), 'utf8');
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

function clientSchema(config) {
  if (config.type === 'base') {
    return {
      name: config.name,
      type: config.type,
      entries: config.entries
        .filter((entry) => entry.target === 'c' || entry.target === 'sc')
        .map((entry) => ({ name: entry.name, type: entry.type }))
    };
  }
  return {
    name: config.name,
    type: config.type,
    keyCount: config.keyCount,
    fields: config.fields
      .filter((field) => field.target === 'c' || field.target === 'sc')
      .map((field) => ({ name: field.name, type: field.type, key: field.column <= config.keyCount }))
  };
}

async function loadDelivery(rootDir, configurations) {
  const deliveryPath = path.join(rootDir, 'config-delivery.json');
  let document;
  try {
    document = JSON.parse(await fs.readFile(deliveryPath, 'utf8'));
  } catch (error) {
    throw new ConfigError(`无法读取 config-delivery.json：${error.message}`);
  }

  const errors = [];
  if (document.schemaVersion !== 1) errors.push(new ConfigError('config-delivery.json schemaVersion 必须为 1'));
  if (!Array.isArray(document.stages) || document.stages.length === 0) {
    errors.push(new ConfigError('config-delivery.json stages 必须是非空数组'));
  }

  const stages = [];
  const stageIds = new Set();
  const tags = new Set();
  const assigned = new Map();
  for (const raw of Array.isArray(document.stages) ? document.stages : []) {
    const stage = { ...raw };
    if (typeof stage.stageId !== 'string' || !ID_PATTERN.test(stage.stageId)) {
      errors.push(new ConfigError(`阶段 stageId "${stage.stageId}" 必须是小写字母开头的字母/数字/下划线`));
    } else if (stageIds.has(stage.stageId)) {
      errors.push(new ConfigError(`重复阶段 stageId "${stage.stageId}"`));
    } else stageIds.add(stage.stageId);

    if (stage.tag !== `config_${stage.stageId}`) {
      errors.push(new ConfigError(`阶段 ${stage.stageId} 的 tag 必须为 config_${stage.stageId}`));
    } else if (tags.has(stage.tag)) {
      errors.push(new ConfigError(`重复配置 tag "${stage.tag}"`));
    } else tags.add(stage.tag);

    if (!LOAD_POLICIES.has(stage.loadPolicy)) {
      errors.push(new ConfigError(`阶段 ${stage.stageId} 的 loadPolicy 只允许 startup/background/onDemand`));
    }
    if (stage.loadPolicy === 'startup' && stage.includeInBase !== true) {
      errors.push(new ConfigError(`startup 阶段 ${stage.stageId} 必须 includeInBase=true`));
    }
    if (typeof stage.includeInBase !== 'boolean') {
      errors.push(new ConfigError(`阶段 ${stage.stageId} 的 includeInBase 必须是 bool`));
    }
    if (!Number.isInteger(stage.dataVersion) || stage.dataVersion <= 0) {
      errors.push(new ConfigError(`阶段 ${stage.stageId} 的 dataVersion 必须是正整数`));
    }
    if (!Array.isArray(stage.configs) || stage.configs.length === 0) {
      errors.push(new ConfigError(`阶段 ${stage.stageId} 必须至少包含一张配置`));
    } else {
      for (const name of stage.configs) {
        if (assigned.has(name)) errors.push(new ConfigError(`配置 ${name} 同时属于 ${assigned.get(name)} 和 ${stage.stageId}`));
        else assigned.set(name, stage.stageId);
      }
    }
    stages.push(stage);
  }

  const names = new Set(configurations.map((config) => config.name));
  for (const name of names) if (!assigned.has(name)) errors.push(new ConfigError(`配置 ${name} 未登记投放阶段`));
  for (const name of assigned.keys()) if (!names.has(name)) errors.push(new ConfigError(`投放清单包含不存在的配置 ${name}`));
  if (errors.length > 0) throw new ConfigValidationError(errors);

  return { schemaVersion: 1, stages };
}

function buildStageArtifacts(delivery, configurations, configArtifacts) {
  const configsByName = new Map(configurations.map((config) => [config.name, config]));
  const artifactsByName = new Map(configArtifacts.map((artifact) => [artifact.name, artifact]));
  return delivery.stages.map((stage) => {
    const catalog = {
      schemaVersion: 1,
      stageId: stage.stageId,
      tag: stage.tag,
      loadPolicy: stage.loadPolicy,
      includeInBase: stage.includeInBase,
      dataVersion: stage.dataVersion,
      configs: stage.configs.map((name) => {
        const config = configsByName.get(name);
        const artifact = artifactsByName.get(name);
        return {
          name,
          type: config.type,
          schemaHash: sha256(clientSchema(config)),
          payloadSha256: sha256(artifact.bytes)
        };
      })
    };
    const name = `ConfigStage_${stage.stageId}`;
    return {
      name,
      data: catalog,
      json: Buffer.from(`${JSON.stringify(catalog, null, 2)}\n`, 'utf8'),
      bytes: Buffer.from(encode(catalog)),
      isCatalog: true
    };
  });
}

module.exports = { buildStageArtifacts, clientSchema, loadDelivery, sha256 };
