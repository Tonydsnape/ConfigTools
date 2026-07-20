'use strict';

const { clientSchema, sha256 } = require('./delivery');

const CS_TYPES = {
  int: 'int',
  long: 'long',
  float: 'float',
  double: 'double',
  string: 'string',
  bool: 'bool',
  json: 'JToken'
};

const CS_READERS = {
  int: 'ReadInt32',
  long: 'ReadInt64',
  float: 'ReadSingle',
  double: 'ReadDouble',
  string: 'ReadString',
  bool: 'ReadBoolean',
  json: 'ReadJson'
};

const RESERVED = new Set(('abstract as base bool break byte case catch char checked class const continue decimal default delegate do double else enum event explicit extern false finally fixed float for foreach goto if implicit in int interface internal is lock long namespace new null object operator out override params private protected public readonly ref return sbyte sealed short sizeof stackalloc static string struct switch this throw true try typeof uint ulong unchecked unsafe ushort using virtual void volatile while').split(' '));

function identifier(value, context) {
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(value)) throw new Error(`${context} "${value}" 不是合法 C# 标识符`);
  return RESERVED.has(value) ? `@${value}` : value;
}

function q(value) {
  return JSON.stringify(String(value));
}

function clientFields(config) {
  return config.type === 'base'
    ? config.entries.filter((entry) => entry.target === 'c' || entry.target === 'sc')
    : config.fields.filter((field) => field.target === 'c' || field.target === 'sc');
}

function readExpression(type, objectName, jsonName, configName) {
  return `ConfigValue.${CS_READERS[type]}(${objectName}, ${q(jsonName)}, ${q(configName)})`;
}

function propertyDeclaration(field) {
  const name = identifier(field.name, '字段名');
  if (field.type === 'json')
    return `        private readonly JToken _${name};\n        public JToken ${name} => _${name}.DeepClone();`;
  return `        public ${CS_TYPES[field.type]} ${name} { get; }`;
}

function assignment(field) {
  const name = identifier(field.name, '字段名');
  return field.type === 'json'
    ? `            this._${name} = ${name}.DeepClone();`
    : `            this.${name} = ${name};`;
}

function generateBase(config, stage) {
  const className = identifier(config.name, '配置名');
  const fields = clientFields(config);
  const ctorParams = fields.map((field) => `${CS_TYPES[field.type]} ${identifier(field.name, '字段名')}`).join(', ');
  const assignments = fields.map(assignment).join('\n');
  const properties = fields.map(propertyDeclaration).join('\n');
  const staticProps = fields.map((field) => `        public static ${CS_TYPES[field.type]} ${identifier(field.name, '字段名')} => RequireSnapshot().${identifier(field.name, '字段名')};`).join('\n');
  const reads = fields.map((field) => `                ${readExpression(field.type, 'obj', field.name, config.name)}`).join(',\n');

  return `
    public sealed class ${className}Snapshot
    {
${properties}

        internal ${className}Snapshot(${ctorParams})
        {
${assignments}
        }
    }

    public static class ${className}
    {
        private static ${className}Snapshot _snapshot;
        public const string StageId = ${q(stage.stageId)};
        public static bool IsLoaded => _snapshot != null;
${staticProps}

        public static UniTask<ConfigLoadResult> EnsureLoadedAsync(ConfigLoadIntent intent = ConfigLoadIntent.Foreground) =>
            ConfigManager.EnsureStageAsync(StageId, intent);

        internal static object Parse(JToken root)
        {
            JObject obj = ConfigValue.RequireObject(root, ${q(config.name)});
            return new ${className}Snapshot(
${reads});
        }

        internal static void Publish(object snapshot) => _snapshot = (${className}Snapshot)snapshot;
        internal static void Clear() => _snapshot = null;

        private static ${className}Snapshot RequireSnapshot() =>
            _snapshot ?? throw new ConfigNotLoadedException(${q(config.name)}, StageId);
    }
`;
}

function generateNormal(config, stage) {
  const className = identifier(config.name, '配置名');
  const rowName = `${className}Row`;
  const fields = clientFields(config);
  const keyFields = config.fields.slice(0, config.keyCount);
  const ctorParams = fields.map((field) => `${CS_TYPES[field.type]} ${identifier(field.name, '字段名')}`).join(', ');
  const assignments = fields.map(assignment).join('\n');
  const properties = fields.map(propertyDeclaration).join('\n');
  const methodParams = keyFields.map((field) => `${CS_TYPES[field.type]} ${identifier(field.name, 'Key 字段')}`).join(', ');
  const methodArgs = keyFields.map((field) => identifier(field.name, 'Key 字段')).join(', ');
  const rowReads = fields.map((field) => `                    ${readExpression(field.type, 'row', field.name, config.name)}`).join(',\n');
  const keyReads = keyFields.map((field, index) => `ConfigValue.${CS_READERS[field.type]}(keys[${index}], ${q(config.name + '.' + field.name)})`).join(', ');

  return `
    public sealed class ${rowName}
    {
${properties}

        internal ${rowName}(${ctorParams})
        {
${assignments}
        }
    }

    internal sealed class ${className}Snapshot
    {
        internal readonly Dictionary<string, ${rowName}> Rows;
        internal ${className}Snapshot(Dictionary<string, ${rowName}> rows) => Rows = rows;
    }

    public static class ${className}
    {
        private static ${className}Snapshot _snapshot;
        public const string StageId = ${q(stage.stageId)};
        public static bool IsLoaded => _snapshot != null;
        public static int Count => RequireSnapshot().Rows.Count;
        public static IReadOnlyCollection<${rowName}> Values => RequireSnapshot().Rows.Values;

        public static UniTask<ConfigLoadResult> EnsureLoadedAsync(ConfigLoadIntent intent = ConfigLoadIntent.Foreground) =>
            ConfigManager.EnsureStageAsync(StageId, intent);

        public static bool TryGet(${methodParams}, out ${rowName} row)
        {
            row = null;
            return _snapshot != null && _snapshot.Rows.TryGetValue(ConfigKey.Compose(${methodArgs}), out row);
        }

        public static ${rowName} Get(${methodParams})
        {
            if (_snapshot == null) throw new ConfigNotLoadedException(${q(config.name)}, StageId);
            if (_snapshot.Rows.TryGetValue(ConfigKey.Compose(${methodArgs}), out ${rowName} row)) return row;
            throw new KeyNotFoundException($"配置 ${config.name} 不存在 Key: {ConfigKey.Display(${methodArgs})}");
        }

        internal static object Parse(JToken root)
        {
            var rows = new Dictionary<string, ${rowName}>(StringComparer.Ordinal);
            foreach (ConfigRowNode node in ConfigValue.EnumerateRows(root, ${config.keyCount}, ${q(config.name)}))
            {
                JToken[] keys = node.Keys;
                JObject row = node.Row;
                string key = ConfigKey.Compose(${keyReads});
                var value = new ${rowName}(
${rowReads});
                if (!rows.TryAdd(key, value)) throw new ConfigDataException($"配置 ${config.name} 包含重复 Key: {key}");
            }
            return new ${className}Snapshot(rows);
        }

        internal static void Publish(object snapshot) => _snapshot = (${className}Snapshot)snapshot;
        internal static void Clear() => _snapshot = null;

        private static ${className}Snapshot RequireSnapshot() =>
            _snapshot ?? throw new ConfigNotLoadedException(${q(config.name)}, StageId);
    }
`;
}

function generateCSharp(configurations, delivery) {
  const byName = new Map(configurations.map((config) => [config.name, config]));
  const stageByConfig = new Map();
  for (const stage of delivery.stages) for (const name of stage.configs) stageByConfig.set(name, stage);

  const bodies = configurations.map((config) => config.type === 'base'
    ? generateBase(config, stageByConfig.get(config.name))
    : generateNormal(config, stageByConfig.get(config.name))).join('\n');

  const descriptors = configurations.map((config) => {
    const stage = stageByConfig.get(config.name);
    const className = identifier(config.name, '配置名');
    return `            new ConfigDescriptor(${q(config.name)}, ${q(stage.stageId)}, ${q(sha256(clientSchema(config)))}, ${q(config.name)}, ${className}.Parse, ${className}.Publish, ${className}.Clear)`;
  }).join(',\n');

  const stages = delivery.stages.map((stage) => `            new ConfigStageDefinition(${q(stage.stageId)}, ${q(stage.tag)}, ConfigLoadPolicy.${stage.loadPolicy[0].toUpperCase() + stage.loadPolicy.slice(1)}, ${stage.includeInBase ? 'true' : 'false'}, ${stage.dataVersion}, ${q(`ConfigStage_${stage.stageId}`)}, new[] { ${stage.configs.map(q).join(', ')} })`).join(',\n');

  return `// <auto-generated />
using System;
using System.Collections.Generic;
using Cysharp.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace GameConfig
{
${bodies}
    internal static class GeneratedConfigCatalog
    {
        internal static readonly ConfigDescriptor[] Configs =
        {
${descriptors}
        };

        internal static readonly ConfigStageDefinition[] Stages =
        {
${stages}
        };
    }
}
`;
}

module.exports = { generateCSharp };
