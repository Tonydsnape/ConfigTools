# 可热更配置导出工具

将当前目录中的 Excel 配置导出为客户端或服务端使用的 JSON 与 MessagePack 文件。客户端导出还会生成强类型 C#、阶段 Catalog，并原子同步到相邻 Unity 工程。

## 使用

- 双击 `ExportClient.bat` 导出客户端配置并同步 Unity。
- 双击 `ExportServer.bat` 导出服务端配置。
- 双击 `Tools/ValidateConfig.bat` 只检测配置合理性，不生成或修改导出产物。
- 首次执行会依据 `package-lock.json` 自动安装依赖，需要能够访问 npm registry。
- 命令行可使用 `npm run export:client`、`npm run export:server`、`npm run validate` 和 `npm test`。
- Unity 工程默认查找 `../MatchingGoUnityDong`，可用 `MATCHINGGO_UNITY_ROOT` 或 `--unity-root` 覆盖；不会记录开发机绝对路径。

输出文件位于 `Client/json`、`Client/bytes`、`Client/generated`、`Server/json`、`Server/bytes`。导出成功时会整体替换目标端目录；校验失败不会修改已有产物。Unity 侧 JSON 仅供 Editor 读取，YooAsset 只收集 `.bytes`。

## Excel 协议

所有 `.xlsx` 的非空工作表都必须遵循以下协议，Excel 自动生成的 `~$` 临时文件会被忽略。

| 单元格 | 内容 |
| --- | --- |
| A1 / B1 | `Name` / 输出文件名 |
| C1 / D1 | `Type` / `base` 或 `normal` |
| A2 / B2 | `Key` / 正整数嵌套层数 |

`base` 表第 3 行必须为 `KeyName / Target / ValueType / Value`，数据从第 4 行开始，`Key` 必须为 1。`Target` 为 `c/s/sc`，`ValueType` 必须显式填写 `int/long/float/double/string/bool/json`；`json` Value 按 JSON5 对象或数组解析。

`normal` 表第 4 行是导出端、第 5 行是字段名、第 6 行是字段类型，数据从第 7 行开始。前 `Key` 列构成嵌套索引，并且必须标记为 `sc`，类型只能为 `int`、`long` 或 `string`。

导出端支持 `c`（客户端）、`s`（服务端）、`sc`（两端）。字段类型支持 `int`、`long`、`float`、`double`、`string`、`bool`、`json`。`long` 在 JSON 和 MessagePack 中始终表示为十进制字符串。

任意结构、类型或数据错误都会输出工作簿、sheet、行号和单元格位置，并终止整个目标端的导出。独立校验工具会同时验证 `config-delivery.json`，并将结果保存到 `Tools/ValidationReport.txt`；检测过程不会修改 Client 或 Server 产物。

## 投放与代码生成

`config-delivery.json` 是资源投放的唯一事实来源。每张配置必须且只能属于一个阶段：

- `stageId`：小写阶段 ID。
- `tag`：固定为 `config_<stageId>`。
- `includeInBase`：是否同时带 `base`，随标准首包并在启动时更新。
- `loadPolicy`：`startup`、`background` 或 `onDemand`。
- `dataVersion`：阶段数据版本，用于诊断和回滚。
- `configs`：该阶段包含的配置名。

`startup` 必须 `includeInBase=true`。只修改值可以在递增 YooAsset PackageVersion 后热更；新增字段、修改类型或 Key 结构需要发布新的 App 版本。

生成访问示例：

```csharp
int gold = GameConfig.GlobalBaseConfig.StartGold;

ConfigLoadResult result = await GameConfig.ItemNormalConfig.EnsureLoadedAsync();
if (result.IsReady && GameConfig.ItemNormalConfig.TryGet(1, 1001, out var item))
    Debug.Log(item.itemName);
```
