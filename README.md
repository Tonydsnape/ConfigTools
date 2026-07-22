# 配置导出工具

将当前目录中的 Excel 配置导出为客户端或服务端使用的 JSON 和 MessagePack 文件。客户端导出还会生成强类型 C#，并原子同步到父目录的 Unity 工程。

## 使用

- 双击 `ExportClient.bat` 导出客户端配置并同步 Unity。
- 双击 `ExportServer.bat` 导出服务端配置。
- 双击 `Tools/ValidateConfig.bat` 只校验 Excel，不修改导出产物。
- 命令行可使用 `npm run export:client`、`npm run export:server`、`npm run validate` 和 `npm test`。
- 配置工程默认位于 Unity 工程的 `Config/`；可以通过 `MATCHINGGO_UNITY_ROOT` 或 `--unity-root` 覆盖 Unity 路径。

客户端产物只有真实配置表：

- `Client/json/<TableName>.json`
- `Client/bytes/<TableName>.bytes`
- `Client/generated/ConfigBindings.g.cs`

同步后的 Unity 目录为：

- `Assets/_GameRes/Config/Editor/json`：仅供 Editor 运行时读取。
- `Assets/_GameRes/Config/Runtime/bytes`：供非 Editor 平台通过 YooAsset 读取。
- `Assets/Scripts/Config/Generated`：强类型访问代码。

导出成功时整体替换目标端产物；校验失败不会覆盖已有有效文件。系统不再生成 `ConfigStage_*`、`config-delivery.json` 或其他阶段清单。

## Excel 协议

所有 `.xlsx` 的非空工作表都必须使用以下元数据：

| 单元格 | 内容 |
| --- | --- |
| A1 / B1 | `Name` / 导出表名 |
| C1 / D1 | `Type` / `base` 或 `normal` |
| A2 / B2 | `Key` / 正整数嵌套 Key 层数 |

`base` 表第 3 行固定为 `KeyName / Target / ValueType / Value`，数据从第 4 行开始，`Key` 必须为 1。

`normal` 表第 4 行是导出端、第 5 行是字段名、第 6 行是字段类型，数据从第 7 行开始。前 `Key` 列组成联合 Key，并且必须导出到客户端和服务端。

导出端支持 `c`、`s`、`sc`。字段类型支持 `int`、`long`、`float`、`double`、`string`、`bool`、`json`；`long` 在 JSON 和 MessagePack 中使用十进制字符串表达。

## Unity 使用

配置系统在游戏启动、YooAsset 清单就绪后统一初始化：

```csharp
ConfigInitializeResult result = await GameConfig.ConfigManager.InitializeAsync();
if (!result.IsReady)
    return;

int gold = GameConfig.GlobalBaseConfig.StartGold;
if (GameConfig.ItemNormalConfig.TryGet(1, 1001, out var item))
    Debug.Log(item.itemName);
```

Editor 固定读取 JSON，Android、iOS 和其他播放器固定读取 MessagePack。所有客户端配置二进制均以独立 YooAsset 地址收集并带 `base` Tag，随基础包内置。任一配置加载或解析失败时不会发布部分快照，启动流程也不会进入业务场景。
