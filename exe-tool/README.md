# 南方分中心舆情质检辅助工具

这是一个独立桌面程序工程，用来把项目中的 Excel 宏处理逻辑迁移为 Windows exe。

## 已复刻的宏功能

- 读取舆情质检明细 Excel。
- 删除 AG 列为空的数据行。
- 按 H 列筛选指定国网公司。
- 计算处理时长，超过 20 分钟记为超时。
- 按运行当天月份自动选择月度专项质检关键词。
- 月度专项同时检查 C 列和 K 列，任意一列命中即进入专项复核。
- 生成舆情提醒复核。
- 生成无效复核，并随机保留 20%，至少 1 条。
- 添加质检人员、质检结果、补录单号。
- 生成日报送文本。
- 导出最终质检明细 Excel。
- 支持在界面中配置不同源表结构的质检方案，包括可增删字段映射、可增删判断规则、公司名单、无效复核抽样比例和超时阈值。
- 支持在界面中配置 1-12 月专项质检策略和关键词。
- 支持复制日报送内容；运行环境支持时可拖入 Excel 文件。

## 目录说明

- `main.py`：桌面界面入口。
- `processor.py`：Excel 处理逻辑。
- `monthly_special.py`：1-12 月专项策略和关键词。
- `config.json`：公司名单、专项匹配目标列、超时阈值。
- `cli.py`：命令行处理入口，便于内网排错。
- `requirements.txt`：依赖清单。
- `build.bat`：Windows 一键打包脚本。
- `build_on_mac_with_wine.md`：Mac 打包 Windows exe 的现实限制和可选方案。

## 本机调试

在 Mac 或 Windows 上可以先运行：

```bash
python3 main.py
```

也可以命令行处理：

```bash
python3 cli.py 舆情质检明细.xlsx -o 输出目录 -i 质检人员
```

## 配置说明

`config.json` 中保留了默认配置；正常使用时建议直接在程序界面的“方案配置”和“专项配置”里维护：

- `companies`：H 列公司筛选名单。
- `special_target_columns`：月度专项关键词匹配的列号，默认 `[3, 11]`，对应 C 列和 K 列。
- `overtime_threshold_minutes`：超时阈值，默认 20 分钟。
- `preserve_cell_styles`：是否逐格保留源表样式，默认 `false`，大表处理更快。
- `active_scheme_id`：默认运行的质检方案。
- `schemes`：不同表格结构的质检方案。每个方案可配置字段映射行、判断规则行、公司名单、无效复核抽样比例和超时阈值。
- `monthly_special_plans`：界面保存后的 1-12 月专项质检策略和关键词配置。

## Windows 打包

Windows 7/10 内网使用建议：

1. 安装 Python 3.8.x。
2. 进入本目录。
3. 双击 `build.bat`。
4. 使用 `dist\南方分中心舆情质检辅助工具\南方分中心舆情质检辅助工具.exe`。

Windows 7 不建议使用 Python 3.9 及以上版本。
