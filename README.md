<!-- markdownlint-disable MD031 MD033 MD036 MD041 -->

<div align="center">

<a href="https://v2.nonebot.dev/store">
  <img src="https://raw.githubusercontent.com/A-kirami/nonebot-plugin-template/resources/nbp_logo.png" width="180" height="180" alt="NoneBotPluginLogo">
</a>

<p>
  <img src="https://raw.githubusercontent.com/lgc-NB2Dev/readme/main/template/plugin.svg" alt="NoneBotPluginText">
</p>

# NoneBot-Plugin-Xlsx

_✨ 基于 NoneBot2 的 Excel 文件导入导出和数据管理插件！ ✨_

[![python3](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](https://opensource.org/licenses/MIT)

</div>

## 📖 介绍

这是一个功能强大的 Excel 数据管理插件，专为游戏代肝记录、签到管理等场景设计。支持 Excel 文件的导入导出、数据统计、进度跟踪等功能。

### 🌟 主要特性

- **📁 Excel 文件管理**: 支持导入现有 Excel 文件到数据库
- **📊 数据导出**: 支持单个游戏导出、批量导出、合并导出
- **🎮 多游戏支持**: 自动识别并注册多个游戏的命令
- **📝 灵活记录**: 支持 `+1` 传统格式和自定义次数格式
- **🎯 进度跟踪**: 自动统计进度，达到目标次数自动祝贺
- **🎨 美化样式**: A列用户名黄色背景，完成记录蓝色背景
- **⚙️ 高度可配置**: 支持自定义完成次数、列宽、文件路径等
- **📤 文件上传**: 支持导出后直接上传到聊天

## 💿 安装

以下提到的方法 任选**其一**即可

<details open>
<summary>[推荐] 直接部署</summary>

将项目文件放置到 NoneBot2 项目目录下即可使用

```bash
# 克隆或下载项目
git clone <repository-url>
cd Xlsx

# 安装依赖
pip install -r requirements.txt
# 或使用 poetry
poetry install
```

</details>

<details>
<summary>手动安装依赖</summary>

```bash
pip install openpyxl
pip install httpx
```

打开 nonebot2 项目根目录下的 `pyproject.toml` 文件, 在 `[tool.nonebot]` 部分配置插件路径：

```toml
[tool.nonebot]
plugin_dirs = ["./plugins"]
```

</details>

## ⚙️ 配置

### 环境变量配置

在 `.env` 文件中配置以下变量（可选）：

```ini
# ===== 基础路径配置 =====
EXCEL_FOLDER="./xlsx"              # Excel文件目录路径

# ===== 调试配置 =====  
DEBUG_MODE=false                   # 是否启用调试模式

# ===== Excel格式配置 =====
ROW_HEIGHT=50.0                    # 行高设置（单位：磅）
NAME_COLUMN_WIDTH=20.0             # A列列宽设置（单位：字符数）

# ===== 文件导入配置 =====
FILE_SELECTION_TIMEOUT=30         # 文件选择超时时间（秒）
PRIVATE_FILE_TIMEOUT=30            # 私聊文件等待超时时间（秒）

# ===== 查询配置 =====
DEFAULT_LOOKUP_COUNT=3             # 默认查询显示的最新记录数
```

### 配置说明

- **EXCEL_FOLDER**: Excel文件存储目录，默认为项目根目录下的 `xlsx` 文件夹
- **DEBUG_MODE**: 开启后会显示详细的调试信息
- **ROW_HEIGHT**: Excel表格行高，默认50磅
- **NAME_COLUMN_WIDTH**: A列（用户名列）宽度，默认20字符
- **COMPLETION_COUNT**: 完成一个周期所需次数，可根据需要调整（如10、30、50、100等）
- **DEFAULT_LOOKUP_COUNT**: 查询命令默认显示的最新记录数，默认3条

## 🎉 使用

### 📥 导入 Excel 文件

首先需要将现有的 Excel 文件导入到数据库：

```
/文档导入                      # 查看可导入的文件列表
/文档导入 原神.xlsx            # 导入指定Excel文件
```

### 🎮 游戏记录命令

导入成功后，会自动注册对应的游戏命令：

```
# 传统格式（向后兼容）
/原神 玩家名 +1                # 为玩家添加1次记录
/绝区零 小明 +1               # 为小明添加1次记录

# 新格式（自定义次数）  
/原神 玩家名 5                # 为玩家添加5次记录
/崩铁 小红 20                # 为小红添加20次记录
/绝区零 带空格的用户 10        # 支持带空格的用户名
```

### 📤 导出功能

```
# 导出单个游戏
/文档导出 原神                # 导出原神数据到Excel文件
/文档导出 原神 --upload       # 导出原神数据并显示文件信息

# 批量导出
/文档导出 all                 # 将所有游戏合并导出到一个Excel文件的不同sheet
/文档导出 all --upload        # 合并导出并显示文件信息
```

### 🎯 游戏管理

```
/创建表格 新游戏名称           # 手动创建新游戏并注册命令
```

### 📊 查询功能

```
/表格查询 原神 张三             # 查询张三在原神中的最新3条记录
/表格查询 原神 张三 5           # 查询张三在原神中的最新5条记录
/表格查询 崩铁 李四 10          # 查询李四在崩铁中的最新10条记录
```

### 📚 获取帮助

```
/xlsx帮助                      # 显示详细的命令帮助信息
```

### 💡 使用场景示例

<details>
<summary>📝 日常签到记录</summary>

```
用户: /原神 张三 +1
机器人: ✅ 已更新 张三 的记录: 06-15_1

用户: /原神 张三 +1  
机器人: ✅ 已更新 张三 的记录: 06-15_2
```

</details>

<details>
<summary>🔄 批量补录记录</summary>

```
用户: /绝区零 李四 10
机器人: ✅ 已为 李四 添加 10 条记录
        记录: 06-15_1, 06-15_2, 06-15_3, 06-15_4, 06-15_5, 06-15_6, 06-15_7, 06-15_8, 06-15_9, 06-15_10
        当前进度: 10/30
```

</details>

<details>
<summary>🎯 完成目标庆祝</summary>

```
用户: /崩铁 王五 5
机器人: ✅ 已为 王五 添加 5 条记录
        记录: 06-15_26, 06-15_27, 06-15_28, 06-15_29, 06-15_30
        🎉 恭喜完成30次！
```

</details>

<details>
<summary>📊 查询用户记录</summary>

```
用户: /表格查询 原神 张三
机器人: 📊 查询结果
        🎮 游戏: 原神
        👤 用户: 张三
        📈 当前进度: 25/30
        📝 总记录数: 25

        🕒 最新 3 条记录:
        1. 2025-06-14 - 第23次
        2. 2025-06-15 - 第24次
        3. 2025-06-16 - 第25次
```

</details>

### ⚠️ 使用限制

- **权限要求**: 所有命令都需要 SUPERUSER 权限
- **自定义次数范围**: 1-100
- **查询记录数量范围**: 1-20
- **达到设定的完成次数后自动开始新周期**
- **支持的文件格式**: `.xlsx`、`.xls`

## 📋 命令列表

### 🎮 动态游戏指令
- **`/<游戏名> <用户名> +1`** - 为指定用户在指定游戏中添加1次记录
- **`/<游戏名> <用户名> <次数>`** - 为指定用户在指定游戏中添加指定次数的记录

### 📁 文件管理指令
- **`/文档导入 [文件名]`** - 导入Excel文件到数据库
  - 不带参数：列出可用的Excel文件
  - 带文件名：导入指定的Excel文件
- **`/文档导出 <游戏名|all> [--upload]`** - 导出数据到Excel文件
  - `<游戏名>`：导出指定游戏的数据
  - `all`：导出所有游戏的数据到一个文件
  - `--upload`：导出后显示文件信息

### 🎯 游戏管理指令
- **`/创建表格 <游戏名>`** - 手动创建新游戏
  - 在数据库中创建新的游戏记录
  - 自动注册对应的动态指令

### 📊 查询指令
- **`/表格查询 <游戏名> <用户名> [记录数量]`** - 查询用户记录
  - 查询指定用户在指定游戏中的记录
  - 默认显示最新3条记录
  - 可指定显示的记录数量（1-20）

### 📚 帮助指令
- **`/xlsx帮助`** - 显示详细的命令帮助信息
  - 包含所有可用命令的说明
  - 使用限制和注意事项
  - 实用提示和技巧

## � 技术特性

### 📊 数据管理

- **SQLite 数据库**: 使用 SQLite 作为数据存储，轻量且可靠
- **多周期支持**: 自动管理用户的多个周期记录
- **数据完整性**: 自动处理数据导入时的格式兼容性

### 🎨 Excel 样式

- **A列黄色背景**: 用户名列统一使用黄色背景，便于识别
- **完成记录蓝色背景**: 已完成周期的记录使用蓝色背景标识
- **无表头设计**: 数据从第1行开始，没有表头行
- **时间戳文件名**: 导出文件自动添加时间戳避免重名

### 📁 文件组织

```
项目结构：
xlsx/                          # Excel文件存储目录
├── 原神.xlsx                  # 原始Excel文件
├── 绝区零.xlsx
├── 崩铁.xlsx
└── exports/                   # 导出文件目录
    ├── 原神_export_06-15-1430.xlsx
    ├── 绝区零_export_06-15-1431.xlsx
    └── all_games_export_06-15-1432.xlsx

plugins/xlsx/                  # 插件源码目录
├── __init__.py
├── __main__.py               # 主逻辑和命令处理
├── config.py                 # 配置管理
├── database.py               # 数据库操作
├── excel_importer.py         # Excel导入功能
└── excel_exporter.py         # Excel导出功能

records.db                    # SQLite数据库文件
```

## 📞 联系与支持

### 🐛 问题反馈

如果遇到问题或有建议，请通过以下方式联系：

- **GitHub Issues**: 在项目仓库提交问题
- **QQ群**: [待补充]
- **邮箱**: [待补充]

### 🤝 贡献指南

欢迎提交 Pull Request 来改进这个项目！

1. Fork 本仓库
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 提交 Pull Request

## 💡 鸣谢

- **NoneBot2**: 感谢 NoneBot2 框架提供的强大功能
- **OpenPyXL**: 感谢 OpenPyXL 库提供的 Excel 操作支持
- **OneBot**: 感谢 OneBot 协议的标准化

## 📝 更新日志

### v0.2.1 (最新)

- 🗑️ **移除base64上传**: 文件上传不支持base64
- 📚 **新增帮助命令**: 添加 `/xlsx帮助` 命令，提供详细的使用说明
- 🎨 **优化用户体验**: 文件导出后显示清晰的文件信息和路径
- 📝 **文档更新**: 更新README文档，反映最新功能变化

### v0.2.0

- ✨ **命令中文化**: 将英文命令改为中文命令，更符合中文用户习惯
  - `/xlsximport` → `/文档导入`
  - `/xlsxexport` → `/文档导出`
  - `/xlsxcreate` → `/创建表格`
  - `/xlsxlookup` → `/表格查询`
- 📊 **新增查询功能**: 支持查询用户在指定游戏中的记录
- ⚙️ **配置优化**: 新增查询相关配置项
- 🗑️ **代码精简**: 移除群文件导入功能，简化代码结构
- 🐛 **Bug修复**: 修复了展开收起文本的显示问题

### v0.1.0

- ✨ 初始版本发布
- 📁 支持 Excel 文件导入导出
- 🎮 支持多游戏动态命令注册
- 📝 支持传统 `+1` 和自定义次数格式
- 🎨 实现 A列黄色背景和完成记录蓝色背景
- ⚙️ 支持可配置的完成次数
- 📤 支持文件上传功能
- 🔄 支持合并导出到单个 Excel 文件的多个 sheet

### 计划功能

- 📊 数据统计和图表功能
- 🔔 完成提醒和通知功能  
- 👥 用户排行榜功能
- 📅 定时任务和自动备份
- 🌐 Web 管理界面
- 🎯 自定义完成目标和奖励系统

---

<div align="center">

**如果这个项目对你有帮助，请给它一个 ⭐ Star！**

</div>
