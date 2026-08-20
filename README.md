# 网易云歌单导出工具

将网易云音乐歌单数据导出为格式精美的 Excel 文件,支持自动美化样式、斑马纹行、冻结表头、自动列宽。

## 功能特性

- 解析网易云歌单 JSON 数据(兼容被 HTML 标签包裹的 JSON)
- 导出 Excel 包含两个 Sheet:
  - **歌单歌曲明细**:序号、歌名、歌手、专辑、时长、播放量、发行时间、歌曲ID
  - **歌单信息汇总**:歌单名称、创建者、歌曲总数、导出时间
- 表头蓝底白字加粗,数据行微软雅黑字体
- 交替行斑马纹背景,提升可读性
- 自动调整列宽(中英文混排宽度计算)
- 冻结首行表头,滚动时表头常驻

## 环境依赖

- Python 3.7+
- 依赖库:`openpyxl`(标准库 `json`、`re`、`unicodedata`、`datetime` 无需安装)

安装依赖:

```bash
pip install openpyxl
```

## 使用方法

### 1. 准备歌单 JSON 数据

在浏览器中访问网易云音乐歌单 API(需登录态 Cookie),例如:

```
https://music.163.com/api/v6/playlist/detail?id=<歌单ID>&n=1000
```

将返回的 JSON 内容保存为 `playlist.json`,放在脚本同目录下。

> 说明:本脚本兼容被 HTML 标签包裹的 JSON(如 `<html><body>{...}</body></html>`),会自动提取 `<body>` 内的 JSON 内容。

### 2. 运行脚本

```bash
python wy_playlist_to_excel.py
```

### 3. 查看结果

运行完成后,在同目录下生成 `网易云歌单导出.xlsx`。

## 配置说明

在 [wy_playlist_to_excel.py](file:///f:/wyyyy/wy_playlist_to_excel.py) 顶部可修改以下配置:

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| `JSON_FILE_PATH` | 输入歌单 JSON 文件路径 | `playlist.json` |
| `OUTPUT_EXCEL_PATH` | 输出 Excel 文件路径 | `网易云歌单导出.xlsx` |
| `HEADERS` | Excel 表头字段 | 序号/歌名/歌手/专辑/时长/播放量/发行时间/歌曲ID |

样式配置(表头颜色、字体、斑马纹颜色、边框等)也可在同文件的「样式配置」区域自定义。

## 字段映射

脚本从网易云 API 返回的 JSON 中提取以下字段:

| 输出字段 | JSON 路径 |
|---------|-----------|
| 序号 | 自动生成 |
| 歌名 | `playlist.tracks[].name` |
| 歌手 | `playlist.tracks[].ar[].name`(多歌手用 `/` 分隔) |
| 专辑 | `playlist.tracks[].al.name` |
| 时长 | `playlist.tracks[].dt`(毫秒转 `分:秒`) |
| 播放量 | `playlist.tracks[].playCount` |
| 发行时间 | `playlist.tracks[].publishTime`(毫秒时间戳转 `YYYY-MM-DD`) |
| 歌曲ID | `playlist.tracks[].id` |

## 目录结构

```
wyyyy/
├── wy_playlist_to_excel.py   # 主脚本
├── playlist.json             # 输入歌单数据(需自行准备)
├── 网易云歌单导出.xlsx         # 运行后生成的输出文件
└── README.md
```

## 常见问题

### Q: 运行报错 `JSONDecodeError: Expecting value`

A: `playlist.json` 内容可能不是有效 JSON。请确认文件为网易云 API 返回的 JSON 数据。若 JSON 被 HTML 包裹(以 `<html>` 开头),脚本已自动处理。

### Q: 运行报错 `NameError: name 'datetime' is not defined`

A: 检查脚本顶部是否已导入 `from datetime import datetime`。

### Q: 歌曲数据不完整或字段为空

A: 网易云 API 可能因未登录或 Cookie 失效返回部分字段。请在浏览器登录网易云后重新获取 API 数据。
