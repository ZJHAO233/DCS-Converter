<div align="center">

# DCS Converter

**Excel 试验条件表 → Markdown / Word 智能转换工具**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![HTML](https://img.shields.io/badge/HTML5-E34F26?style=flat&logo=html5&logoColor=white)]()
[![CSS](https://img.shields.io/badge/CSS3-1572B6?style=flat&logo=css3&logoColor=white)]()
[![JS](https://img.shields.io/badge/JavaScript-F7DF1E?style=flat&logo=javascript&logoColor=black)]()

将 DCS 分散控制系统的试验条件 Excel 表格，一键转换为结构清晰的 Markdown 文档或 Word 文件。

[快速开始](#快速开始) · [功能特性](#功能特性) · [转换示例](#转换示例) · [Excel 格式要求](#excel-格式要求)

</div>

---

## 快速开始

### 1. 启动服务

```bash
# PowerShell 5.1+
.\server.ps1
```

浏览器自动打开 `http://localhost:18888/`

### 2. 上传 Excel

拖拽或点击上传区域导入 `.xlsx` / `.xls` 文件

### 3. 转换导出

选择 Sheet → 点击转换 → 下载 `.md` / `.docx` 或复制内容

---

## 功能特性

<table>
<tr>
<td width="50%">

**导入**
- 拖拽 / 点击上传
- 支持 `.xlsx` / `.xls`
- 解析进度条显示
- 错误原因详细提示

</td>
<td width="50%">

**导出**
- 下载 Markdown 文件
- 一键复制到剪贴板
- 下载 Word (.docx) 文件
- 自动命名（按 Sheet 名）

</td>
</tr>
<tr>
<td>

**智能转换**
- 按序号自动分组
- 三层嵌套逻辑识别
- 括号自动包裹
- 标题逻辑自动推断

</td>
<td>

**界面**
- 暗色 / 亮色主题切换
- 多 Sheet 目录切换
- 合并 / 单 Sheet 模式
- 偏好自动保存

</td>
</tr>
</table>

---

## 转换示例

### 基础转换

**Excel 输入：**

| 序号 | 条件内容 | | |
|------|----------|--|--|
| 1 | 输煤急停（取反） | | |
| 2 | 或 | 输煤联锁模式 | |
| | 与 | 输煤自动模式 | |

**Markdown 输出：**

```markdown
### 1.启动允许条件（与）

- 输煤急停（取反）
- 输煤联锁模式且输煤自动模式
```

### 三层嵌套

**Excel 输入：**

| 序号 | | 条件内容 | | |
|------|-|----------|--|--|
| 1 | | J0BNR02J1_LS | | |
| 2 | 或 | 输煤自动模式 | | |
| | | 上煤预启动，5S脉冲 | | |
| | | | 或 | 上煤方式1.1-1.4选择有效 |
| | | | | 上煤方式2.1-2.4选择有效 |

**Markdown 输出：**

```markdown
### 1.自动启条件（与）

- J0BNR02J1_LS
- 输煤自动模式且上煤预启动，5S脉冲且（上煤方式1.1-1.4选择有效或上煤方式2.1-2.4选择有效）
```

---

## Excel 格式要求

### 列结构

| A列 | B列 | C列 | D列 | E列 |
|:---:|:---:|:---:|:---:|:---:|
| 序号 | 分隔符/条件 | 分隔符/条件 | 条件 | 条件 |

### 识别规则

| 内容 | 格式 | 转换结果 |
|------|------|----------|
| 子标题 | `一、xxx` | `## 一、xxx` |
| 分组标题 | `1.xxx（与）` | `### 1.xxx（与）` |
| 分隔符 | `与` / `或` | 条件连接符 |
| 条件内容 | 具体文字 | `- 条件内容` |

---

## 项目结构

```
DCS-Converter/
├── dcs_converter.html    # 主程序（单文件应用）
├── xlsx.min.js           # SheetJS Excel 解析库
├── server.ps1            # 本地 HTTP 服务器
└── README.md
```

---

## 技术栈

- **前端**: 原生 HTML / CSS / JavaScript
- **Excel 解析**: [SheetJS](https://sheetjs.com/)
- **Word 导出**: [html-docx-js](https://github.com/nicktomlin/html-docx-js)
- **服务器**: PowerShell HTTP Server

---

## 许可

[MIT License](LICENSE)
