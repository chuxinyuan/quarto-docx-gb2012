# 党政机关公文 Quarto Extensions

这是一个专门用于生成符合《党政机关公文格式》（GB/T 9704-2012）国家标准的 Word 版公文的 Quarto 扩展。

## 开发准备

- Quarto 官方文档

    扩展开发指南：参考 Quarto 官方提供的扩展开发文档，学习如何创建和配置 Quarto 扩展，了解扩展的结构、配置文件编写方法以及如何将自定义模板和脚本集成到扩展中。
    Word 输出格式定制：深入学习 Quarto 关于 Word 输出格式定制的文档，掌握如何通过引用模板文件来控制 Word 文档的生成格式，包括样式设置、页面布局等细节。

- Lua 编程语言

    Lua 基础教程：学习 Lua 编程语言的基础知识，包括语法、数据类型、控制结构、函数定义等，为编写 Lua 过滤器脚本打下坚实的基础。
    Pandoc Lua 过滤器开发：参考 Pandoc 官方文档中关于 Lua 过滤器开发的部分，了解如何利用 Lua 脚本对文档转换过程进行干预和定制，
    掌握如何编写高效的过滤器脚本来处理复杂的格式要求。

- GB/T 9704—2012 标准文件（公文格式要求）

    官方标准文档：`./assets/9704-2012-gbt-cd-300.pdf`，
    仔细研读其中的各项规定，深入理解公文格式的具体要求，为插件开发提供准确的标准依据。

- 相关案例研究

    现有 Quarto 扩展案例：`./assets/apaquarto`，研究它的结构、配置文件和代码实现，借鉴其成功经验和优秀做法。
    公文排版软件分析：分析市场上现有的公文排版软件，了解它们在处理 GB/T 9704—2012 标准时的功能特点和操作流程，从中汲取有益的设计理念。

## 项目目标

定制一个 Quarto Extensions，利用这个插件可以很方便的生成满足《党政机关公文格式》（GB/T 9704-2012）国家标准格式要求的 Word 版公文，插件既方便自己又方便分享给别人。

- 完全符合 GB/T 9704-2012 标准
- 自动应用公文格式要求
- 支持各种公文类型（通知、报告、函、纪要、讲话稿等）
- 正确的字体、字号、行距设置
- 标准的页面布局和页边距
- 支持密级、紧急程度标识
- 支持发文字号、签发人等元素

## 技术实现

本扩展包含以下核心组件：

1. **`_extension.yml`** - 扩展配置文件
2. **`gbmetadata.lua`** - 处理文档元数据
3. **`gbstyler.lua`** - 应用 Word 样式
4. **`gblinenumber.lua`** - 行号处理

## 扩展安装

- 命令行安装

```bash
quarto use template chuxinyuan/quarto-docx-gb2012
```

- R 代码安装

```r
quarto::quarto_use_template("chuxinyuan/quarto-docx-gb2012", no_prompt = TRUE)
```

开发人员请将整个 `_extensions/official-gb2012` 文件夹复制到您的 Quarto 项目根目录，方便调试。

## 使用指南

### 快速开始

项目包含一个完整的示例：`example.qmd`

运行以下命令查看效果：

```bash
quarto render example.qmd --to docx
```

开始使用前，请确保系统已安装公文正常显示所需要的字体：

- 仿宋_GB2312
- 小标宋体
- 宋体
- 黑体
- 楷体

### 文档结构

````qmd
---
title: "公文标题"
format: docx
author: "发文机关"
date: "2024年10月22日"
document_type: "通知"
issuing_authority: "发文机关名称"
document_number: "发文字号"
urgency_level: "普通"
confidentiality_level: "公开"
signatory: "签发人姓名"
recipients: "主送机关"
---

::: {.identifier}
公开
:::

::: {.issuing-authority}
某某机关
:::

::: {.document-number}
某办发〔2024〕15号
:::

::: {.signatory}
签发人：张三
:::

# 公文标题

正文内容...

::: {.date}
发文机关
年月日
:::
````

### 元数据字段

| 字段 | 说明 | 示例 |
|------|------|------|
| `document_type` | 公文类型 | "通知"、"报告"、"函" |
| `issuing_authority` | 发文机关 | "某某机关办公室" |
| `document_number` | 发文字号 | "某办发〔2024〕15号" |
| `urgency_level` | 紧急程度 | "特急"、"加急"、"普通" |
| `confidentiality_level` | 密级 | "绝密"、"机密"、"秘密"、"公开" |
| `signatory` | 签发人 | "张三" |
| `recipients` | 主送机关 | "各部门、各单位" |
| `attachments` | 附件列表 | ["附件1：xxx", "附件2：xxx"] |

## 技术支持

如遇到问题，请检查：
1. Quarto 版本 >= 1.4.549
2. 系统字体安装情况
3. 文件路径和权限

如果扩展无法正常加载：
1. 确保 `_extensions` 文件夹在项目根目录
2. 检查 `_extension.yml` 文件格式是否正确
3. 使用 `quarto check` 检查 Quarto 安装

## 项目许可

本项目采用 MIT 许可证。
