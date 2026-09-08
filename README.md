# ChemDraw Native PPT

![ChemDraw Native PPT 效果图](assets/synthetic-example/hero-figure.png)

<p align="center">
  <a href="https://github.com/LHH90538/chemdraw-native-ppt"><img src="https://img.shields.io/github/license/LHH90538/chemdraw-native-ppt?style=flat-square&label=license" alt="License"></a>
  <a href="https://github.com/LHH90538/chemdraw-native-ppt"><img src="https://img.shields.io/badge/Codex-Skill-111827?style=flat-square" alt="Codex Skill"></a>
  <a href="https://www.cambridgesoft.com/Ensemble_for_Chemistry/ChemDraw/Default.aspx"><img src="https://img.shields.io/badge/ChemDraw-native%20objects-2563eb?style=flat-square" alt="Native ChemDraw objects"></a>
  <img src="https://img.shields.io/badge/PowerPoint-editable-f97316?style=flat-square" alt="Editable PowerPoint">
  <img src="https://img.shields.io/badge/platform-Windows-0ea5e9?style=flat-square" alt="Windows">
</p>

<p align="center"><a href="#中文">中文</a> · <a href="#english">English</a></p>

## 中文

**把化学文献图片变成可编辑的 PowerPoint，而不是截图。**

`chemdraw-native-ppt` 是一个面向 Codex 的专业 skill，用于将化学反应图、合成路线和文献示意图重建为可编辑幻灯片：分子保留为真正的 **ChemDraw 原生对象**，文字、箭头、标签和简单图形保留为可编辑的 PowerPoint 对象。

### 安装

直接对 Codex 说：

> 请帮我安装 `chemdraw-native-ppt` 这个 skill。

安装后重启 Codex。使用时说：

> 请使用 `chemdraw-native-ppt` 把这张化学文献图转换成可编辑的 PowerPoint，保留分子结构、特殊键、文字、箭头和原图排版，并交付 PPTX 和可复现项目。

### 能做什么

- 将化学结构重建为可双击编辑的 ChemDraw 对象
- 将文字、箭头、标签和简单图形重建为 PowerPoint 对象
- 支持多分子、多页幻灯片和反应路线图
- 保存源图、项目清单、启动脚本、预览图和验证信息
- 对模糊或遮挡的结构明确标注不确定性，不凭空猜测连接关系

### 环境要求

- Windows
- 已安装并可使用的 ChemDraw 和 Microsoft PowerPoint
- Windows PowerShell 5.1 或更高版本
- Python 3.10+（仅用于初始化项目）

### 必须安装 ChemDraw 吗？

如果要生成可编辑的 **ChemDraw 原生分子对象**，必须安装 Windows 版 ChemDraw；如果还要完成最终的 PPTX 导出，则同时必须安装 Microsoft PowerPoint。

仅让 Codex 分析图片、重建结构或准备项目文件时，可以暂时不安装 ChemDraw。但没有 ChemDraw，无法完成原生对象嵌入，也无法保证用户之后能在 PowerPoint 中双击编辑分子。ChemDraw、PowerPoint 及其许可证不包含在本项目中。本项目是“图形重建工作流”，不是端到端 OCSR 模型；复杂结构仍需根据原图进行人工或代理复核。

### 项目文档

- [SKILL.md](SKILL.md)：Codex 的完整工作流程和验收标准
- [项目格式](references/project-format.md)：坐标、缩放、清单和输出
- [重建规范](references/reconstruction.md)：结构识别和证据记录
- [ChemDraw 样式](references/cdxml-style.md)：键型、标签和排版
- [排错指南](references/troubleshooting.md)：导入和 Office 自动化问题

## English

**Turn chemical literature figures into editable PowerPoint—not screenshots.**

`chemdraw-native-ppt` is a Codex skill for reconstructing chemical reaction schemes, synthesis routes, and literature figures as editable slides. Molecules remain genuine **native ChemDraw objects**, while text, arrows, labels, and simple graphics remain editable PowerPoint objects.

### Install

Simply ask Codex:

> Please install the `chemdraw-native-ppt` skill for me.

Then restart Codex and ask:

> Please use `chemdraw-native-ppt` to convert this chemical literature figure into an editable PowerPoint. Preserve the molecular structures, special bonds, text, arrows, and layout, and deliver both the PPTX and a reproducible project.

### Features

- Native, editable ChemDraw objects embedded in PowerPoint
- Editable PowerPoint text, arrows, labels, and simple graphics
- Multi-molecule, multi-slide, and reaction-scheme support
- Source references, project manifests, launchers, previews, and validation output
- Explicit uncertainty reporting for obscured or ambiguous chemistry

### Requirements

- Windows
- ChemDraw and Microsoft PowerPoint
- Windows PowerShell 5.1+
- Python 3.10+ for project initialization

### Do users need ChemDraw?

Yes—generating editable **native ChemDraw molecule objects** requires the Windows version of ChemDraw. Microsoft PowerPoint is also required to complete the final PPTX export.

Codex can analyze a figure, reconstruct structures, and prepare a project without ChemDraw. However, native object embedding and reliable double-click editing in PowerPoint cannot be completed without ChemDraw. ChemDraw, PowerPoint, and their licenses are not bundled. This is a figure-reconstruction workflow, not an end-to-end OCSR model; complex structures still require review against the source figure.

Released under the [MIT License](LICENSE). Not affiliated with ChemDraw, Revvity, Microsoft, or OpenAI.
