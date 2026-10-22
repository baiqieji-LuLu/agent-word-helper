# word_helper.py

一个基于纯 Python 的 CLI 工具，用于在不依赖本地 Microsoft Word 软件和 COM 自动化组件的情况下，安全地创建和读取 Word DOCX 文档。支持 Markdown 转 DOCX 排版、基于 LaTeX 的原生 OMML 数学公式生成，以及从现有 DOCX 文件中提取纯文本与表格。

本项目专为 AI Agent 和自动化文档工作流设计，底层基于 `python-docx` 和 OOXML 构建，彻底告别 Word COM 自动化依赖。数学公式的转换通过 `LaTeX → MathML → OMML` 管线实现，并需要配合外部的 XSLT 样式表使用。

## 为什么会有这个项目？

如果你正在构建 AI Agent 或自动化工作流，并且需要与 Word 文档交互，大概率面临过以下两个问题（我不知道你有没有遇到过，反正我遇到了）：

1. COM 接口崩溃：当 Agent 通过 `win32com` 在后台静默调用本地 Windows Word 时，经常会因为沙盒权限限制或残留的孤儿进程，触发“Word 无法创建工作文件，请检查临时环境变量”的报错，或者陷入无限弹窗的死循环。
2. 公式灾难：大模型生成的数学公式，插入 Word 后全变成了乱码、无法编辑的 Unicode 字符，或者是低分辨率的模糊图片。

`word_helper.py` 是一个黑盒 CLI 工具，正是为了解决上述问题而生。

## 核心特性

- 无 Word COM 依赖：无需启动 Microsoft Word 即可创建和读取 DOCX 文件，彻底避开与 COM 相关的自动化崩溃。
- 原生 OMML 公式渲染：内置 `LaTeX -> MathML -> OMML` 转换管线。只需传入标准的 `$...$`（行内公式）或 `$$...$$`（独立公式），即可输出原生、在 Word 中双击可编辑的数学公式。
- Markdown 转 Docx：支持常见的 Markdown 语法（标题、列表、表格、代码块），并内置了标准化的排版样式（A4 纸、中文字体支持等）。
- 双向读写：不仅能生成文档，还能从现有的 `.docx` 文件中提取纯文本和表格，方便你的 Agent 进行读取分析。

## 安装指南

1. Clone 本仓库。
2. 安装 Python 依赖：

   ```bash
   pip install python-docx latex2mathml lxml
   ```

## 前置条件：MML2OMML.XSL

为了将 MathML 转换为 OMML，本工具需要依赖微软官方的 MML2OMML.XSL 样式表。由于版权限制，该文件并未包含在本仓库中。

如何获取该文件：

如果你的 Windows 电脑安装了 Microsoft Office，请在你的 Office 安装目录中搜索 MML2OMML.XSL（通常位于 `C:\Program Files\Microsoft Office\root\Office16\`）。

将该文件复制到与 word_helper.py 相同的目录下。

如果你没有安装 Office，可以在网上搜索下载该文件，或从 Office 的 ISO 镜像中提取。

## 快速开始

1. 从字符串创建文档（含公式）：

   注：在 PowerShell 下运行含 `$` 符号的命令时，请使用单引号包裹，以防变量展开。

   ```bash
   python word_helper.py create --text '质能方程: $E=mc^2$' -o output.docx --format markdown
   ```

2. 从 Markdown 文件创建文档：

   ```bash
   python word_helper.py create notes.md -o notes.docx --format markdown
   ```

3. 读取文档中的正文与表格：

   ```bash
   python word_helper.py read output.docx
   ```

## 为 AI Agent 打造

本仓库内置了一份开箱即用的 AGENT_GUIDE.md。

当你在给 LLM Agent（如 Claude、GPT、DeepSeek 等）写 Prompt 时，只需将这份指南注入到系统提示词（System Prompt）中。它会强制命令 Agent 使用本脚本，而不是去瞎编 win32com 代码，从而为你省下几十万 Token 的浪费，并避免无休止的调试死循环。

## 进阶测试：让 AI Agent 替你跑一遍（推荐方式）

由于 word_helper.py 的核心定位是为大模型服务，最贴合真实场景的测试方式，就是让 Agent 自己去调用它。

第一步：准备环境

确保你的 Agent 拥有当前项目目录的读取和执行权限，并将仓库里的 AGENT_GUIDE_WORD.md 喂给它（或直接放入系统提示词中）。

第二步：向 Agent 发送测试指令

直接复制下面这句话发给你的 Agent：

“请先阅读 AGENT_GUIDE_WORD.md 了解规范。然后，使用 word_helper.py 将demo.md 转换为一个名为demo_test.docx 的 Word 文档。转换成功后，请使用 read 命令验证一下内容，并向我汇报生成的文件路径。”

第三步：观察与验证

正常情况下，Agent 会自动执行命令并生成文件。用 Word 打开 demo_test.docx，你应该能看到：完美的标题与表格排版，以及由 LaTeX 完美转换而来的、可双击编辑的原生数学公式，并且整个过程中没有触发任何一次 COM 报错。

## 开源协议

本项目基于 MIT License 开源。
