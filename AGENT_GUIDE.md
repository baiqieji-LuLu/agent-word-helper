# Agent Guidelines: Word Document Generation

Do NOT read the source code of `word_helper.py`; simply use the commands and `--help`.
Runtime Environment: Please replace this with your local Python path.

## Core Commands
1. Generate from a string (with formulas):
   python word_helper.py create --text "content" -o "output.docx" --format markdown --title "Title"
2. Generate from a file (Recommended: write long texts to .md first, then convert):
   python word_helper.py create "input.md" -o "output.docx" --format markdown
3. Read body text and tables:
   python word_helper.py read "file.docx"
4. Read all parts (headers, footers, footnotes, endnotes):
   python word_helper.py read "file.docx" --all
5. Structured read (JSON):
   python word_helper.py read "file.docx" --json -o "output.json"

## Strict Rules
- Do NOT write your own Python scripts to parse/generate docx files. Always use `word_helper.py`.
- Do NOT invoke local COM/Word software.
- Retry the same command at most once. If it still fails, send the complete error log to the user.
- For situations not covered by this guide, refer to `README.md`.


# Agent 使用规范：Word 文档生成

不需要阅读 word_helper.py 源码；用命令和 `--help` 即可。
运行环境：请替换为你本地的 Python 路径。

## 核心命令
1. 字符串生成（含公式）：
   python word_helper.py create --text "内容" -o "输出.docx" --format markdown --title "标题"
2. 文件生成（长篇先写 .md 再转换，推荐）：
   python word_helper.py create "输入.md" -o "输出.docx" --format markdown
3. 读正文与表格：
   python word_helper.py read "文件.docx"
4. 读全部（页眉页脚、脚注尾注）：
   python word_helper.py read "文件.docx" --all
5. 结构化读取：
   python word_helper.py read "文件.docx" --json -o "输出.json"

## 规则
- 禁止自己写 python 脚本解析/生成 docx，一律用 word_helper.py。
- 禁止调用本机 COM/Word 软件。
- 同一条命令最多重试 1 次，仍失败把完整报错发给用户。
- 本指南未覆盖的情况再查 README.md。