# DOCX from Markdown

将 Markdown 文档转换为符合 `.docx` 模板格式的 Word 文档。

本 skill 是一个**流水线执行器**，串联调用 `docx-deterministic-generator` skill 生成的两个脚本。

---

## 1. 前提条件

本 skill 的 `scripts/` 目录必须包含以下文件：

```
skills/docx-from-md/scripts/
├── run_pipeline.py           ← 本 skill 自带，流水线编排脚本
├── fix_md_tables.py          ← 本 skill 自带，表格预处理脚本
├── md_to_json.py             ← 用户从 docx-deterministic-generator 产出复制过来
└── render_docx.py            ← 用户从 docx-deterministic-generator 产出复制过来
```

如果 `md_to_json.py` 或 `render_docx.py` 不存在，运行时会报错并提示：

> 请先使用 `docx-deterministic-generator` skill 分析 `.docx` 模板，
> 生成 `md_to_json.py` 和 `render_docx.py`，然后复制到本目录。

---

## 2. 使用方法

### 方式一：让 Claude 执行

告诉 Claude：

> 用 `$docx-from-md` 把 `设计手册/skill测试/xxx.md` 转成 Word

Claude 会调用 `scripts/run_pipeline.py` 完成转换。

### 方式二：直接命令行运行

```bash
python3 skills/docx-from-md/scripts/run_pipeline.py \
  --md 你的文档.md \
  --template 模板.docx \
  --output generated.docx
```

参数说明：

| 参数 | 必填 | 说明 |
|---|---|---|
| `--md` | ✅ | 输入的 Markdown 文档（按模板章节格式编写） |
| `--template` | ✅ | 原始 `.docx` 模板 |
| `--output` | ✅ | 输出 `.docx` 路径 |
| `--keep-json` | ❌ | 保留中间 JSON 文件（默认生成后删除） |

---

## 3. 流水线步骤

```
[1/3] 修复非矩形表格  (fix_md_tables.py)
[2/3] Markdown → JSON  (md_to_json.py)
[3/3] JSON → DOCX      (render_docx.py)
完成：generated.docx
```

中间 JSON 文件默认保存在临时位置，成功后自动删除。加 `--keep-json` 保留在输出同目录（如 `generated.json`）。

---

## 4. 两个 skill 的分工

```
docx-deterministic-generator        docx-from-md
─────────────────────────────        ─────────────────────────
输入：模板.docx                      输入：文档.md + 模板.docx
产出：render_docx.py                 产出：generated.docx
      md_to_json.py
      ...
```

`docx-deterministic-generator` 是**一次性**的（模板不变就不用重跑）。  
`docx-from-md` 是**每次写文档都要跑的**（MD → DOCX）。

---

## 5. 错误场景

| 错误 | 原因 | 解决方式 |
|---|---|---|
| 缺少 `md_to_json.py` 或 `render_docx.py` | 没有运行过模板分析 | 先跑 `docx-deterministic-generator`，复制脚本过来 |
| 非矩形表格 | interface spec table 不同子段列数不一致 | 自动修复（`fix_md_tables.py`），无需手动处理 |
| MD 格式错误 | 标题层级或表头不符合模板 | 修改 MD，使其符合模板章节格式 |
| JSON 不符合 schema | MD 内容有缺失或格式错误 | 检查 `--keep-json` 输出的 JSON，对照 schema 修正 |
| DOCX 生成失败 | 模板与脚本不匹配 | 确认 `--template` 和生成脚本的 `md_to_json.py`/`render_docx.py` 来自同一次模板分析 |
