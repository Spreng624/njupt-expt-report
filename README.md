# NJUPT 实验报告 Markdown → DOCX Skill

这个 Skill 先把带 YAML 头部的 Markdown 实验内容写完并校验，再通过占位符替换生成南京邮电大学实验报告 DOCX。版式基准为用户手工调整的 `实验报告模板.docx`，项目内保存固定快照。

## 目录

```text
njupt-expt-report/
├─ SKILL.md                        # Codex Skill 入口与强制工作流
├─ references/                     # Markdown 与模板占位符规范
├─ templates/实验报告模板.docx     # 固定版式快照
├─ scripts/validate_markdown.py    # 生成前校验 Markdown
├─ scripts/build_report.py         # Markdown → DOCX
├─ scripts/validate_docx.py        # 结构与字号校验
├─ scripts/render_preview.ps1      # Word 导出 PDF / PNG
├─ examples/report.md              # 可直接运行的示例
├─ examples/report.docx            # 匿名化示例输出
└─ output/                         # 生成结果
```

## 生成报告

先生成并校验 Markdown，再转换 DOCX。不要直接修改生成的 DOCX。

一键完成 Markdown 校验、DOCX 转换和 DOCX 校验：

```powershell
powershell -ExecutionPolicy Bypass -File convert.ps1 `
  -Markdown examples/report.md `
  -Output output/示例实验报告.docx
```

也可以分步运行：

```powershell
python scripts/validate_markdown.py examples/report.md
python scripts/build_report.py examples/report.md -o output/示例实验报告.docx
python scripts/validate_docx.py output/示例实验报告.docx
```

如果 Python 不在 `PATH` 中，可通过参数指定解释器：

```powershell
$py = "path\to\python.exe"
& $py scripts/validate_markdown.py examples/report.md
& $py scripts/build_report.py examples/report.md -o output/示例实验报告.docx
& $py scripts/validate_docx.py output/示例实验报告.docx
```

生成前预览：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/render_preview.ps1 `
  -InputDocx output/示例实验报告.docx `
  -OutputDirectory output/preview-final
```

## Markdown 约定

正文从“一、实验目的和要求”开始，不要写 Markdown H1：

```markdown
---
course_name: 示例课程
experiment_name: 实验一 示例实验
experiment_date: 2026-01-01
academic_year: 2025 / 2026
term: 第一
department: 示例学院
instructor: 示例教师
student_name: 示例学生
student_id: B00000000
class_name: B00000000
major: 示例专业
experiment_type: 验证
experiment_hours: 2
experiment_time: 2026-01-01
---

**一、** **实验目的和要求**

正文内容。

**二、** **实验环境(实验设备)**

正文内容。

**三、** **自定义中间标题**

正文内容。

**四、** **实验小结**

正文内容。

**五、** **支撑毕业要求指标点**

正文内容。

**六、** **指导教师评语**
```

第四节和第五节必须各有 3 个三级标题；第六节正文必须留空，因为模板直接保留指导教师评分表。

也接受 `## 一、实验目的和要求` 形式的二级标题，但生成结果不使用 Word 的标题样式。

## 自动化保证

- 封面“实验报告”为初号字（42 pt）。
- 封面字段、正文页眉、二级标题为四号字（14 pt）。
- 正文、列表、小标题、表格、图注和代码为五号字（10.5 pt）。
- 全文只使用 `Normal` 和普通列表样式，不创建或引用 `Title`、`Subtitle`、`Heading 1–9`。
- 正文元信息表直接使用模板的 13 列主表，不新增内嵌元信息表。
- Markdown 必须恰好包含 6 个二级标题。转换器只在 `{{B1}}`–`{{B5_n}}` 占位符位置插入内容。
- 转换器不删除、插入或克隆正文行，也不修改行高、分页、单元格合并、边框或评分区。
- 文本报告中除 `word/document.xml` 外的所有 OOXML 包部件都与模板逐字节一致；只有插入图片时才允许新增媒体和关系部件。
- 模板封面分页、正文外框、评分区和页面尺寸保持不变。
- 生成后可用 `validate_docx.py` 自动检查禁用样式、标题、字号和正文表结构。

详细规则见 [Markdown 规范](references/markdown-authoring.md)。

Skill 工作流见 [SKILL.md](SKILL.md)。模板占位符清单见 [references/template-placeholders.md](references/template-placeholders.md)。

## 更新模板快照

项目默认使用 `templates/实验报告模板.docx`。只有在用户确认新版模板后，才替换快照并更新本文件中的 SHA-256。用于公开分发时，先运行 `scripts/sanitize_docx_metadata.py` 清除 DOCX 作者、最后修改者和 Office 自定义保存记录。

当前模板快照 SHA-256：

```text
25AA3B3E8822E4D9FDAB2242CD8D6439A4776D517E82F6AFC6815138DFF95A54
```
