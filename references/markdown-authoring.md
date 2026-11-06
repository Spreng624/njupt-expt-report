# Markdown Authoring Specification

## Content-First Gate

Complete and validate the Markdown before generating DOCX. The Markdown is the editable source of truth; regenerated DOCX files are disposable outputs.

## YAML Front Matter

The file must begin with:

```markdown
---
course_name: 示例课程
experiment_name: 实验一 示例实验
experiment_date: 2026-01-01
academic_year: 2025 / 2026
term: 第一
department: 示例学院
college: 示例学院
instructor: 示例教师
student_name: 示例学生
student_id: B00000000
class_name: B00000000
major: 示例专业
experiment_type: 验证
experiment_hours: 2
experiment_time: 2026-01-01
---
```

Required keys are validated by `scripts/validate_markdown.py`. `college` is optional and defaults to `department`.

## Section Contract

Use exactly six bold second-level headings:

```markdown
**一、** **实验目的和要求**
**二、** **实验环境(实验设备)**
**三、** **自定义实验方法与内容**
**四、** **实验小结**
**五、** **支撑毕业要求指标点**
**六、** **指导教师评语**
```

`## 一、实验目的和要求` syntax is also accepted, but the generated DOCX never uses Word heading styles.

Sections 4 and 5 must each contain exactly three third-level headings:

```markdown
### （一）第一项
...
### （二）第二项
...
### （三）第三项
...
```

Section 6 must remain empty because the template has no `{{B6}}` body slot.

## Content Syntax

Supported block types:

- paragraphs;
- unordered and ordered lists;
- fenced code blocks;
- pipe tables;
- images written as `![caption](relative/path.png)`;
- quoted notes beginning with `>`.

Supported inline syntax:

- `**bold**`;
- `` `inline code` ``;
- `[label](url)`, rendered as `label（url）`.

Use relative image paths. A missing image is a hard validation error.

## Screenshot Handoff

The Agent executes the experiment first and writes a complete report. For every evidence point,
leave a named screenshot placeholder:

```markdown
<!-- SHOT: screenshots/01-interface.png | 运行 ip -4 addr show eth1 -->

![确认实验接口](screenshots/01-interface.png)
```

The user reproduces the command and saves their screenshot at the exact relative path. The
Markdown validator reports every pending screenshot before DOCX conversion.

For terminal-only experiments, the Agent or user may generate the PNG directly with
`scripts/capture_terminal.py`. See [terminal-capture.md](terminal-capture.md).

## Quality Gate

Before conversion:

- record actual commands, observed results, and errors;
- use images for relevant execution evidence;
- explain failure causes and resolution attempts;
- do not leave empty required sections;
- do not place experimental prose before the first heading.
