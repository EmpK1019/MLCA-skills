---
name: pptx
description: 'Use this skill whenever a PowerPoint deck is involved: creating a
  new presentation, generating a polished pitch deck or report deck, editing an
  existing .pptx file, extracting text from slides, reworking a template, or producing
  a downloadable slide deck from user notes, attachments, or research. Trigger
  for .pptx/.ppt files, and also when the user mentions deck, slides, presentation,
  pitch deck, roadshow, keynote,汇报,演示文稿, or 幻灯片.'
license: Proprietary. LICENSE.txt has complete terms
distribution_source: bundled
publisher_name: Anthropic, PBC
id: pptx
version: '3.1'
trigger_hints:
  - .pptx
  - .ppt
  - ppt
  - slides
  - deck
  - presentation
  - pitch deck
  - keynote
  - roadshow
  - 演示文稿
  - 幻灯片
  - 汇报
allowed_tools:
  - render_pptx_from_spec
  - run_generated_script
  - read_skill_resource
  - read_chat_attachment
  - search_web
  - mcp__sequential-thinking__sequentialthinking
mandatory_reads:
  - deck_workflow.md
  - audit_workflow.md
output_contract:
  primary_artifact: pptx
  optional_sidecars:
    - deck-spec.json
    - sources.md
    - deck.html
scripts:
  - path: scripts/extract_text.py
    purpose: Extract slide text from PPTX files for chat attachment reading.
  - path: scripts/thumbnail.py
    purpose: Generate slide thumbnails for template review.
  - path: scripts/add_slide.py
    purpose: Add or duplicate slides in unpacked PPTX packages.
  - path: scripts/clean.py
    purpose: Clean unpacked PPTX package artifacts before repacking.
  - path: scripts/render_deck.js
    purpose: Reference renderer for PptxGenJS deck generation.
  - path: scripts/export_slides.js
    purpose: Export PPTX slides to PNG images for visual audit (PowerPoint COM → LibreOffice fallback).
  - path: scripts/export_slides_to_images.ps1
    purpose: PowerShell script for PowerPoint COM slide export on Windows.
  - path: scripts/audit_deck_structural.js
    purpose: Structural audit engine — checks fill rate, font consistency, text overflow, alignment.
  - path: scripts/office/unpack.py
    purpose: Unpack PPTX archives for XML-level editing.
  - path: scripts/office/pack.py
    purpose: Repack edited PPTX archives.
  - path: scripts/office/validate.py
    purpose: Validate PPTX package structure after editing.
  - path: scripts/office/soffice.py
    purpose: Run LibreOffice conversions in a controlled environment.
tags:
  - presentation
  - pptx
  - slides
  - deck
form: []
prompt_template: ''
category: contextual
icon: skill
show_in_template_bar: false
show_in_drawer: true
order: 104
---

# PPTX Skill

你的任务是帮用户生成或处理演示文稿。收到请求后，**第一件事是判断模式**，然后按该模式的指令执行。不要跳过模式判断直接开始生成。

---

## 模式判断（必须先完成）

按以下规则判断，**从上到下匹配第一个命中的**：

| # | 条件 | 模式 | 下一步 |
|---|------|------|--------|
| 1 | 用户有现成 .pptx 文件要编辑/修改/提取文字 | **编辑模式** | `read_skill_resource("pptx", "editing.md")` |
| 2 | 会话中已有 deck.html，且用户说"转成pptx/转成PowerPoint/导出pptx" | **转换模式** | `read_skill_resource("pptx", "mode_convert.md")` |
| 3 | 用户消息含"pptx""PowerPoint""可编辑""要在PPT里改" | **原生PPTX模式** | `read_skill_resource("pptx", "mode_native_pptx.md")` |
| 4 | 用户明确说"HTML演示""网页演示""高保真HTML" | **HTML模式** | `read_skill_resource("pptx", "mode_html_deck.md")` |
| 5 | 用户未指定格式（默认） | **原生PPTX模式** | `read_skill_resource("pptx", "mode_native_pptx.md")` |

### 硬规则（不可违反）

- 含"pptx""PowerPoint""可编辑"关键词 → **必须**走原生PPTX模式，不得选HTML
- 含"转换""转成""改成pptx"且已有 deck.html → **必须**走转换模式，不得重新生成
- 只有用户**明确要求** HTML/网页演示时才走 HTML 模式
- 未指定格式时**默认原生PPTX模式**

---

## 所有模式的共同纪律

无论哪个模式，都必须遵守 deck_workflow.md 中的四阶段纪律：

1. **THINK 阶段**：**必须调用 `mcp__sequential-thinking__sequentialthinking`** 分析需求（受众、目标、风格、主题选择）
2. **PLAN 阶段**：**必须调用 `mcp__sequential-thinking__sequentialthinking`** 规划结构（叙事主线、每页布局、色彩一致性）
3. **EXECUTE 阶段**：按模式指令渲染
4. **TRACK 阶段**：自检质量

每个阶段都需要认真思考并输出分析结果，不要跳过直接渲染。THINK 和 PLAN 阶段不调用 sequential-thinking 视为违规。

---

## 质量标准

- 一个好的 deck 有叙事主线，不只是堆内容的页面
- 每页 slide 必须有明确目的：定位、说服、对比、解释、量化、或总结
- 大多数 slide 应包含超越纯文字段落的视觉结构：卡片、统计数字、时间线、对比、网格、引用、流程图、或分节页
- 封面和结尾页的视觉权重应高于中间页
- 控制密度——如果一页感觉拥挤，拆分它
- **白底为王**——所有内容页使用白色背景，通过色彩模块（卡片色条、数据高亮）注入视觉丰富度

---

## 参考资源索引

确定模式后，按需加载以下资源：

| 需要时 | 调用 |
|--------|------|
| 选主题预设 | `read_skill_resource("pptx", "design_system.md")` |
| 查 slide kind 参考 | `read_skill_resource("pptx", "slide_patterns.md")` |
| PPTX spec 格式 | `read_skill_resource("pptx", "deck_spec_schema.json")` |
| PPTX spec 示例 | `read_skill_resource("pptx", "deck_spec_example.json")` |
| 引用/来源规则 | `read_skill_resource("pptx", "citation_playbook.md")` |
| 修订已有 deck | `read_skill_resource("pptx", "revision_playbook.md")` |
| PptxGenJS API | `read_skill_resource("pptx", "pptxgenjs.md")` |

---

## 核心循环

1. 判断模式 → 2. 加载模式指令 → 3. THINK → 4. PLAN → 5. EXECUTE → 6. TRACK

每一步都不可跳过。如果你发现自己想直接开始渲染——停下来，回到第 3 步。
