# 原生 PPTX 模式

你已确认用户需要可编辑的 .pptx 文件。严格按以下步骤执行。

---

## 第一步：加载资源

**在同一轮中并行调用**以下 `read_skill_resource`（不要一个一个加载，一次性全部调用）：

1. `read_skill_resource("pptx", "design_system.md")` — 主题预设和色彩 token
2. `read_skill_resource("pptx", "slide_patterns.md")` — slide kind 参考库
3. `read_skill_resource("pptx", "deck_spec_schema.json")` — spec 结构定义
4. `read_skill_resource("pptx", "deck_spec_example.json")` — spec 完整示例

如果用户上传了附件，用 `read_chat_attachment` 读取（可与上面并行）。
如果 deck 需要引用数据，用 `search_web` 收集来源。

---

## 第二步：THINK — 分析需求

**必须调用 `mcp__sequential-thinking__sequentialthinking`** 进行深度分析，不得跳过。

在 sequential-thinking 中逐步分析：

- 受众是谁，他们期望什么风格
- 核心信息是什么，deck 要达成什么目标
- 哪个主题预设最合适（mckinsey / dark-tech / genspark）
- 预估 slide 数量

---

## 第三步：PLAN — 规划结构

**必须调用 `mcp__sequential-thinking__sequentialthinking`** 进行结构规划，不得跳过。

在 sequential-thinking 中逐步规划：

- 叙事主线：开头（cover + 背景/问题）→ 中间（论证/展示）→ 结尾（总结/行动号召）
- 为每页选择 `kind`——参考 slide_patterns.md 中的类型，根据内容需要自由选择，可以重复使用同一种 kind
- 确认主题 token（主色、强调色、字体）全程一致
- 如有引用需求，规划 sources slide

规划完成后输出 deck-spec.json 草稿。

---

## 第四步：EXECUTE

调用 `render_pptx_from_spec`，传入完整的 deck spec JSON：

```json
{
  "title": "Deck 标题（必填）",
  "outputFile": "presentation.pptx（必填）",
  "themePreset": "mckinsey | dark-tech | genspark",
  "theme": { ... },
  "slides": [
    { "kind": "cover", "title": "...", ... },
    ...
  ]
}
```

**硬性约束**：
- `title` 字段必填，不能为空
- `outputFile` 字段必填，必须以 `.pptx` 结尾
- `themePreset` 必须是 `mckinsey`、`dark-tech`、`genspark` 之一

**禁止**：
- 不得使用 `run_generated_script` 生成 .pptx（那是 HTML 模式的工具）
- 不得使用 python-pptx 从零手写代码生成

---

## 第五步：TRACK — 视觉审计（必做，不得跳过）

**这不是可选步骤。每次渲染完成后都必须执行以下操作，不能用口头自检代替。**

**关键约束**：`run_generated_script` 在独立 workdir 中执行，所有输出文件必须写入 workdir（即 `process.env.MLCA_OUTPUT_DIR` 或 `process.cwd()`），否则工具无法收集到生成文件。不要使用上一个步骤返回的绝对路径，而应使用 workdir 内的文件名。

### 5.1 导出截图（内部审计用，不暴露给用户）

渲染成功后，立即调用 `run_generated_script` 导出每页 PNG。

**重要**：
- 设置 `collect_all_outputs: false`，防止 PNG 截图作为附件暴露给用户
- 不要使用 `execSync` 调用外部 node 进程——改用 `require()` 直接调用模块导出函数
- 输出目录设为 workdir

```javascript
const path = require('path');
const fs = require('fs');

// 所有输出必须写入 workdir（MLCA_OUTPUT_DIR）
const outputDir = process.env.MLCA_OUTPUT_DIR || process.cwd();
fs.mkdirSync(outputDir, { recursive: true });

// 在 workdir 中查找上一步渲染生成的 PPTX 文件
const pptxFiles = fs.readdirSync(outputDir).filter(f => f.endsWith('.pptx'));
if (pptxFiles.length === 0) {
  throw new Error('未在 workdir 中找到 PPTX 文件');
}
const pptxPath = path.join(outputDir, pptxFiles[0]);

// 直接 require 模块，避免 execSync 子进程
const scriptDir = path.join(
  process.env.LCA_PROJECT_ROOT || process.env.MLCA_PROJECT_ROOT || '',
  'engine_data', 'skills', 'pptx', 'scripts'
);
const { exportSlides } = require(path.join(scriptDir, 'export_slides.js'));

// 将结果写入 JSON 汇总文件
exportSlides({
  inputFile: pptxPath,
  outputDir: outputDir,
  width: 1920,
  height: 1080,
}).then(result => {
  fs.writeFileSync(path.join(outputDir, 'export_result.json'), JSON.stringify(result, null, 2));
  console.log('Exported', result.slideCount, 'slides, method:', result.method);
}).catch(err => {
  fs.writeFileSync(path.join(outputDir, 'export_result.json'), JSON.stringify({ success: false, error: err.message }));
  console.error('Export failed:', err.message);
});
```

工具调用参数：
- `output_filenames_json`: `["export_result.json"]`
- `collect_all_outputs`: `false`（关键！防止 PNG 截图作为附件暴露给用户）

如果截图导出失败（无 PowerPoint 且无 LibreOffice），跳到 5.3 执行结构化审计。

### 5.2 视觉审查（截图可用时）

将导出的 PNG 截图作为附件读取（`read_chat_attachment`），逐页检查：

- 每页是否铺满（无大片留白）
- 元素是否对齐、对称
- 同级文字大小是否一致
- 文本是否截断或溢出
- 箭头/连接线是否指向正确

### 5.3 结构化审计（始终执行）

**审计数据在渲染时自动生成**，无需在 spec 中手动声明 `_auditOutput` 字段。渲染完成后，workdir 中会自动出现 `*_audit.json` 文件（文件名与 outputFile 对应）。

调用 `run_generated_script` 执行结构化审计：

```javascript
const path = require('path');
const fs = require('fs');

const workdir = process.env.MLCA_OUTPUT_DIR || process.cwd();

// 在 workdir 中查找审计数据文件（文件名包含 _audit）
const auditFiles = fs.readdirSync(workdir).filter(f => f.includes('_audit') && f.endsWith('.json'));
if (auditFiles.length === 0) {
  throw new Error('未在 workdir 中找到审计数据文件');
}
const auditPath = path.join(workdir, auditFiles[0]);

const scriptDir = path.join(
  process.env.LCA_PROJECT_ROOT || process.env.MLCA_PROJECT_ROOT || '',
  'engine_data', 'skills', 'pptx', 'scripts'
);
const { runAudit } = require(path.join(scriptDir, 'audit_deck_structural.js'));

const auditData = JSON.parse(fs.readFileSync(auditPath, 'utf8'));
const auditResult = runAudit(auditData);

// 将审计结果写入文件（必须写入 workdir 才能被收集）
const resultPath = path.join(workdir, 'audit_result.json');
fs.writeFileSync(resultPath, JSON.stringify(auditResult, null, 2), 'utf8');
console.log('Audit result written to:', resultPath);
console.log('Summary:', JSON.stringify(auditResult.summary));
```

`output_filenames_json` 设为 `["audit_result.json"]`。
`collect_all_outputs` 设为 `false`（审计结果是内部使用的，不暴露给用户）。

检查审计结果：
- **error ≥ 1** → 必须修正后重新渲染
- **warning ≥ 3** → 建议修正后重新渲染
- **0 error 且 warning ≤ 2** → 通过

### 5.4 修正与重渲染（如需要）

根据 5.2/5.3 的问题列表，修改 deck spec 中有问题的 slide：

- 填充率低 → 增加内容、增大元素尺寸、减少边距
- 字号不一致 → 统一为同级最小字号
- 文本溢出 → 缩减文本、增大文本框、降低字号
- 对齐偏差 → 调整 x/y 坐标
- 重叠 → 调整间距或元素位置

修改后重新调用 `render_pptx_from_spec`，最多重渲染 2 次（共 3 轮）。

### 5.5 完成通知

通过所有检查后，告知用户："已生成 N 页演示文稿（.pptx），可直接下载编辑。"

**注意**：
- 只向用户展示 .pptx 文件作为交付物
- 审计截图（PNG）和审计结果（JSON）是内部审查用的，不要在最终回复中列出或提及
- 如果审计过程中发现并修正了问题，可以简要说明修正了什么

**绝对不要在没有执行 5.1-5.3 的情况下声称自检通过。口头打勾不等于实际审计。**

---

## 修订

如果用户要求修改已生成的 deck：

1. 读取现有 deck-spec.json
2. 修改受影响的 slide（不要从零重建）
3. 重新调用 `render_pptx_from_spec`
4. 重新执行第五步的视觉审计
5. 详见 `revision_playbook.md`