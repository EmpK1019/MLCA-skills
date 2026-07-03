# 视觉审计流程

本文件定义 PPTX 生成后的质量审计流程。在 `deck_workflow.md` 的 TRACK 阶段执行。

**重要：这是一个强制流程，不是可选步骤。每次渲染后都必须执行，不能用口头自检代替实际工具调用。**

---

## 审计触发条件

每次 `render_pptx_from_spec` 成功渲染后，**必须**自动进入审计流程。跳过审计视为严重违规。

---

## 审计流程

**关键约束**：`run_generated_script` 在独立 workdir 中执行，所有输出文件必须写入 workdir（即 `process.env.MLCA_OUTPUT_DIR` 或 `process.cwd()`），否则工具无法收集到生成文件。不要使用上一个步骤返回的绝对路径，而应使用 workdir 内的文件名。

### Step 1：导出截图（内部审计用，不暴露给用户）

调用 `run_generated_script`，使用 `require()` 直接调用模块导出函数（不要用 `execSync` 子进程），将输出目录设为 workdir：

```javascript
const path = require('path');
const fs = require('fs');

// 所有输出必须写入 workdir
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

导出策略（自动选择，无需手动配置）：
1. Windows：PowerPoint COM（通过 PowerShell 脚本）
2. macOS：AppleScript 调用 Mac 版 PowerPoint
3. 以上都不可用：LibreOffice headless（PPTX → PDF → PNG）
4. 都不可用：跳过截图，直接执行 Step 3

工具调用参数：
- `output_filenames_json`: `["export_result.json"]`
- `collect_all_outputs`: `false`（关键！防止 PNG 截图作为附件暴露给用户）

### Step 2：视觉审查（截图可用时）

将导出的 PNG 截图作为附件读取（`read_chat_attachment`），逐页检查：

| 类别 | 检查项 | 不通过标准 |
|------|--------|-------------|
| 填充 | 页面是否有大片留白 | 留白 > 30% 且非 cover/section/closing |
| 文字 | 文字是否可读 | 截断、溢出、重叠 |
| 文字 | 同级字号是否一致 | 页标题差 > 4pt，卡片标题差 > 2pt |
| 对齐 | 元素是否对齐 | 同行元素 y 偏差 > 0.05"，同列 x 偏差 > 0.05" |
| 对齐 | 网格模块是否等大 | 同行卡片宽/高差 > 10% |
| 连线 | 箭头指向是否正确 | 方向与语义相反或断开 |
| 配色 | 主题色是否一致 | 同类元素使用不同色系 |

### Step 3：结构化审计（始终执行，不可跳过）

**此步骤不依赖截图，无论 Step 1 是否成功都必须执行。**

审计数据在渲染时自动生成，无需在 spec 中手动声明 `_auditOutput` 字段。渲染完成后，workdir 中会自动出现 `*_audit.json` 文件。

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

审计规则见下方"结构化审计规则"。

### Step 4：判断是否需要重渲染

- **error ≥ 1** → 必须修正后重新渲染
- **warning ≥ 3** → 建议修正后重新渲染
- **0 error 且 warning ≤ 2** → 通过

### Step 5：修正并重渲染（如需要）

根据 Step 2/3 的问题列表，修改 deck spec 中有问题的 slide：

- 填充率低 → 增加内容、增大元素尺寸、减少边距
- 字号不一致 → 统一为同级最小字号
- 文本溢出 → 缩减文本、增大文本框、降低字号
- 对齐偏差 → 调整 x/y 坐标
- 重叠 → 调整间距或元素位置

修改后重新调用 `render_pptx_from_spec`，最多重渲染 2 次（共 3 轮）。

### Step 6：完成通知

通过所有检查后，告知用户："已生成 N 页演示文稿（.pptx），可直接下载编辑。"

---

## 结构化审计规则

| 规则 | 严重度 | 条件 |
|------|--------|------|
| fill-rate | error | 内容填充率 < 18%（非 cover/section/closing） |
| fill-rate | warning | 内容填充率 < 28% |
| title-font-consistency | warning | 同页标题字号差异 > 4pt |
| tiny-text | warning | 字号 < 7pt 的文本元素 |
| text-overflow-risk | error | 文字密度 > 1.5× 文本框面积 |
| text-overflow-risk | info | 文字密度 > 1.0× 文本框面积 |
| element-overlap | warning | 文本元素重叠面积 > 小元素 50% |
| cover-title | warning | 封面缺少 ≥20pt 标题 |
| closing-content | warning | 结尾页缺少实质内容 |

---

## 平台兼容性

### Windows
- PowerPoint COM 导出：通过 PowerShell 脚本调用 `export_slides_to_images.ps1`
- LibreOffice headless：检测 `D:\tmp\LibreOfficePortable\App\libreoffice\program\soffice.exe` 等路径

### macOS
- PowerPoint 导出：通过 AppleScript 调用 Mac 版 PowerPoint
- LibreOffice headless：检测 `/Applications/LibreOffice.app/Contents/MacOS/soffice`

### 两者都不可用时
- 跳过 Step 1-2（截图和视觉审查）
- **Step 3 结构化审计仍然必须执行**