# 执行纪律（所有模式通用）

无论你选择了哪个模式（原生PPTX / HTML / 转换），都必须遵守以下四阶段纪律。跳过任何阶段都是严重违规。

**阶段播报规则**：每进入一个新阶段时，先输出一行阶段标记，让用户知道当前进度：

```
进入阶段 → 输出文本
THINK   → "## 🔍 THINK — 分析需求"
PLAN    → "## 📋 PLAN — 规划结构"
EXECUTE → "## ⚡ EXECUTE — 生成文件"
TRACK   → "## ✅ TRACK — 质量自检"
```

这行标记必须在该阶段的任何其他操作之前输出。

---

## 🔍 THINK — 理解需求，不生成任何内容

进入此阶段时先输出：`## 🔍 THINK — 分析需求`

**必须调用 `mcp__sequential-thinking__sequentialthinking`** 进行深度分析。不得跳过此工具调用直接输出分析结果。

在 sequential-thinking 中逐步思考以下问题：

1. 受众画像和期望
2. 核心目标和关键信息
3. 语气和视觉风格方向（白底 + 什么色系的强调色最合适）
4. 交付模式确认
5. 如需引用，先完成搜索/收集

完成 sequential-thinking 分析后才能进入 PLAN 阶段。

---

## 📋 PLAN — 输出结构化大纲，不渲染任何 slide

进入此阶段时先输出：`## 📋 PLAN — 规划结构`

**必须调用 `mcp__sequential-thinking__sequentialthinking`** 进行结构规划。不得跳过此工具调用直接输出大纲。

在 sequential-thinking 中逐步规划以下内容：

1. 叙事主线（开头 → 中间 → 结尾）
2. 每页 slide 的布局类型和用途（slide_patterns.md 中的 kind 作为参考库，不是强制约束——如果内容需要，可以自由组合或重复使用同一种 kind）
3. 视觉节奏感（通过内容密度、色彩对比、留白来创造节奏，而非机械地切换 kind）
4. 主题色彩一致性
5. 引用/数据来源映射
6. slide 数量和信息密度控制

完成 sequential-thinking 规划后才能进入 EXECUTE 阶段。

---

## ⚡ EXECUTE — 按模式指令渲染

进入此阶段时先输出：`## ⚡ EXECUTE — 生成文件`

执行你所在模式文件（mode_native_pptx.md / mode_html_deck.md / mode_convert.md）中的 EXECUTE 步骤。

---

## ✅ TRACK — 质量门禁，渲染后自检 + 视觉审计循环

进入此阶段时先输出：`## ✅ TRACK — 质量自检`

### 静态自检

完成渲染后，逐项检查：

- [ ] 叙事完整性：有清晰的开头、展开、收尾
- [ ] 视觉美感：白底干净、色彩模块丰富沉稳、卡片间距均匀、字号层次清晰、留白充足
- [ ] 信息密度：每页信息量适中，不过载也不空洞
- [ ] 主题一致性：颜色、字体、间距全程统一
- [ ] 文字完整性：所有文本都正确渲染，无截断或丢失
- [ ] 输出了可下载的成品文件（.pptx 或 .html）

### 视觉审计循环（最多 3 轮，必须执行，不得跳过）

**重要：这不是可选步骤。不能用口头打勾代替实际的工具调用。**

每次渲染完成后，必须执行以下操作：

1. **导出截图**：调用 `run_generated_script` 执行 `export_slides.js`，将每页导出为 PNG
   - Windows 自动使用 PowerPoint COM
   - 无 Office 时自动 fallback 到 LibreOffice
   - 都不可用时跳到第 3 步

2. **视觉审查**：读取导出的 PNG 截图，逐页检查填充率、对齐、字号一致性、文本溢出、箭头指向

3. **结构化审计**：在 `render_pptx_from_spec` 的 spec 中加入 `_auditOutput` 字段，渲染后调用 `run_generated_script` 执行 `audit_deck_structural.js`

4. **判断**：error ≥ 1 或 warning ≥ 3 → 修改 spec → 重新 `render_pptx_from_spec`

5. **循环**：最多 3 轮，直到 0 error 且 warning ≤ 2

详细操作步骤见 `mode_native_pptx.md` 第五步。

---

## 美学原则（所有主题通用）

以下是视觉设计的指导原则：

1. **白色为底**：所有内容页使用白色或极浅灰背景，让色彩模块成为视觉焦点
2. **色彩模块**：通过卡片的彩色顶边/左边、图标色块、数据高亮来注入色彩，而非大面积色块背景
3. **沉稳丰富**：使用深灰/墨色作为文字主色，搭配 2-3 种饱和度适中的强调色（如深青、琥珀、紫罗兰）
4. **卡片圆角**：PPTX 用 ROUNDED_RECTANGLE (0.1" radius)，HTML 用 `border-radius: 8px`
5. **间距呼吸**：元素之间留足空间，宁可少放内容也不要拥挤
6. **字号层次**：页标题 28-36px bold → 卡片标题 14-18px bold → 正文 11-14px → 注释 8-10px
7. **视觉权重**：封面 > 分节页 > 内容页，通过字号和色彩浓度体现
8. **画布自适应**：支持 LAYOUT_WIDE (13.33"×7.5") / LAYOUT_16x9 (10"×5.625") / LAYOUT_16x10 (10"×6.25") / LAYOUT_4x3 (10"×7.5")
