# HTML 演示文稿模式

你已确认用户需要高保真 HTML 演示文稿。严格按以下步骤执行。

---

## 第一步：加载资源

调用 `read_skill_resource("pptx", "design_system.md")` 获取主题预设和色彩 token。

如果用户上传了附件，用 `read_chat_attachment` 读取。
如果 deck 需要引用数据，用 `search_web` 收集来源。

---

## 第二步：THINK — 分析需求

**必须调用 `mcp__sequential-thinking__sequentialthinking`** 进行深度分析，不得跳过。

在 sequential-thinking 中逐步分析：

- 受众是谁，他们期望什么风格
- 核心信息是什么，deck 要达成什么目标
- 哪个主题预设最合适（mckinsey / dark-tech / genspark）
- 预估 slide 数量
- 确认用户确实需要 HTML 而非 PPTX（如果用户说了"pptx"相关词汇，应该走原生PPTX模式）

---

## 第三步：PLAN — 规划结构

**必须调用 `mcp__sequential-thinking__sequentialthinking`** 进行结构规划，不得跳过。

在 sequential-thinking 中逐步规划：

- 叙事主线：开头 → 中间 → 结尾
- 为每页选择视觉结构（参考 slide_patterns.md 的 kind 作为布局灵感，自由选择）
- 确认主题 token（主色、强调色、字体）全程一致
- 规划 CSS 变量和组件复用策略

---

## 第四步：EXECUTE

调用 `run_generated_script` **一次**，参数：

```json
{
  "output_filenames_json": "[\"deck.html\"]"
}
```

生成脚本必须满足：

### 单次生成规则（不可违反）

- **一次调用生成所有 slide**——不得分批生成再合并
- 输出一个自包含的 `deck.html`，包含：
  - 所有 slide 内容
  - 完整的 CSS（主题色、字号、间距）
  - 导航功能（键盘左右箭头、点击翻页、进度条）
  - 内联 SVG/图标（如需要）
- 输出必须可在浏览器中独立打开，无外部依赖（CDN 字体除外）
- 默认 16:9 水平翻页，除非用户指定其他比例

### 脚本结构建议

对于 >10 页的 deck，保持脚本结构清晰：
1. 顶部定义主题 token（颜色、字体、间距）为 JS 变量
2. 用 helper 函数包装每页 slide 的 HTML
3. 按顺序生成所有 slide HTML
4. 组装最终文档

### Slide 结构约束

每个 slide 必须是顶层 `<section class="slide" data-slide-index="N">...</section>`，直接位于 `<body>` 内（或单个容器 div 内）。

---

## 第五步：TRACK

渲染完成后自检：

- [ ] deck.html 可在浏览器中独立打开
- [ ] 所有 slide 都在一个文件中（不是分批生成的）
- [ ] 导航功能正常（键盘、点击、进度条）
- [ ] 叙事有清晰的开头、展开、收尾
- [ ] 视觉美感：白底干净、色彩模块丰富沉稳
- [ ] 主题色彩和字号全程一致

告知用户："已生成 N 页演示文稿，可点击预览。"

---

## 修订

如果用户要求修改已生成的 HTML deck：

1. 读取现有 deck.html 内容
2. 修改受影响的 slide（不要从零重建整个 deck）
3. 重新调用 `run_generated_script` 生成更新后的完整 deck.html
4. 详见 `revision_playbook.md`

---

## 格式转换提醒

如果用户在 HTML deck 生成后说"转成pptx"/"转成PowerPoint"，**不要重新生成**。切换到转换模式（mode_convert.md），基于现有 HTML 内容转换为 .pptx。
