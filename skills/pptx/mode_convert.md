# 格式转换模式

用户已有一个演示文稿（通常是 deck.html），要求转换为 .pptx 格式。

**核心原则：转换 ≠ 重新生成。** 你必须保留原有内容和结构，只改变输出格式。

---

## 判断条件

当以下条件同时满足时进入此模式：
- 当前会话中已有生成的 deck.html（或用户上传了现有演示文稿）
- 用户说了"转成pptx"/"转成PowerPoint"/"转换格式"/"导出为pptx"等

如果用户没有现成的演示文稿，而是要求从零创建 pptx，那不是转换——应该走原生PPTX模式。

---

## 第一步：读取现有内容

1. 读取现有 deck.html 的完整内容（通过文件系统或 `read_chat_attachment`）
2. 分析其结构：
   - 有多少页 slide
   - 每页的标题和主要内容
   - 使用的主题色和字体
   - 每页的布局类型

---

## 第二步：加载资源

1. `read_skill_resource("pptx", "slide_patterns.md")` — slide kind 参考库
2. `read_skill_resource("pptx", "deck_spec_schema.json")` — spec 结构

---

## 第三步：THINK — 分析转换策略

**必须调用 `mcp__sequential-thinking__sequentialthinking`** 进行转换策略分析，不得跳过。

在 sequential-thinking 中逐步分析：

1. 现有 HTML 的结构（页数、每页内容类型、布局风格）
2. 为每页匹配最合适的 slide kind
3. 转换策略：哪些内容可以直接映射，哪些需要适配 PPTX 布局限制

---

## 第四步：EXECUTE

基于现有内容构建 deck-spec.json，然后调用 `render_pptx_from_spec`：

1. **提取**：从 HTML 中提取每页的标题、正文、数据点、图表描述
2. **映射**：为每页选择最匹配的 slide kind
3. **构建**：组装完整的 deck spec JSON
4. **渲染**：调用 `render_pptx_from_spec`

```json
{
  "title": "（从原 HTML 的 cover slide 提取）",
  "outputFile": "presentation.pptx",
  "themePreset": "（匹配原 HTML 的视觉风格）",
  "slides": [
    // 每页对应原 HTML 的一页，保留所有文字内容
  ]
}
```

---

## 硬性约束

- **不得忽略现有内容重新生成**——每页 slide 的文字内容必须来自原 HTML
- **不得改变叙事结构**——slide 顺序和分组保持不变
- **可以适当调整布局**——HTML 的复杂布局可能需要简化以适配 PPTX 的 kind 系统
- **主题尽量匹配**——选择与原 HTML 视觉风格最接近的 themePreset

---

## 第五步：TRACK

转换完成后自检：

- [ ] 输出的 .pptx 页数与原 HTML 一致
- [ ] 每页的标题和核心文字内容都保留了
- [ ] 没有丢失重要信息
- [ ] 主题风格与原 HTML 基本一致

告知用户："已将 HTML 演示文稿转换为 .pptx 格式（N 页），可直接下载编辑。"

---

## 常见情况

| 情况 | 处理 |
|------|------|
| HTML 中有复杂 CSS 动画 | 转换时忽略动画，保留静态内容 |
| HTML 中有内联 SVG 图表 | 在 PPTX 中用文字描述或简化图形替代 |
| HTML 页面内容过多 | 可拆分为多页 PPTX slide，但需告知用户 |
| 原 HTML 只有 3-4 页 | 正常转换，不要擅自扩充内容 |
