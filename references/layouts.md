# 页面布局库（Layouts）

10 种常用 PPT 页面布局，每种都可用 pptxgenjs 实现。
所有布局都兼容 16:9 宽屏（13.33" x 7.5"）。

---

## 使用方式

每种布局对应的 JS 辅助函数，在 `scripts/generate-ppt.js` 中预定义。
生成时直接调用函数，传入主题和内容即可。

---

## 布局索引

| 编号 | 名称 | 适合场景 | 对应函数 |
|:---:|------|---------|---------|
| L01 | **封面页 Cover** | 首页，大标题 + 副标题 + 日期 | `addCover(slide, theme, title, subtitle)` |
| L02 | **章节幕封 Section** | 每个 Part 的过渡页 | `addSectionSlide(slide, theme, num, title)` |
| L03 | **标题+正文 Content** | 标准内容页 | `addContentSlide(slide, theme, title, body)` |
| L04 | **大字报 Data Hero** | 关键数据/指标展示 | `addDataHero(slide, theme, number, label)` |
| L05 | **左文右图 Image-Text** | 图文混排 | `addImageText(slide, theme, title, body, imgDesc)` |
| L06 | **图片网格 Grid** | 多图展示 | `addImageGrid(slide, theme, title, images)` |
| L07 | **对比页 Comparison** | Before/After 对比 | `addComparison(slide, theme, left, right)` |
| L08 | **三段式 Cards** | 三个并列要点 | `addThreeCards(slide, theme, title, cards)` |
| L09 | **时间线 Timeline** | 发展历程/流程 | `addTimeline(slide, theme, title, steps)` |
| L10 | **引用页 Quote** | 名人名言/核心观点 | `addQuoteSlide(slide, theme, quote, author)` |
| L11 | **Bento Grid** | 多信息混合展示 | `addBentoGrid(slide, theme, title, items)` |
| L12 | **结束页 Closing** | 谢谢/联系方式 | `addClosing(slide, theme)` |

---

## L01 · 封面页

```js
function addCover(slide, theme, title, subtitle, meta) {
  slide.background = { fill: theme.bg };
  // 顶部装饰线
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 4, h: 0.04, fill: { color: theme.accent1 }
  });
  // 大标题
  slide.addText(title, {
    x: 0.8, y: 2.0, w: 11.5, h: 1.5,
    fontSize: 44, color: theme.fg, fontFace: "Arial", bold: true
  });
  // 副标题
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.8, y: 3.6, w: 11.5, h: 0.7,
      fontSize: 22, color: theme.accent1, fontFace: "Arial"
    });
  }
  // 底部信息
  if (meta) {
    slide.addText(meta, {
      x: 0.8, y: 6.5, w: 5, h: 0.4,
      fontSize: 11, color: theme.textGray, fontFace: "Arial"
    });
  }
}
```

---

## L02 · 章节幕封

```js
function addSectionSlide(slide, theme, sectionNum, title, subtitle) {
  slide.background = { fill: theme.bg };
  // 左侧竖线
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.8, y: 2.0, w: 0.04, h: 2.5, fill: { color: theme.accent1 }
  });
  // 章节编号
  slide.addText(sectionNum, {
    x: 1.2, y: 2.0, w: 2, h: 0.5,
    fontSize: 14, color: theme.accent1, fontFace: "Arial", bold: true
  });
  // 标题
  slide.addText(title, {
    x: 1.2, y: 2.8, w: 10, h: 1.0,
    fontSize: 36, color: theme.fg, fontFace: "Arial", bold: true
  });
  // 副标题
  if (subtitle) {
    slide.addText(subtitle, {
      x: 1.2, y: 4.0, w: 10, h: 0.5,
      fontSize: 16, color: theme.textGray, fontFace: "Arial"
    });
  }
}
```

---

## L03 · 标题+正文（标准内容页）

```js
function addContentSlide(slide, theme, title, bodyItems) {
  slide.background = { fill: theme.bg };
  slide.addText(title, {
    x: 0.8, y: 1.0, w: 11.5, h: 0.8,
    fontSize: 28, color: theme.fg, fontFace: "Arial", bold: true
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.8, y: 1.9, w: 3, h: 0.02, fill: { color: theme.accent1 }
  });
  bodyItems.forEach((item, i) => {
    slide.addText(item, {
      x: 1.0, y: 2.4 + i * 0.7, w: 11, h: 0.6,
      fontSize: 16, color: theme.fg, fontFace: "Arial",
      bullet: { code: "2022" }, // bullet char
    });
  });
}
```

---

## L04 · 数据大字报

```js
function addDataHero(slide, theme, number, label, note) {
  slide.background = { fill: theme.bg };
  slide.addShape(pptx.ShapeType.rect, {
    x: 1.0, y: 1.5, w: 11.3, h: 4.5,
    fill: { color: theme.cardBg },
    rectRadius: 0.08,
  });
  slide.addText(number, {
    x: 1.0, y: 2.0, w: 11.3, h: 2.5,
    fontSize: 72, color: theme.accent1, fontFace: "Arial", bold: true,
    align: "center", valign: "middle",
  });
  slide.addText(label, {
    x: 1.0, y: 4.2, w: 11.3, h: 0.6,
    fontSize: 18, color: theme.fg, fontFace: "Arial",
    align: "center",
  });
  if (note) {
    slide.addText(note, {
      x: 1.0, y: 4.9, w: 11.3, h: 0.4,
      fontSize: 12, color: theme.textGray, fontFace: "Arial",
      align: "center",
    });
  }
}
```

---

## L05 · 左文右图

```js
function addImageText(slide, theme, title, body, imgPlaceholder) {
  slide.background = { fill: theme.bg };
  // 左列文字
  slide.addText(title, {
    x: 0.8, y: 1.5, w: 5.5, h: 0.8,
    fontSize: 28, color: theme.fg, fontFace: "Arial", bold: true
  });
  slide.addText(body, {
    x: 0.8, y: 2.6, w: 5.5, h: 3.0,
    fontSize: 16, color: theme.fg, fontFace: "Arial"
  });
  // 右列图片占位
  slide.addShape(pptx.ShapeType.rect, {
    x: 7.0, y: 1.5, w: 5.3, h: 5.0,
    fill: { color: theme.cardBg },
    line: { color: theme.accent1, width: 0.5 },
    rectRadius: 0.04,
  });
  slide.addText(imgPlaceholder || "[图片占位]", {
    x: 7.0, y: 3.5, w: 5.3, h: 1.0,
    fontSize: 14, color: theme.textGray, fontFace: "Arial",
    align: "center", valign: "middle",
  });
}
```

---

## L06 · 图片网格

```js
function addImageGrid(slide, theme, title, images) {
  // images: [{label, desc}]
  slide.background = { fill: theme.bg };
  slide.addText(title, {
    x: 0.8, y: 0.5, w: 11.5, h: 0.7,
    fontSize: 24, color: theme.fg, fontFace: "Arial", bold: true
  });
  const cols = 3, perRow = Math.ceil(images.length / cols);
  images.forEach((img, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const xb = 0.8 + col * 4.0, yb = 1.5 + row * 2.8;
    slide.addShape(pptx.ShapeType.rect, {
      x: xb, y: yb, w: 3.6, h: 2.2,
      fill: { color: theme.cardBg },
      line: { color: theme.accent1, width: 0.5 },
      rectRadius: 0.04,
    });
    slide.addText(img.label || "[图]", {
      x: xb, y: yb + 0.3, w: 3.6, h: 1.2,
      fontSize: 14, color: theme.fg, fontFace: "Arial",
      align: "center", valign: "middle",
    });
    slide.addText(img.desc || "", {
      x: xb + 0.2, y: yb + 1.6, w: 3.2, h: 0.5,
      fontSize: 10, color: theme.textGray, fontFace: "Arial",
      align: "center",
    });
  });
}
```

---

## L07 · 对比页

```js
function addComparison(slide, theme, leftTitle, leftBody, rightTitle, rightBody) {
  slide.background = { fill: theme.bg };
  slide.addText(leftTitle, {
    x: 0.8, y: 1.0, w: 5.5, h: 0.7,
    fontSize: 24, color: theme.fg, fontFace: "Arial", bold: true
  });
  slide.addText(rightTitle, {
    x: 7.0, y: 1.0, w: 5.5, h: 0.7,
    fontSize: 24, color: theme.accent1, fontFace: "Arial", bold: true,
    align: "right",
  });
  // Separator
  slide.addShape(pptx.ShapeType.rect, {
    x: 6.6, y: 1.2, w: 0.02, h: 5.0,
    fill: { color: theme.textGray, opacity: 30 },
  });
  slide.addText(leftBody, {
    x: 0.8, y: 2.0, w: 5.5, h: 4.0,
    fontSize: 15, color: theme.fg, fontFace: "Arial"
  });
  slide.addText(rightBody, {
    x: 7.0, y: 2.0, w: 5.5, h: 4.0,
    fontSize: 15, color: theme.fg, fontFace: "Arial"
  });
}
```

---

## L08 · 三段式卡片

```js
function addThreeCards(slide, theme, title, cards) {
  // cards: [{icon, title, desc}]
  slide.background = { fill: theme.bg };
  slide.addText(title, {
    x: 0.8, y: 0.5, w: 11.5, h: 0.8,
    fontSize: 28, color: theme.fg, fontFace: "Arial", bold: true
  });
  cards.forEach((card, i) => {
    const xb = 0.8 + i * 4.0;
    slide.addShape(pptx.ShapeType.rect, {
      x: xb, y: 1.8, w: 3.6, h: 4.5,
      fill: { color: theme.cardBg },
      line: { color: theme.accent1, width: 0.5 },
      rectRadius: 0.08,
    });
    slide.addText(card.title, {
      x: xb + 0.3, y: 2.2, w: 3.0, h: 0.6,
      fontSize: 20, color: theme.fg, fontFace: "Arial", bold: true
    });
    slide.addText(card.desc, {
      x: xb + 0.3, y: 3.0, w: 3.0, h: 3.0,
      fontSize: 14, color: theme.textGray, fontFace: "Arial"
    });
  });
}
```

---

## L09 · 时间线

```js
function addTimeline(slide, theme, title, steps) {
  // steps: [{year/num, title, desc}]
  slide.background = { fill: theme.bg };
  slide.addText(title, {
    x: 0.8, y: 0.5, w: 11.5, h: 0.8,
    fontSize: 28, color: theme.fg, fontFace: "Arial", bold: true
  });
  // Timeline line
  const lineY = 4.0;
  slide.addShape(pptx.ShapeType.rect, {
    x: 1.0, y: lineY, w: 11.3, h: 0.02,
    fill: { color: theme.accent1 },
  });
  steps.forEach((step, i) => {
    const xb = 1.0 + (i * 11.3) / Math.max(steps.length - 1, 1);
    // Dot
    slide.addShape(pptx.ShapeType.ellipse, {
      x: xb - 0.1, y: lineY - 0.1, w: 0.22, h: 0.22,
      fill: { color: theme.accent1 },
    });
    // Label above
    slide.addText(step.num || "", {
      x: xb - 0.8, y: lineY - 1.5, w: 1.6, h: 0.4,
      fontSize: 12, color: theme.accent1, fontFace: "Arial", bold: true,
      align: "center",
    });
    slide.addText(step.title, {
      x: xb - 0.8, y: lineY - 1.0, w: 1.6, h: 0.8,
      fontSize: 13, color: theme.fg, fontFace: "Arial",
      align: "center",
    });
    // Desc below
    slide.addText(step.desc || "", {
      x: xb - 0.8, y: lineY + 0.4, w: 1.6, h: 1.5,
      fontSize: 11, color: theme.textGray, fontFace: "Arial",
      align: "center",
    });
  });
}
```

---

## L10 · 引用页

```js
function addQuoteSlide(slide, theme, quote, author) {
  slide.background = { fill: theme.bg };
  // 大引号
  slide.addText("\u201C", {
    x: 0.8, y: 1.0, w: 1.5, h: 1.5,
    fontSize: 72, color: theme.accent1, fontFace: "Arial", bold: true
  });
  slide.addText(quote, {
    x: 1.5, y: 2.0, w: 10.5, h: 3.0,
    fontSize: 24, color: theme.fg, fontFace: "Arial",
    italic: true, lineSpacing: 32,
  });
  slide.addText("\u2014 " + (author || ""), {
    x: 1.5, y: 5.2, w: 10.5, h: 0.5,
    fontSize: 14, color: theme.textGray, fontFace: "Arial",
    align: "right",
  });
}
```

---

## L11 · Bento Grid（信息混合展示）

```js
function addBentoGrid(slide, theme, title, items) {
  // items: [{title, desc, size: "large"|"small"|"wide"}]
  slide.background = { fill: theme.bg };
  slide.addText(title, {
    x: 0.8, y: 0.5, w: 11.5, h: 0.7,
    fontSize: 26, color: theme.fg, fontFace: "Arial", bold: true
  });
  items.forEach((item, i) => {
    const positions = [
      { x: 0.8, y: 1.6, w: 5.5, h: 2.8 },  // large left
      { x: 6.8, y: 1.6, w: 5.7, h: 1.3 },  // small right top
      { x: 6.8, y: 3.1, w: 2.7, h: 1.3 },  // small right bottom L
      { x: 9.8, y: 3.1, w: 2.7, h: 1.3 },  // small right bottom R
    ];
    const p = positions[i] || positions[0];
    slide.addShape(pptx.ShapeType.rect, {
      x: p.x, y: p.y, w: p.w, h: p.h,
      fill: { color: theme.cardBg },
      line: { color: theme.accent1, width: 0.5 },
      rectRadius: 0.06,
    });
    slide.addText(item.title, {
      x: p.x + 0.3, y: p.y + 0.3, w: p.w - 0.6, h: 0.5,
      fontSize: i === 0 ? 22 : 16, color: theme.fg, fontFace: "Arial", bold: true
    });
    slide.addText(item.desc || "", {
      x: p.x + 0.3, y: p.y + 1.0, w: p.w - 0.6, h: p.h - 1.5,
      fontSize: i === 0 ? 15 : 12, color: theme.textGray, fontFace: "Arial"
    });
  });
}
```

---

## L12 · 结束页

```js
function addClosing(slide, theme, thanks, subtitle) {
  slide.background = { fill: theme.bg };
  // 顶部装饰线
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 0.04, fill: { color: theme.accent1 }
  });
  slide.addText(thanks || "谢谢大家", {
    x: 0.8, y: 2.5, w: 12, h: 1.2,
    fontSize: 48, color: theme.fg, fontFace: "Arial", bold: true,
    align: "center",
  });
  if (subtitle) {
    slide.addShape(pptx.ShapeType.rect, {
      x: 5.5, y: 3.8, w: 2.5, h: 0.03, fill: { color: theme.accent1 }
    });
    slide.addText(subtitle, {
      x: 0.8, y: 4.0, w: 12, h: 0.6,
      fontSize: 18, color: theme.textGray, fontFace: "Arial",
      align: "center",
    });
  }
}
```
