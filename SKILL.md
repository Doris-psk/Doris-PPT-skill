---
name: ppt-generator
description: "Professional PowerPoint presentation generation using pptxgenjs. Use when the user asks to: (1) create a PowerPoint/PPT/PPTX/slides/presentation, (2) generate courseware or lecture slides, (3) make a slide deck for an event or speech, (4) convert content/text/outline into a presentation, (5) create a pitch deck or proposal presentation. Features 10 built-in high-end design themes inspired by guizang-ppt-skill (归藏PPT) magazine-style and Swiss-style design systems. Works with Node.js and the pptxgenjs library to produce editable .pptx files with custom theming, layouts, and styling."
---

# PPT Generator

Generates professional, editable PowerPoint (.pptx) presentations using pptxgenjs.

拥有 **10 套高端设计主题**，覆盖杂志风、瑞士国际主义风、深色科技风三大体系。
风格灵感来自 **guizang-ppt-skill（歸藏PPT）** 和 **ppt-agent-skill**。

## 快速开始

```bash
cd scripts/
node your-presentation.js
```

## 工作流

1. **需求澄清** — 问用户：主题、页数、受众、风格偏好、有无素材
2. **选主题** — 从 `references/themes.md` 的 10 套主题中推荐一套（见风格选择矩阵）
3. **规划大纲** — 创建 slide-by-slide 大纲，确认
4. **写脚本** — 引用主题 + 布局辅助函数，生成内容
5. **运行 & 交付** — `node script.js` → `.pptx` 文件
6. **自检** — 按 `references/checklist.md` 逐项检查

## 10 套内置主题（从 references/themes.md 选用）

### 🖋 电子杂志风（5套）—— 归藏风格
| 主题 | 色系 | 适合 |
|:----:|:----:|:----:|
| 🖋 墨水经典 | 暖米白+纯墨黑 | 通用/商业发布 |
| 🌊 靛蓝瓷 | 瓷白+深靛蓝 | 科技/研究/AI |
| 🌿 森林墨 | 象牙+森林绿 | 自然/文化/非虚构 |
| 🍂 牛皮纸 | 暖米+深棕 | 人文/历史/文学 |
| 🌙 沙丘 | 沙色+炭灰 | 艺术/设计/时尚 |

### 🇨🇭 瑞士国际主义风（4套）
| 主题 | 锚点色 | 适合 |
|:----:|::------:|:----:|
| 🔵 克莱因蓝 IKB | `#002FA7` | 商业/AI产品 |
| 🟡 柠檬黄 | `#FFD500` | 年轻品牌/零售 |
| 🟢 柠檬绿 | `#C5E803` | 生态/健康 |
| 🟠 安全橙 | `#FF6B35` | 新闻/工业/运动 |

### 🌌 深色科技风（1套）
| 主题 | 色系 | 适合 |
|:----:|:----:|:----:|
| 🌌 深空科技 | 深蓝+青色荧光 | 科技发布会/AI |

## 12 种页面布局（从 references/layouts.md 选用）

| 编号 | 名称 | 函数 | 说明 |
|:----:|:----:|:----:|:------|
| L01 | 封面页 | `addCover()` | 大标题+副标题+日期 |
| L02 | 章节幕封 | `addSectionSlide()` | Part 过渡页 |
| L03 | 标准内容 | `addContentSlide()` | 标题+要点列表 |
| L04 | 数据大字报 | `addDataHero()` | 突出关键数字 |
| L05 | 左文右图 | `addImageText()` | 图文混排 |
| L06 | 图片网格 | `addImageGrid()` | 多图展示 |
| L07 | 对比页 | `addComparison()` | Before/After |
| L08 | 三段卡片 | `addThreeCards()` | 并列要点 |
| L09 | 时间线 | `addTimeline()` | 流程/发展历程 |
| L10 | 引用页 | `addQuoteSlide()` | 名言/核心观点 |
| L11 | Bento Grid | `addBentoGrid()` | 混合信息展示 |
| L12 | 结束页 | `addClosing()` | 谢谢/联系方式 |

## 脚本参考

### 标准脚本结构

```js
const PptxGenJS = require("pptxgenjs");
const pptx = new PptxGenJS();

// 1. 设置
pptx.defineLayout({ name: "WIDE", width: 13.33, height: 7.5 });
pptx.layout = "WIDE";

// 2. 选主题（从 references/themes.md 选择）
const THEME = {
  name: "墨水经典",
  bg: "F1EFEA", fg: "0A0A0B",
  cardBg: "E8E5DE", accent1: "0A0A0B",
  accent2: "6B5B4F", textGray: "6B5B4F",
};

// 3. 添加布局辅助函数（从 references/layouts.md 复制）
// ...

// 4. 创建幻灯片
const s1 = pptx.addSlide();
addCover(s1, THEME, "主标题", "副标题", "作者 · 日期");

const s2 = pptx.addSlide();
addSectionSlide(s2, THEME, "01", "章节标题", "副标题");

// ... 更多幻灯片

// 5. 保存
pptx.writeFile({ fileName: "presentation.pptx" })
  .then(() => console.log("PPT generated!"))
  .catch(e => console.error(e));
```

## 设计原则

### 通用原则
1. **一致性** — 全 deck 使用同一套主题色
2. **卡片布局** — 内容放在带边框的圆角矩形中
3. **视觉层级** — 大标题(36-48pt) > 中标题(18-24pt) > 正文(14-16pt)
4. **留白** — 每页只讲一个核心观点
5. **左对齐** — 文字优先左对齐，标题居中

### 杂志风原则（归藏风格）
1. **克制优于炫技** — 装饰少而精，靠字体对比和留白
2. **图片是第一公民** — 图片只裁底部，顶部和左右完整
3. **信息层级由字号定义** — 主标题 > 副标题 > lead > body > meta
4. **节奏靠封面/章节页交替** — hero 和 non-hero 交替不累眼睛

### 瑞士风原则
1. **单一锚点色** — 一份 deck 只用一个 accent 色
2. **极致字号对比** — 主标题与正文比例 ≥ 8:1
3. **无衬线统一** — 全程无衬线字体
4. **直角纯色** — 无渐变/阴影/圆角（分割线除外）
5. **网格至上** — 所有元素左对齐 + 大幅留白

## 参考文件索引

| 文件 | 使用时机 | 内容 |
|:----:|:--------:|:------|
| `references/themes.md` | Step 2 选主题 | 10 套主题色板 + 选择指南 |
| `references/layouts.md` | Step 4 写脚本 | 12 种布局辅助函数 |
| `references/typography.md` | 全流程 | 字体/字号/排版规范 |
| `references/checklist.md` | 生成后自检 | P0-P3 质量检查清单 |
| `scripts/generate-ppt.js` | 模板脚本 | 可复用的完整模板 |

> 💡 本 skill 的主题和布局设计借鉴了 **guizang-ppt-skill**（歸藏）的杂志风和瑞士风设计体系，以及 **ppt-agent-skill** 的 26 风格系统。感谢开源社区的贡献。
