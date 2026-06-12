// ============================================================
// PPT Generator — 完整模板脚本
// 10 套主题 × 12 种布局 = 无限可能
// 使用方式: node scripts/generate-ppt.js
// ============================================================
const path = require("path");
const PptxGenJS = require(path.join(__dirname, "..", "node_modules", "pptxgenjs"));
const pptx = new PptxGenJS();

// ===================== 布局设置 =====================
pptx.defineLayout({ name: "WIDE", width: 13.33, height: 7.5 });
pptx.layout = "WIDE";

// ===================== 10 套主题 =====================
// 完整说明见 references/themes.md
const THEMES = {
  // --- 杂志风 ---
  ink_classic: {
    name: "墨水经典", bg: "F1EFEA", fg: "0A0A0B",
    cardBg: "E8E5DE", accent1: "0A0A0B",
    accent2: "6B5B4F", textGray: "6B5B4F",
  },
  indigo_porcelain: {
    name: "靛蓝瓷", bg: "F1F3F5", fg: "0A1F3D",
    cardBg: "E4E8EC", accent1: "0A1F3D",
    accent2: "2A4A7F", textGray: "4A6A8A",
  },
  forest_ink: {
    name: "森林墨", bg: "F5F1E8", fg: "1A2E1F",
    cardBg: "ECE7DA", accent1: "1A2E1F",
    accent2: "3A5E3F", textGray: "5A7E5F",
  },
  kraft_paper: {
    name: "牛皮纸", bg: "EEDFC7", fg: "2A1E13",
    cardBg: "E0D0B6", accent1: "2A1E13",
    accent2: "5A4A3A", textGray: "7A6A5A",
  },
  dune: {
    name: "沙丘", bg: "F0E6D2", fg: "1F1A14",
    cardBg: "E3D7BF", accent1: "1F1A14",
    accent2: "4A3A2A", textGray: "6A5A4A",
  },
  // --- 瑞士风 ---
  ikb_blue: {
    name: "克莱因蓝 IKB", bg: "FFFFFF", fg: "1A1A1A",
    cardBg: "F5F5F5", accent1: "002FA7",
    accent2: "001A5E", textGray: "666666",
  },
  lemon_yellow: {
    name: "柠檬黄", bg: "FFFFFF", fg: "1A1A1A",
    cardBg: "F5F5F5", accent1: "FFD500",
    accent2: "CCAA00", textGray: "666666",
  },
  lemon_green: {
    name: "柠檬绿", bg: "FFFFFF", fg: "1A1A1A",
    cardBg: "F5F5F5", accent1: "C5E803",
    accent2: "9EBA00", textGray: "666666",
  },
  safety_orange: {
    name: "安全橙", bg: "FFFFFF", fg: "1A1A1A",
    cardBg: "F5F5F5", accent1: "FF6B35",
    accent2: "CC552A", textGray: "666666",
  },
  // --- 深色科技 ---
  dark_tech: {
    name: "深空科技", bg: "0A0E27", fg: "FFFFFF",
    cardBg: "111534", accent1: "00D4FF",
    accent2: "7C3AED", textGray: "8892B0",
  },
};

// ===================== 12 种布局辅助函数 =====================
// 完整说明见 references/layouts.md

/** L01: 封面页 */
function addCover(slide, T, title, subtitle, meta) {
  slide.background = { fill: T.bg };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 4, h: 0.04, fill: { color: T.accent1 } });
  slide.addText(title, { x: 0.8, y: 2.0, w: 11.5, h: 1.5, fontSize: 44, color: T.fg, fontFace: "Arial", bold: true });
  if (subtitle) slide.addText(subtitle, { x: 0.8, y: 3.6, w: 11.5, h: 0.7, fontSize: 22, color: T.accent1, fontFace: "Arial" });
  if (meta) slide.addText(meta, { x: 0.8, y: 6.5, w: 5, h: 0.4, fontSize: 11, color: T.textGray, fontFace: "Arial" });
}

/** L02: 章节幕封 */
function addSectionSlide(slide, T, sectionNum, title, subtitle) {
  slide.background = { fill: T.bg };
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 2.0, w: 0.04, h: 2.5, fill: { color: T.accent1 } });
  slide.addText(sectionNum, { x: 1.2, y: 2.0, w: 2, h: 0.5, fontSize: 14, color: T.accent1, fontFace: "Arial", bold: true });
  slide.addText(title, { x: 1.2, y: 2.8, w: 10, h: 1.0, fontSize: 36, color: T.fg, fontFace: "Arial", bold: true });
  if (subtitle) slide.addText(subtitle, { x: 1.2, y: 4.0, w: 10, h: 0.5, fontSize: 16, color: T.textGray, fontFace: "Arial" });
}

/** L03: 标准内容页 */
function addContentSlide(slide, T, title, items) {
  slide.background = { fill: T.bg };
  slide.addText(title, { x: 0.8, y: 1.0, w: 11.5, h: 0.8, fontSize: 28, color: T.fg, fontFace: "Arial", bold: true });
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.9, w: 3, h: 0.02, fill: { color: T.accent1 } });
  items.forEach((item, i) => {
    slide.addText("\u25B8  " + item, { x: 1.0, y: 2.4 + i * 0.7, w: 11, h: 0.6, fontSize: 16, color: T.fg, fontFace: "Arial" });
  });
}

/** L04: 数据大字报 */
function addDataHero(slide, T, number, label, note) {
  slide.background = { fill: T.bg };
  slide.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.5, w: 11.3, h: 4.5, fill: { color: T.cardBg }, rectRadius: 0.08 });
  slide.addText(number, { x: 1.0, y: 2.0, w: 11.3, h: 2.5, fontSize: 72, color: T.accent1, fontFace: "Arial", bold: true, align: "center", valign: "middle" });
  slide.addText(label, { x: 1.0, y: 4.2, w: 11.3, h: 0.6, fontSize: 18, color: T.fg, fontFace: "Arial", align: "center" });
  if (note) slide.addText(note, { x: 1.0, y: 4.9, w: 11.3, h: 0.4, fontSize: 12, color: T.textGray, fontFace: "Arial", align: "center" });
}

/** L05: 左文右图 */
function addImageText(slide, T, title, body, imgLabel) {
  slide.background = { fill: T.bg };
  slide.addText(title, { x: 0.8, y: 1.5, w: 5.5, h: 0.8, fontSize: 28, color: T.fg, fontFace: "Arial", bold: true });
  slide.addText(body, { x: 0.8, y: 2.6, w: 5.5, h: 3.0, fontSize: 16, color: T.fg, fontFace: "Arial" });
  slide.addShape(pptx.ShapeType.rect, { x: 7.0, y: 1.5, w: 5.3, h: 5.0, fill: { color: T.cardBg }, line: { color: T.accent1, width: 0.5 }, rectRadius: 0.04 });
  slide.addText(imgLabel || "[图片占位]", { x: 7.0, y: 3.5, w: 5.3, h: 1.0, fontSize: 14, color: T.textGray, fontFace: "Arial", align: "center", valign: "middle" });
}

/** L06: 图片网格 */
function addImageGrid(slide, T, title, images) {
  slide.background = { fill: T.bg };
  slide.addText(title, { x: 0.8, y: 0.5, w: 11.5, h: 0.7, fontSize: 24, color: T.fg, fontFace: "Arial", bold: true });
  const cols = 3;
  images.forEach((img, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const xb = 0.8 + col * 4.0, yb = 1.5 + row * 2.8;
    slide.addShape(pptx.ShapeType.rect, { x: xb, y: yb, w: 3.6, h: 2.2, fill: { color: T.cardBg }, line: { color: T.accent1, width: 0.5 }, rectRadius: 0.04 });
    slide.addText(img.label || "[图]", { x: xb, y: yb + 0.3, w: 3.6, h: 1.2, fontSize: 14, color: T.fg, fontFace: "Arial", align: "center", valign: "middle" });
    if (img.desc) slide.addText(img.desc, { x: xb + 0.2, y: yb + 1.6, w: 3.2, h: 0.5, fontSize: 10, color: T.textGray, fontFace: "Arial", align: "center" });
  });
}

/** L07: 对比页 */
function addComparison(slide, T, leftTitle, leftBody, rightTitle, rightBody) {
  slide.background = { fill: T.bg };
  slide.addText(leftTitle, { x: 0.8, y: 1.0, w: 5.5, h: 0.7, fontSize: 24, color: T.fg, fontFace: "Arial", bold: true });
  slide.addText(rightTitle, { x: 7.0, y: 1.0, w: 5.5, h: 0.7, fontSize: 24, color: T.accent1, fontFace: "Arial", bold: true, align: "right" });
  slide.addShape(pptx.ShapeType.rect, { x: 6.6, y: 1.2, w: 0.02, h: 5.0, fill: { color: T.textGray } });
  slide.addText(leftBody, { x: 0.8, y: 2.0, w: 5.5, h: 4.0, fontSize: 15, color: T.fg, fontFace: "Arial" });
  slide.addText(rightBody, { x: 7.0, y: 2.0, w: 5.5, h: 4.0, fontSize: 15, color: T.fg, fontFace: "Arial" });
}

/** L08: 三段式卡片 */
function addThreeCards(slide, T, title, cards) {
  slide.background = { fill: T.bg };
  slide.addText(title, { x: 0.8, y: 0.5, w: 11.5, h: 0.8, fontSize: 28, color: T.fg, fontFace: "Arial", bold: true });
  cards.forEach((card, i) => {
    const xb = 0.8 + i * 4.0;
    slide.addShape(pptx.ShapeType.rect, { x: xb, y: 1.8, w: 3.6, h: 4.5, fill: { color: T.cardBg }, line: { color: T.accent1, width: 0.5 }, rectRadius: 0.08 });
    slide.addText(card.title, { x: xb + 0.3, y: 2.2, w: 3.0, h: 0.6, fontSize: 20, color: T.fg, fontFace: "Arial", bold: true });
    slide.addText(card.desc, { x: xb + 0.3, y: 3.0, w: 3.0, h: 3.0, fontSize: 14, color: T.textGray, fontFace: "Arial" });
  });
}

/** L09: 时间线 */
function addTimeline(slide, T, title, steps) {
  slide.background = { fill: T.bg };
  slide.addText(title, { x: 0.8, y: 0.5, w: 11.5, h: 0.8, fontSize: 28, color: T.fg, fontFace: "Arial", bold: true });
  const lineY = 4.0;
  slide.addShape(pptx.ShapeType.rect, { x: 1.0, y: lineY, w: 11.3, h: 0.02, fill: { color: T.accent1 } });
  steps.forEach((step, i) => {
    const xb = 1.0 + (i * 11.3) / Math.max(steps.length - 1, 1);
    slide.addShape(pptx.ShapeType.ellipse, { x: xb - 0.1, y: lineY - 0.1, w: 0.22, h: 0.22, fill: { color: T.accent1 } });
    slide.addText(step.num || "", { x: xb - 0.8, y: lineY - 1.5, w: 1.6, h: 0.4, fontSize: 12, color: T.accent1, fontFace: "Arial", bold: true, align: "center" });
    slide.addText(step.title, { x: xb - 0.8, y: lineY - 1.0, w: 1.6, h: 0.8, fontSize: 13, color: T.fg, fontFace: "Arial", align: "center" });
    if (step.desc) slide.addText(step.desc || "", { x: xb - 0.8, y: lineY + 0.4, w: 1.6, h: 1.5, fontSize: 11, color: T.textGray, fontFace: "Arial", align: "center" });
  });
}

/** L10: 引用页 */
function addQuoteSlide(slide, T, quote, author) {
  slide.background = { fill: T.bg };
  slide.addText("\u201C", { x: 0.8, y: 1.0, w: 1.5, h: 1.5, fontSize: 72, color: T.accent1, fontFace: "Arial", bold: true });
  slide.addText(quote, { x: 1.5, y: 2.0, w: 10.5, h: 3.0, fontSize: 24, color: T.fg, fontFace: "Arial", italic: true });
  slide.addText("\u2014 " + (author || ""), { x: 1.5, y: 5.2, w: 10.5, h: 0.5, fontSize: 14, color: T.textGray, fontFace: "Arial", align: "right" });
}

/** L11: Bento Grid */
function addBentoGrid(slide, T, title, items) {
  slide.background = { fill: T.bg };
  slide.addText(title, { x: 0.8, y: 0.5, w: 11.5, h: 0.7, fontSize: 26, color: T.fg, fontFace: "Arial", bold: true });
  const positions = [
    { x: 0.8, y: 1.6, w: 5.5, h: 2.8 },
    { x: 6.8, y: 1.6, w: 5.7, h: 1.3 },
    { x: 6.8, y: 3.1, w: 2.7, h: 1.3 },
    { x: 9.8, y: 3.1, w: 2.7, h: 1.3 },
  ];
  items.slice(0, 4).forEach((item, i) => {
    const p = positions[i];
    slide.addShape(pptx.ShapeType.rect, { x: p.x, y: p.y, w: p.w, h: p.h, fill: { color: T.cardBg }, line: { color: T.accent1, width: 0.5 }, rectRadius: 0.06 });
    slide.addText(item.title, { x: p.x + 0.3, y: p.y + 0.3, w: p.w - 0.6, h: 0.5, fontSize: i === 0 ? 22 : 16, color: T.fg, fontFace: "Arial", bold: true });
    if (item.desc) slide.addText(item.desc, { x: p.x + 0.3, y: p.y + 1.0, w: p.w - 0.6, h: p.h - 1.5, fontSize: i === 0 ? 15 : 12, color: T.textGray, fontFace: "Arial" });
  });
}

/** L12: 结束页 */
function addClosing(slide, T, thanks, subtitle) {
  slide.background = { fill: T.bg };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 0.04, fill: { color: T.accent1 } });
  slide.addText(thanks || "谢谢大家", { x: 0.8, y: 2.5, w: 12, h: 1.2, fontSize: 48, color: T.fg, fontFace: "Arial", bold: true, align: "center" });
  if (subtitle) {
    slide.addShape(pptx.ShapeType.rect, { x: 5.5, y: 3.8, w: 2.5, h: 0.03, fill: { color: T.accent1 } });
    slide.addText(subtitle, { x: 0.8, y: 4.0, w: 12, h: 0.6, fontSize: 18, color: T.textGray, fontFace: "Arial", align: "center" });
  }
}

// ===================== 示例：生成一份演示文稿 =====================
// 使用墨水经典（杂志风）主题
const T = THEMES.ink_classic;

// 1. 封面
const s1 = pptx.addSlide();
addCover(s1, T, "演示文稿标题", "副标题 / 一句话说明", "作者名  ·  2025年6月");

// 2. 章节页
const s2 = pptx.addSlide();
addSectionSlide(s2, T, "01", "第一部分", "章节副标题说明");

// 3. 内容页
const s3 = pptx.addSlide();
addContentSlide(s3, T, "核心要点", ["要点一：这是第一个关键信息", "要点二：这是第二个关键信息", "要点三：这是第三个关键信息"]);

// 4. 数据大字报
const s4 = pptx.addSlide();
addDataHero(s4, T, "86%", "用户满意度", "基于 2025 年 Q1 调研数据");

// 5. 三段卡片
const s5 = pptx.addSlide();
addThreeCards(s5, T, "三大优势", [
  { title: "高效", desc: "AI 驱动的工作流程，大幅提升效率" },
  { title: "智能", desc: "深度学习算法，精准理解用户需求" },
  { title: "可靠", desc: "企业级稳定性，99.9% 可用性保障" },
]);

// 6. 对比页
const s6 = pptx.addSlide();
addComparison(s6, T,
  "传统方式", "• 手动操作\n• 耗时数小时\n• 容易出错\n• 依赖专业经验",
  "AI 方式", "• 自动化处理\n• 几分钟完成\n• 精准度高\n• 零门槛使用"
);

// 7. 引用
const s7 = pptx.addSlide();
addQuoteSlide(s7, T, "AI 不会取代人类，但会用 AI 的人会取代不会用 AI 的人。", "行业观察");

// 8. 时间线
const s8 = pptx.addSlide();
addTimeline(s8, T, "发展历程", [
  { num: "2023", title: "概念验证", desc: "初步探索" },
  { num: "2024", title: "产品落地", desc: "正式发布" },
  { num: "2025", title: "规模扩展", desc: "全面推广" },
  { num: "2026", title: "生态构建", desc: "平台化" },
]);

// 9. 结束页
const s9 = pptx.addSlide();
addClosing(s9, T, "谢谢大家", "Vibe Coding Academy  ·  2025");

// ===================== 保存 =====================
const outPath = path.join(__dirname, "..", "presentation.pptx");
pptx.writeFile({ fileName: outPath })
  .then(() => console.log("PPT generated: " + outPath + "  (" + pptx.slides.length + " slides)"))
  .catch(e => console.error("Error:", e));
