const path = require("path");
const PptxGenJS = require(path.join(__dirname, "..", "node_modules", "pptxgenjs"));
const pptx = new PptxGenJS();
pptx.defineLayout({ name: "WIDE", width: 13.33, height: 7.5 });
pptx.layout = "WIDE";

// ===================== 10 套完整主题 =====================
const THEMES = {
  ink_classic: {
    name: "墨水经典", nameEn: "Ink Classic", badge: "🖋",
    bg: "F1EFEA", fg: "0A0A0B", cardBg: "E8E5DE",
    accent1: "0A0A0B", accent2: "6B5B4F", textGray: "6B5B4F",
    family: "杂志风", desc: "暖米白纸底 × 纯墨黑文字，Monocle 杂志感。通用分享、商业发布最安全的选择。",
    hex: ["F1EFEA","E8E5DE","0A0A0B","6B5B4F"],
  },
  indigo_porcelain: {
    name: "靛蓝瓷", nameEn: "Indigo Porcelain", badge: "🌊",
    bg: "F1F3F5", fg: "0A1F3D", cardBg: "E4E8EC",
    accent1: "0A1F3D", accent2: "2A4A7F", textGray: "4A6A8A",
    family: "杂志风", desc: "瓷白底色 × 深靛蓝文字，冷静理性。适合科技、研究、AI 产品发布会。",
    hex: ["F1F3F5","E4E8EC","0A1F3D","2A4A7F"],
  },
  forest_ink: {
    name: "森林墨", nameEn: "Forest Ink", badge: "🌿",
    bg: "F5F1E8", fg: "1A2E1F", cardBg: "ECE7DA",
    accent1: "1A2E1F", accent2: "3A5E3F", textGray: "5A7E5F",
    family: "杂志风", desc: "象牙底色 × 森林绿文字，沉稳有呼吸感。适合自然、可持续、文化类内容。",
    hex: ["F5F1E8","ECE7DA","1A2E1F","3A5E3F"],
  },
  kraft_paper: {
    name: "牛皮纸", nameEn: "Kraft Paper", badge: "🍂",
    bg: "EEDFC7", fg: "2A1E13", cardBg: "E0D0B6",
    accent1: "2A1E13", accent2: "5A4A3A", textGray: "7A6A5A",
    family: "杂志风", desc: "暖米底 × 深棕文字，像牛皮信封般温暖。适合怀旧、人文、历史、文学分享。",
    hex: ["EEDFC7","E0D0B6","2A1E13","5A4A3A"],
  },
  dune: {
    name: "沙丘", nameEn: "Dune", badge: "🌙",
    bg: "F0E6D2", fg: "1F1A14", cardBg: "E3D7BF",
    accent1: "1F1A14", accent2: "4A3A2A", textGray: "6A5A4A",
    family: "杂志风", desc: "沙色底 × 炭灰文字，克制高级。适合艺术、设计、创意、时尚分享。",
    hex: ["F0E6D2","E3D7BF","1F1A14","4A3A2A"],
  },
  ikb_blue: {
    name: "克莱因蓝 IKB", nameEn: "IKB Blue", badge: "🔵",
    bg: "FFFFFF", fg: "1A1A1A", cardBg: "F5F5F5",
    accent1: "002FA7", accent2: "001A5E", textGray: "666666",
    family: "瑞士风", desc: "纯白底 × 克莱因蓝锚点色。极致简洁，适合商业发布、AI 产品、方法论。",
    hex: ["FFFFFF","F5F5F5","002FA7","001A5E"],
  },
  lemon_yellow: {
    name: "柠檬黄", nameEn: "Lemon Yellow", badge: "🟡",
    bg: "FFFFFF", fg: "1A1A1A", cardBg: "F5F5F5",
    accent1: "FFD500", accent2: "CCAA00", textGray: "666666",
    family: "瑞士风", desc: "纯白底 × 柠檬黄高亮。年轻活力，适合零售、消费品、Y2K 复古。",
    hex: ["FFFFFF","F5F5F5","FFD500","CCAA00"],
  },
  lemon_green: {
    name: "柠檬绿", nameEn: "Lemon Green", badge: "🟢",
    bg: "FFFFFF", fg: "1A1A1A", cardBg: "F5F5F5",
    accent1: "C5E803", accent2: "9EBA00", textGray: "666666",
    family: "瑞士风", desc: "纯白底 × 柠檬绿高亮。清新自然，适合生态、可持续、健康、Z 世代品牌。",
    hex: ["FFFFFF","F5F5F5","C5E803","9EBA00"],
  },
  safety_orange: {
    name: "安全橙", nameEn: "Safety Orange", badge: "🟠",
    bg: "FFFFFF", fg: "1A1A1A", cardBg: "F5F5F5",
    accent1: "FF6B35", accent2: "CC552A", textGray: "666666",
    family: "瑞士风", desc: "纯白底 × 安全橙高亮。警示感、动感，适合新闻、工业、运动、活力主题。",
    hex: ["FFFFFF","F5F5F5","FF6B35","CC552A"],
  },
  dark_tech: {
    name: "深空科技", nameEn: "Dark Tech", badge: "🌌",
    bg: "0A0E27", fg: "FFFFFF", cardBg: "111534",
    accent1: "00D4FF", accent2: "7C3AED", textGray: "8892B0",
    family: "科技风", desc: "深空蓝底 × 青色荧光。赛博感十足，适合科技发布会、AI 产品演示、极客分享。",
    hex: ["0A0E27","111534","00D4FF","7C3AED"],
  },
};

function addCard(s, x, y, w, h, fill, line) {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: fill }, line: { color: line || fill, width: 0.5 }, rectRadius: 0.06 });
}

function addColorSwatch(s, x, y, hex, label) {
  s.addShape(pptx.ShapeType.rect, { x, y, w: 0.7, h: 0.7, fill: { color: hex }, rectRadius: 0.04, line: { color: "CCCCCC", width: 0.3 } });
  s.addText(label + "\n#" + hex, { x: x + 0.85, y: y, w: 2.0, h: 0.7, fontSize: 9, color: "666666", fontFace: "Arial", valign: "middle" });
}

// ===== S1: 封面 =====
const s1 = pptx.addSlide();
s1.background = { fill: "0A0E27" };
s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 4, h: 0.04, fill: { color: "00D4FF" } });
s1.addShape(pptx.ShapeType.rect, { x: 0, y: 7.46, w: 13.33, h: 0.04, fill: { color: "7C3AED" } });
s1.addText("PPT 设计风格全览", { x: 0.8, y: 2.0, w: 12, h: 1.2, fontSize: 44, color: "FFFFFF", fontFace: "Arial", bold: true });
s1.addText("10 套主题 × 3 大风格体系", { x: 0.8, y: 3.3, w: 12, h: 0.7, fontSize: 22, color: "00D4FF", fontFace: "Arial" });
s1.addShape(pptx.ShapeType.rect, { x: 0.8, y: 4.2, w: 3, h: 0.02, fill: { color: "00D4FF" } });
s1.addText("ppt-generator  ·  归藏风格灵感  ·  Swiss Design  ·  Dark Tech", { x: 0.8, y: 4.5, w: 12, h: 0.5, fontSize: 13, color: "8892B0", fontFace: "Arial" });
s1.addText("2025  ·  Vibe Coding Academy", { x: 0.8, y: 6.5, w: 5, h: 0.4, fontSize: 11, color: "8892B0", fontFace: "Arial" });

// ===== S2: 总览 =====
const s2 = pptx.addSlide();
s2.background = { fill: "F5F5F5" };
s2.addText("10 套主题一览", { x: 0.8, y: 0.4, w: 8, h: 0.7, fontSize: 28, color: "1A1A1A", fontFace: "Arial", bold: true });
s2.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.1, w: 2, h: 0.02, fill: { color: "002FA7" } });

const overview = [
  { badge: "🖋", name: "墨水经典", fam: "杂志风", hex: "F1EFEA" },
  { badge: "🌊", name: "靛蓝瓷", fam: "杂志风", hex: "F1F3F5" },
  { badge: "🌿", name: "森林墨", fam: "杂志风", hex: "F5F1E8" },
  { badge: "🍂", name: "牛皮纸", fam: "杂志风", hex: "EEDFC7" },
  { badge: "🌙", name: "沙丘", fam: "杂志风", hex: "F0E6D2" },
  { badge: "🔵", name: "克莱因蓝 IKB", fam: "瑞士风", hex: "FFFFFF" },
  { badge: "🟡", name: "柠檬黄", fam: "瑞士风", hex: "FFFFFF" },
  { badge: "🟢", name: "柠檬绿", fam: "瑞士风", hex: "FFFFFF" },
  { badge: "🟠", name: "安全橙", fam: "瑞士风", hex: "FFFFFF" },
  { badge: "🌌", name: "深空科技", fam: "科技风", hex: "0A0E27" },
];
overview.forEach((o, i) => {
  const col = i % 5, row = Math.floor(i / 5);
  const xb = 0.5 + col * 2.55, yb = 1.5 + row * 2.8;
  addCard(s2, xb, yb, 2.3, 2.4, o.hex, "DDDDDD");
  s2.addText(o.badge, { x: xb + 0.3, y: yb + 0.3, w: 1.7, h: 0.6, fontSize: 28, align: "center" });
  s2.addText(o.name, { x: xb + 0.1, y: yb + 1.0, w: 2.1, h: 0.4, fontSize: 13, color: "1A1A1A", fontFace: "Arial", bold: true, align: "center" });
  s2.addText(o.fam, { x: xb + 0.1, y: yb + 1.4, w: 2.1, h: 0.3, fontSize: 10, color: "666666", fontFace: "Arial", align: "center" });
  s2.addText("#" + o.hex, { x: xb + 0.1, y: yb + 1.75, w: 2.1, h: 0.3, fontSize: 9, color: "999999", fontFace: "Consolas", align: "center" });
});

// ===== 风格展示页 =====
const themeOrder = ["ink_classic","indigo_porcelain","forest_ink","kraft_paper","dune","ikb_blue","lemon_yellow","lemon_green","safety_orange","dark_tech"];

themeOrder.forEach((key, idx) => {
  const T = THEMES[key];
  const s = pptx.addSlide();

  // 用该主题的背景和前景
  s.background = { fill: T.bg };

  // 顶部装饰线
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 0.04, fill: { color: T.accent1 } });

  // 风格标签
  const tagColor = T.family === "科技风" ? "00D4FF" : (T.family === "瑞士风" ? T.accent1 : T.fg);
  s.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 0.3, w: 1.5, h: 0.35, fill: { color: tagColor }, rectRadius: 0.04 });
  s.addText(T.family, { x: 0.8, y: 0.3, w: 1.5, h: 0.35, fontSize: 10, color: T.family === "科技风" ? "0A0E27" : (T.family === "瑞士风" ? "FFFFFF" : T.bg), fontFace: "Arial", bold: true, align: "center", valign: "middle" });

  // 主题名
  s.addText(T.badge + "  " + T.name, { x: 2.5, y: 0.25, w: 8, h: 0.5, fontSize: 22, color: T.fg, fontFace: "Arial", bold: true });
  s.addText(T.nameEn, { x: 2.5, y: 0.7, w: 8, h: 0.3, fontSize: 11, color: T.textGray, fontFace: "Arial" });

  // 左侧：色板
  addCard(s, 0.8, 1.3, 5.5, 5.5, T.cardBg, T.accent1);
  s.addText("🎨 色彩系统", { x: 1.0, y: 1.4, w: 5.0, h: 0.4, fontSize: 14, color: T.fg, fontFace: "Arial", bold: true });
  s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.85, w: 5.0, h: 0.01, fill: { color: T.accent1, opacity: 30 } });

  // 色块展示
  const swatches = [
    { hex: T.bg, label: "背景色" },
    { hex: T.cardBg, label: "卡片色" },
    { hex: T.fg, label: "文字色" },
    { hex: T.accent1, label: "强调色" },
    { hex: T.accent2, label: "次强调" },
    { hex: T.textGray, label: "灰色文字" },
  ];
  swatches.forEach((sw, i) => {
    const row = Math.floor(i / 2), col = i % 2;
    const xb = 1.0 + col * 2.5, yb = 2.0 + row * 1.1;
    s.addShape(pptx.ShapeType.rect, { x: xb, y: yb, w: 2.1, h: 0.85, fill: { color: sw.hex }, rectRadius: 0.04, line: { color: "CCCCCC", width: 0.3 } });
    const textColor = (sw.hex === "FFFFFF" || sw.hex === "F1EFEA" || sw.hex === "F1F3F5" || sw.hex === "F5F1E8" || sw.hex === "EEDFC7" || sw.hex === "F0E6D2" || sw.hex === "E8E5DE" || sw.hex === "E4E8EC" || sw.hex === "ECE7DA" || sw.hex === "E0D0B6" || sw.hex === "E3D7BF" || sw.hex === "F5F5F5") ? T.fg : "FFFFFF";
    s.addText(sw.label, { x: xb + 0.1, y: yb + 0.1, w: 1.9, h: 0.3, fontSize: 10, color: textColor, fontFace: "Arial", bold: true });
    s.addText("#" + sw.hex, { x: xb + 0.1, y: yb + 0.45, w: 1.9, h: 0.3, fontSize: 9, color: textColor, fontFace: "Consolas" });
  });

  // 右侧：介绍 + 示例
  addCard(s, 6.8, 1.3, 5.7, 2.5, T.bg, T.accent1);
  s.addText("📖 风格介绍", { x: 7.0, y: 1.4, w: 5.3, h: 0.4, fontSize: 14, color: T.fg, fontFace: "Arial", bold: true });
  s.addShape(pptx.ShapeType.rect, { x: 7.0, y: 1.85, w: 5.3, h: 0.01, fill: { color: T.accent1, opacity: 30 } });
  s.addText(T.desc, { x: 7.0, y: 2.0, w: 5.3, h: 1.5, fontSize: 14, color: T.fg, fontFace: "Arial", lineSpacing: 22 });

  // 适合场景
  addCard(s, 6.8, 4.1, 5.7, 2.7, T.cardBg, T.accent2);
  s.addText("✅ 适合场景", { x: 7.0, y: 4.2, w: 5.3, h: 0.4, fontSize: 13, color: T.fg, fontFace: "Arial", bold: true });
  const sceneMap = {
    "墨水经典": "• 通用商业分享\n• 发布会 / 路演\n• 任何场景都安全的默认选择",
    "靛蓝瓷": "• 科技 / AI 产品发布\n• 数据分享 / 研究汇报\n• 工程师文化",
    "森林墨": "• 自然 / 可持续主题\n• 文化 / 非虚构内容\n• 户外品牌",
    "牛皮纸": "• 人文 / 历史 / 文学\n• 书评 / 生活方式\n• 独立杂志风格",
    "沙丘": "• 艺术 / 设计 / 时尚\n• 私享会 / 画廊\n• 审美优先的场合",
    "克莱因蓝 IKB": "• 商业发布 / AI 产品\n• 方法论分享\n• 通用瑞士风默认选择",
    "柠檬黄": "• 年轻品牌 / 潮流\n• 零售 / 消费品\n• Y2K / 复古",
    "柠檬绿": "• 生态 / 可持续 / 健康\n• Z 世代品牌\n• 绿色科技",
    "安全橙": "• 新闻 / 工业 / 运动\n• 活力主题\n• 警示 / 动感",
    "深空科技": "• 科技发布会\n• AI 产品演示\n• 极客 / 赛博朋克",
  };
  s.addText(sceneMap[T.name] || "", { x: 7.0, y: 4.65, w: 5.3, h: 2.0, fontSize: 12, color: T.fg, fontFace: "Arial", lineSpacing: 20 });

  // 底部装饰
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 7.46, w: 13.33, h: 0.02, fill: { color: T.accent1, opacity: 30 } });
  s.addText((idx + 3) + " / 14", { x: 12.0, y: 7.0, w: 1.0, h: 0.35, fontSize: 9, color: T.textGray, fontFace: "Arial", align: "right" });
});

// ===== S13: 风格选择指南 =====
const s13 = pptx.addSlide();
s13.background = { fill: "F5F5F5" };
s13.addText("🎯 风格选择指南", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 28, color: "1A1A1A", fontFace: "Arial", bold: true });
s13.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.1, w: 2, h: 0.02, fill: { color: "002FA7" } });

const guide = [
  { scene: "通用分享 / 商业发布 / 不知道选啥", rec: "🖋 墨水经典", why: "最安全，杂志感强" },
  { scene: "AI / 科技 / 产品发布会", rec: "🌊 靛蓝瓷 或 🌌 深空科技", why: "冷静专业 / 赛博感" },
  { scene: "数据汇报 / 工程分享", rec: "🔵 克莱因蓝 IKB", why: "瑞士风极致简洁" },
  { scene: "行业观察 / 文化 / 非虚构", rec: "🌿 森林墨 或 🍂 牛皮纸", why: "沉稳人文气息" },
  { scene: "设计 / 艺术 / 创意 / 时尚", rec: "🌙 沙丘", why: "克制高级" },
  { scene: "年轻品牌 / 消费品 / 潮流", rec: "🟡 柠檬黄 或 🟢 柠檬绿", why: "活力鲜明" },
  { scene: "新闻 / 工业 / 运动", rec: "🟠 安全橙", why: "动感警示" },
];
guide.forEach((g, i) => {
  const yb = 1.5 + i * 0.8;
  addCard(s13, 0.8, yb, 11.8, 0.65, "FFFFFF", "DDDDDD");
  s13.addText(g.scene, { x: 1.0, y: yb + 0.05, w: 4.5, h: 0.55, fontSize: 13, color: "1A1A1A", fontFace: "Arial", valign: "middle" });
  s13.addText("→", { x: 5.6, y: yb + 0.05, w: 0.5, h: 0.55, fontSize: 14, color: "002FA7", fontFace: "Arial", valign: "middle", align: "center" });
  s13.addText(g.rec, { x: 6.3, y: yb + 0.05, w: 3.5, h: 0.55, fontSize: 14, color: "002FA7", fontFace: "Arial", bold: true, valign: "middle" });
  s13.addText(g.why, { x: 10.0, y: yb + 0.05, w: 2.5, h: 0.55, fontSize: 11, color: "666666", fontFace: "Arial", valign: "middle", align: "right" });
});

s13.addShape(pptx.ShapeType.rect, { x: 0.8, y: 7.0, w: 11.8, h: 0.005, fill: { color: "002FA7", opacity: 20 } });
s13.addText("💡 原则：一份 deck 只用一套主题，不中途换色。不允许自定义 hex 值——从预设中挑选。", { x: 1.0, y: 7.1, w: 11.5, h: 0.3, fontSize: 11, color: "666666", fontFace: "Arial", align: "center" });

// ===== S14: 三大风格对比 =====
const s14 = pptx.addSlide();
s14.background = { fill: "F5F5F5" };
s14.addText("三大风格体系对比", { x: 0.8, y: 0.4, w: 10, h: 0.7, fontSize: 28, color: "1A1A1A", fontFace: "Arial", bold: true });
s14.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.1, w: 2, h: 0.02, fill: { color: "002FA7" } });

const families = [
  {
    name: "🖋 杂志风", sub: "归藏灵感 · Monocle 美学",
    bg: "F1EFEA", fg: "0A0A0B",
    items: ["衬线字体标题", "暖色纸底 + 墨色文字", "WebGL 流体背景", "克制留白，信息层级分明", "适合叙事、故事性表达"],
  },
  {
    name: "🇨🇭 瑞士风", sub: "Vignelli · 国际主义",
    bg: "FFFFFF", fg: "1A1A1A",
    items: ["无衬线统一字体", "纯白底 + 单一锚点色", "极细网格 + 点阵背景", "直角纯色，无渐变阴影", "适合事实、数据、分析"],
  },
  {
    name: "🌌 科技风", sub: "赛博 · 未来感",
    bg: "0A0E27", fg: "FFFFFF",
    items: ["无衬线 + 等宽字体", "深色底 + 荧光强调色", "青色/紫色荧光点缀", "卡片叠加 + 毛玻璃", "适合发布会、极客文化"],
  },
];
families.forEach((f, i) => {
  const xb = 0.8 + i * 4.0;
  addCard(s14, xb, 1.5, 3.7, 5.5, f.bg, "DDDDDD");
  s14.addText(f.name, { x: xb + 0.2, y: 1.7, w: 3.3, h: 0.5, fontSize: 18, color: f.fg, fontFace: "Arial", bold: true });
  s14.addText(f.sub, { x: xb + 0.2, y: 2.2, w: 3.3, h: 0.3, fontSize: 10, color: "666666", fontFace: "Arial" });
  s14.addShape(pptx.ShapeType.rect, { x: xb + 0.2, y: 2.6, w: 3.3, h: 0.01, fill: { color: "DDDDDD" } });
  f.items.forEach((item, j) => {
    s14.addText("✦ " + item, { x: xb + 0.2, y: 2.8 + j * 0.65, w: 3.3, h: 0.55, fontSize: 12, color: f.fg, fontFace: "Arial" });
  });
});

// ===== S15: 结束页 =====
const s15 = pptx.addSlide();
s15.background = { fill: "0A0E27" };
s15.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 0.04, fill: { color: "00D4FF" } });
s15.addShape(pptx.ShapeType.rect, { x: 0, y: 7.46, w: 13.33, h: 0.04, fill: { color: "7C3AED" } });
s15.addText("谢谢观看", { x: 0.8, y: 2.2, w: 12, h: 1.0, fontSize: 44, color: "FFFFFF", fontFace: "Arial", bold: true, align: "center" });
s15.addShape(pptx.ShapeType.rect, { x: 5.5, y: 3.3, w: 2.5, h: 0.03, fill: { color: "00D4FF" } });
s15.addText("PPT Generator  ·  10 Themes  ·  3 Styles", { x: 0.8, y: 3.6, w: 12, h: 0.5, fontSize: 16, color: "8892B0", fontFace: "Arial", align: "center" });
s15.addText("从归藏杂志风到瑞士国际主义，再到深空科技", { x: 0.8, y: 4.3, w: 12, h: 0.5, fontSize: 14, color: "00D4FF", fontFace: "Arial", bold: true, align: "center" });
s15.addText("Vibe Coding Academy  ·  2025", { x: 0.8, y: 6.5, w: 12, h: 0.4, fontSize: 12, color: "8892B0", fontFace: "Arial", align: "center" });

// Save
const outPath = path.join(__dirname, "..", "PPT_设计风格全览.pptx");
pptx.writeFile({ fileName: outPath })
  .then(() => console.log("Done: " + outPath + "  (" + pptx.slides.length + " slides)"))
  .catch(e => console.error(e));
