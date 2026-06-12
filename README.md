<div align="center">

# 🎨 Doris-PPT-Skill

### 一款为 AI Agent 设计的专业 PowerPoint 演示文稿生成技能

**10 套高端设计主题 · 12 种页面布局 · 开箱即用的编辑体验**

[![GitHub Stars](https://img.shields.io/github/stars/Doris-psk/Doris-PPT-skill?style=flat-square&logo=github)](https://github.com/Doris-psk/Doris-PPT-skill/stargazers)
[![License](https://img.shields.io/badge/License-MIT-00D4FF?style=flat-square)](LICENSE)
[![Skill](https://img.shields.io/badge/Skill-Agent-7C3AED?style=flat-square)](https://github.com/Doris-psk/Doris-PPT-skill)
[![PowerPoint](https://img.shields.io/badge/Output-PPTX-0A0E27?style=flat-square)](#)
[![Codex](https://img.shields.io/badge/Codex-Ready-00D4FF?style=flat-square)](#)
[![Claude Code](https://img.shields.io/badge/Claude%20Code-Ready-7C3AED?style=flat-square)](#)
[![Feishu](https://img.shields.io/badge/Feishu%E9%A3%9E%E4%B9%A6-%E7%94%B1%E6%AD%A4%E5%BC%80%E5%A7%8B-002FA7?style=flat-square)](https://waytoagi.feishu.cn/wiki/SoKBwndUFiFhqhkoESJcRtZlnde?from=from_copylink)

---

[✨ 特性](#-特性) • [🎨 主题预览](#-10-套主题一览) • [🚀 快速开始](#-快速开始) • [📂 目录结构](#-目录结构) • [🎯 使用示例](#-使用示例) • [💡 灵感来源](#-灵感来源)

</div>

---

## ✨ 特性

- 🎨 **10 套预置高端主题** — 灵感来自归藏杂志风 + 瑞士国际主义 + 深空科技，每一套都经过精心调色
- 📐 **12 种页面布局** — 封面、章节幕封、数据大字报、Bento Grid、对比页、时间线…覆盖所有演示场景
- 🖋 **原生可编辑 PPTX** — 直接输出 `.pptx` 格式，PowerPoint / WPS 打开即可修改
- 🚀 **零配置运行** — 无需 API Key，无需注册，`node script.js` 一步生成
- 🧩 **模块化辅助函数** — 布局函数即调即用，像搭积木一样构建演示文稿
- 📋 **质量检查清单** — 内置 P0-P3 分级质检，保证每份输出都有专业水准

---

## 🎨 10 套主题一览

### 🖋 杂志风 — 归藏灵感 · Monocle 美学

| 主题 | 色系 | 适合场景 |
|:----:|:----:|:---------|
| 🖋 **墨水经典** | 暖米白 + 纯墨黑 | 通用分享、商业发布、任何安全选择 |
| 🌊 **靛蓝瓷** | 瓷白 + 深靛蓝 | 科技、AI 产品、研究汇报 |
| 🌿 **森林墨** | 象牙 + 森林绿 | 自然、文化、非虚构内容 |
| 🍂 **牛皮纸** | 暖米 + 深棕 | 人文、历史、文学、书评 |
| 🌙 **沙丘** | 沙色 + 炭灰 | 艺术、设计、时尚、画廊 |

### 🇨🇭 瑞士风 — Vignelli · 国际主义

| 主题 | 锚点色 | 适合场景 |
|:----:|::------:|:---------|
| 🔵 **克莱因蓝 IKB** | `#002FA7` | 商业发布、AI 产品、方法论 |
| 🟡 **柠檬黄** | `#FFD500` | 年轻品牌、零售、消费品 |
| 🟢 **柠檬绿** | `#C5E803` | 生态、健康、Z 世代品牌 |
| 🟠 **安全橙** | `#FF6B35` | 新闻、工业、运动、活力主题 |

### 🌌 科技风 — 赛博 · 未来感

| 主题 | 色系 | 适合场景 |
|:----:|:----:|:---------|
| 🌌 **深空科技** | 深蓝 + 青色荧光 | 科技发布会、AI 产品演示、极客分享 |

> 📥 [**下载风格全览 PPT**](PPT_Design_Showcase.pptx) — 查看所有主题的实际效果

---

## 🚀 快速开始

### 安装

```bash
# 方式一：克隆仓库
git clone https://github.com/Doris-psk/Doris-PPT-skill.git
cd Doris-PPT-skill
npm install

# 方式二：使用 npx skills（适用于 Codex / Claude Code）
npx skills add https://github.com/Doris-psk/Doris-PPT-skill --skill Doris-PPT-skill
```

### 使用

对 AI Agent 说一句话：

> **"用 Doris-PPT-skill 做个 PPT，靛蓝瓷主题，15页，关于 AI 落地的话题"**

### 手动生成

```bash
cd scripts
node generate-ppt.js
```

就会在当前目录生成 `presentation.pptx`，直接打开即可编辑。

---

## 🎯 使用示例

### 1. 基础用法 — 一句话生成

```
用户: 帮我做个关于 FDE 模式的 PPT，靛蓝瓷主题，15页
Agent: 好的，读取 meeting notes → 规划大纲 → 生成脚本 → 输出 PPTX
```

### 2. 高级用法 — 自定义内容

```js
const T = THEMES.indigo_porcelain;  // 选主题

const s1 = pptx.addSlide();
addCover(s1, T, "标题", "副标题", "作者 · 日期");

const s2 = pptx.addSlide();
addSectionSlide(s2, T, "01", "第一部分", "说明");

const s3 = pptx.addSlide();
addThreeCards(s3, T, "三大优势", [
  { title: "高效", desc: "AI 驱动的工作流程" },
  { title: "智能", desc: "深度学习算法" },
  { title: "可靠", desc: "企业级稳定性" },
]);

pptx.writeFile({ fileName: "output.pptx" });
```

### 3. 成果展示

| 风格全览 | 实战案例 |
|:-------:|:--------:|
| [![PPT_Design_Showcase](https://img.shields.io/badge/📥-Download_PPT_Design_Showcase-0A1F3D?style=for-the-badge)](PPT_Design_Showcase.pptx) | [![FDE_Case_Study](https://img.shields.io/badge/📥-Download_FDE_Case_Study-0A1F3D?style=for-the-badge)](FDE_Case_Study.pptx) |

---

## 📂 目录结构

```
Doris-PPT-skill/
├── SKILL.md                      # 技能核心指令（Codex / Claude Code 使用）
├── README.md                     # 本文件
├── package.json                  # Node.js 配置
├── .gitignore
├── agents/
│   └── openai.yaml               # UI 元数据
├── references/
│   ├── themes.md                 # 10 套主题色板（含 HEX 值）
│   ├── layouts.md                # 12 种布局辅助函数
│   ├── typography.md             # 排版规范
│   └── checklist.md              # 质量检查清单
├── scripts/
│   ├── generate-ppt.js           # 完整模板脚本
│   └── generate-style-showcase.js # 风格全览生成脚本
├── PPT_Design_Showcase.pptx      # 🎨 10 套主题风格展示
└── FDE_Case_Study.pptx           # 📊 实战案例展示
```

---

## 💡 灵感来源

| 项目 | 贡献 |
|:----|:-----|
| [guizang-ppt-skill](https://github.com/op7418/guizang-ppt-skill)（歸藏） | 杂志风和瑞士风设计体系、主题色板 |
| [ppt-agent-skill](https://github.com/Akxan/ppt-agent-skill) | 26 风格系统、调研-策划-设计分离的工作流 |
| Massimo Vignelli · 瑞士国际主义 | 网格系统、极致简约的设计哲学 |
| *Monocle* 杂志 | 克制优雅的 editorial 排版 |

---

## 🤝 贡献

欢迎 Issue 和 PR！

- 🐛 发现 Bug？[开 Issue](https://github.com/Doris-psk/Doris-PPT-skill/issues/new)
- 💡 有新主题想法？欢迎提交 PR
- 📖 完善文档或添加新布局

---

<div align="center">

**如果这个项目对你有帮助，欢迎 ⭐ Star 支持！**

由 [Doris-psk](https://github.com/Doris-psk) ❤️ 制作

</div>
