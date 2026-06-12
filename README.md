# PPT Generator Skill

> 一款为 AI Agent（Codex / Claude Code）设计的 PowerPoint 演示文稿生成技能。
> 拥有 **10 套高端设计主题**，覆盖杂志风、瑞士国际主义风、深空科技风三大体系。

## ✨ 特性

- 🎨 **10 套预置主题** — 从归藏杂志风到瑞士国际主义，开箱即用
- 📐 **12 种页面布局** — 封面、章节幕封、数据大字报、Bento Grid 等
- 🖋 **可编辑 PPTX** — 原生 PowerPoint 格式，直接打开修改
- 🚀 **零配置运行** — 安装后即可使用，无需 API Key
- 📦 **单文件 HTML 可选** — 也可配合 guizang-ppt-skill 输出网页版

## 🎯 10 套主题一览

| 风格体系 | 主题 | 适合场景 |
|:-------:|:----|:---------|
| 🖋 **杂志风** | 墨水经典、靛蓝瓷、森林墨、牛皮纸、沙丘 | 人文分享、行业观察、商业发布 |
| 🇨🇭 **瑞士风** | 克莱因蓝 IKB、柠檬黄、柠檬绿、安全橙 | 科技产品、数据汇报、设计分享 |
| 🌌 **科技风** | 深空科技 | 科技发布会、AI 产品演示 |

## 🚀 快速开始

### 安装

```bash
# 克隆到 Codex skills 目录
git clone https://github.com/Doris-psk/ppt-generator-skill.git
# 或者用 npx skills
npx skills add https://github.com/Doris-psk/ppt-generator-skill --skill ppt-generator
```

### 使用

对 AI Agent 说：

> "用 ppt-generator 做个 PPT，靛蓝瓷主题，15页，关于 AI 落地的内容"

### 生成示例

```bash
cd scripts
npm install
node generate-ppt.js    # 生成示例演示文稿
```

## 📂 目录结构

```
ppt-generator-skill/
├── SKILL.md              ← 技能主指令
├── agents/openai.yaml    ← UI 元数据
├── references/
│   ├── themes.md         ← 10 套主题色板
│   ├── layouts.md        ← 12 种布局函数
│   ├── typography.md     ← 排版规范
│   └── checklist.md      ← 质量检查清单
├── scripts/
│   ├── generate-ppt.js           ← 完整模板脚本
│   └── generate-style-showcase.js ← 风格全览生成脚本
├── PPT_设计风格全览.pptx         ← 10 套风格展示
├── FDE_案例展示.pptx             ← 实际案例
└── package.json
```

## 🎨 灵感来源

- **guizang-ppt-skill**（歸藏）— 杂志风和瑞士风设计体系
- **ppt-agent-skill**（Akxan）— 26 风格系统与工作流
- **Massimo Vignelli** — 瑞士国际主义设计语言

## 📄 License

MIT
