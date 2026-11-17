---
name: ppt-studio
description: 用「内容 + 风格 + 页数」直接生成可编辑的 .pptx。内置苹果官网风、通用简洁风、电商/淘宝风、行政/公文风四种样式包，风格与版式解耦，新增风格只需加一个 JSON。使用前必须先向用户提问确认内容、风格、页数三项，不可自行假设。
metadata:
  tags: ppt, pptx, presentation, deck, 幻灯片, 汇报, 封面, 排版, style-pack
---

# PPT Studio

把一份内容做成一整套 PPT：**样式包负责长相，版式原型负责结构，内容 spec 负责文字**。
三者解耦 —— 换风格只改 JSON，换内容只改 spec，换结构只加一个函数。

## 什么时候用

用户说「帮我做一份 PPT / 演示文稿 / 汇报材料 / 封面」等表述时启用。

## ⛔ 第 0 步：必须先提问（硬性规则，不可跳过）

**在写下任何内容、创建任何文件、运行任何命令之前，必须先调用 `ask_user` 提问。**

把这三件事问清楚，缺一件都不许开工：

| 输入 | 要问什么 | 绝不可以这样做 |
|---|---|---|
| **内容 / 主题** | 做什么主题？可以贴文字、给链接、说个题目 | ❌ 不许自己从仓库、README、历史对话里挑主题 |
| **风格** | 苹果风 / 通用简洁 / 电商 / 公文 / 其他描述 | ❌ 不许默默用默认 `free` |
| **页数** | 具体数字，或「你按内容估」 | ❌ 不许自己估一个页数就开做 |

### 措辞含糊也算没给

以下这些话**都算「未指定」**，必须提问，不能当成授权：

- 「帮我做个 PPT」「测试一下」「随便做一个」「你看着办」
- 「用我的项目做个介绍」（**没说风格和页数，也算缺**）
- 「做个汇报」（只有场景，没有主题/风格/页数）

### 允许跳过的唯一情况

用户在**同一句话里明确给出了全部三项**（主题 + 风格 + 页数），例如
「用 DeepSeek 发展史、苹果风、10 页」——此时三项齐全，直接开工，不必再问。

否则：**先问，问完再动手**。用户已经给了的部分不要重复问，只问缺的。

### 提问模板

```python
ask_user(
  message="要做一份 PPT，先确认三件事——",
  requestedSchema={"properties": {
    "content": {"type": "string",  "title": "内容 / 主题",
                "description": "可以贴文字、给链接、说个题目"},
    "style":   {"type": "string",  "title": "风格",
                "enum": ["free 通用简洁（默认）", "apple 苹果官网风",
                         "taobao 电商/淘宝风", "gov 行政/公文风",
                         "先不确定，你看内容推荐", "其他（我描述一下）"]},
    "pages":   {"type": "string",  "title": "页数",
                "enum": ["5 页", "8 页", "10 页", "15 页", "你按内容估"]}
  }, "required": ["content"]}
)
```

用户选「先不确定」→ 才可以给推荐；用户选「你按内容估」→ 才可以自己定页数。
**这两个都是用户明确授权后才成立的例外，不能拿来做默认值。**

## 输入（三件套）

| 输入 | 来源 |
|---|---|
| **内容** | 用户在提问中的回答，或用户提供的文字/链接 |
| **风格** | 用户选择的样式包，或用户对风格的描述映射 |
| **页数** | 用户给的数字，或用户在提问中明确授权「你估」 |

风格 ↔ 说法对照：

| 用户可能说 | 用 | 说明 |
|---|---|---|
| 苹果风、发布会、极简、高级感 | `apple` | 大留白、黑白灰、超大字号、深色页环境光 |
| 随便、通用、简洁、现代、默认 | `free` | 中性现代排版，最稳 |
| 淘宝、电商、促销、种草、有冲击力 | `taobao` | 橙红、贴纸卡片、粗描边硬投影 |
| 行政、公文、汇报、正式、国企/政府 | `gov` | 深蓝 + 红顶条、宋体标题、直角、表格化 |

```bash
python <skill>/scripts/build_deck.py --list        # 看当前所有样式与版式
```

## 工作流（三件套确认之后）

> 前置条件：第 0 步的三项已经问清。没问清不许进入这里。

### 1. 定骨架

固定首尾，中间按页数分配：

```
cover（1） + 内容页（N-2） + closing（1）
```

一页 = 一个观点。内容点数量 X 与页数 N 的关系：

- `X == N-2`：一一对应，最理想
- `X > N-2`：合并同类项 → 用 `table`（对照）、`lanes`（2×2）、`cards`（并列）
- `X < N-2`：拆 —— 把最重的点拆成 `stats`（数据）＋ `statement`（结论），或加过渡页
- 页数被指定得很小（≤4）：`cover` → `stats`/`lanes` 承载全部要点 → `closing`

### 2. 选版式

| 内容形态 | `type` | 说明 |
|---|---|---|
| 开场 / 主题 | `cover` | eyebrow + 大字标题 + 副标题 + 页脚 |
| 一句判断、结论先行 | `statement` | 大标题 + 两行以内正文 |
| 一个核心数字 | `bigstat` | 巨字数字 + 注释 + 右侧 3 条清单 |
| 2–4 个并列数据 | `stats` | 横排数字，竖线分隔 |
| 2–4 个并列要点 | `cards` | 卡片组，可带 tag |
| 4 条主线 / 并列能力 | `lanes` | 2×2 卡片矩阵 |
| 演进、阶段、路线图 | `timeline` | 8 个节点以内 |
| 文字较多（背景/依据） | `bullets` | 左标题 + 右条目，最多 3 条 |
| 横向对比 | `table` | 只用横向发丝线；`bold_cols` 强调列 |
| 收尾 / 行动号召 | `closing` | 居中或左对齐大字 + 链接 |

### 3. 写 spec

照 `examples/spec.example.json` 写 `deck.json`：`slides[]` 每项一个 `type`。
字段见下方「spec 速查」。**文案要按风格包的 `voice` 约束来写**（读一下 `styles/<style>.json` 的 `voice` 字段）。

### 4. 生成 + 自检

```bash
python <skill>/scripts/build_deck.py \
    --spec deck.json --style apple -o ~/Desktop/主题_风格.pptx --preview
```

`--preview` 会在同名 `_preview/` 目录里出逐页 PNG 和 `_sheet.png` 联系表。
**必须打开 `_sheet.png` 逐页看**，跑下面的检查清单，有问题就改 spec 重跑（不要手改 pptx）。

### 5. 交付

- 产物：`.pptx`（可编辑），需要时附 PDF 和 `_sheet.png`
- 命名：`主题_风格.pptx`
- 回话里说明：几页、用了哪些版式、哪些文案是补的、下一步能怎么调

## 自检清单（每次都要过）

开工前：

- [ ] **问过了吗**：内容 / 风格 / 页数是否都来自用户的明确答复？（不是自己推的）

出片后：
- [ ] **压字**：大标题与副标题/数据之间有没有重叠？（`disp_bottom` 已处理，但视觉上仍要确认）
- [ ] **孤字**：有没有一行只剩 1–2 个字？（缩短文案或换版式）
- [ ] **溢出**：卡片内文字有没有出框？表格列宽够不够？
- [ ] **字重**：中文是否过粗？（样式包 `weight_mode` 应为 `split`）
- [ ] **节奏**：连续两页同底色就换一页 `tone`（`dark`/`light`/`tint`）
- [ ] **空格**：中文与数字/英文之间留一个半角空格（`671B 参数` 而不是 `671B参数`）
- [ ] **信息量**：每页 ≤ 1 个观点、≤ 4 个并列项、正文 ≤ 3 行

## spec 速查

```jsonc
{
  "title": "主题",
  "style": "apple",              // 也可以用 --style 覆盖
  "width": 13.3333, "height": 7.5,   // 16:9；要 4:3 就 10 x 7.5
  "slides": [
    { "type": "cover", "tone": "dark", "eyebrow": "小标签",
      "title": "大字标题", "subtitle": ["副标题第一行", "第二行"],
      "foot_left": "左下角", "foot_right": "右下角" },

    { "type": "statement", "tone": "tint", "eyebrow": "标签",
      "title": "一句判断。", "body": ["补充一", "补充二"] },

    { "type": "bigstat", "tone": "tint", "title": "标题", "caption": "单位说明",
      "number": "67B", "notes": ["注释一", "注释二"],
      "items": [ { "title": "条目", "body": "说明" } ] },

    { "type": "stats", "tone": "dark", "title": "标题", "subtitle": "副标题",
      "stats": [ { "value": "236B", "label": "总参数", "note": "可选" } ],
      "footnote": "页脚注释" },

    { "type": "cards", "tone": "light", "title": "标题", "subtitle": "副标题",
      "cards": [ { "title": "卡1", "body": "说明", "tag": "可选标签" } ] },

    { "type": "lanes", "tone": "tint", "title": "标题",
      "items": [ { "title": "块", "detail": "展开", "tag": "口号" } ] },  // 默认 2 列

    { "type": "timeline", "tone": "light", "title": "标题", "subtitle": "副标题",
      "nodes": [ { "label": "2024.05", "name": "V2", "tag": "关键词", "pos": 0.25 } ] },

    { "type": "bullets", "tone": "light", "title": "标题",
      "bullets": [ { "title": "小标题", "body": "一到两行说明" } ] },

    { "type": "table", "tone": "light", "title": "标题",
      "columns": ["列1", "列2"], "widths": [2.0, 4.0],
      "bold_cols": [1], "muted_cols": [0],
      "rows": [["a", "b"]], "footnote": "来源" },

    { "type": "closing", "tone": "dark", "title": "收尾大字。",
      "subtitle": "一句收束", "link": "下一步 →", "align": "center" }
  ]
}
```

## 目录

```
ppt-studio/
├── SKILL.md                    # 本文件：工作流 + 选型 + 自检
├── styles/
│   ├── _SCHEMA.md              # ★ 新增风格必读：token 逐项说明
│   ├── apple.json  free.json  taobao.json  gov.json
├── scripts/
│   ├── deck_kit.py             # 渲染引擎（token 驱动 + 10 个版式原型）
│   ├── build_deck.py           # CLI：spec + style → pptx
│   └── preview.py              # Keynote→PDF→PNG，出联系表
└── examples/
    ├── spec.example.json       # 一份完整用例（6 页）
    └── spec.all-archetypes.json# 冒烟测试（10 页，每种版式一次）
```

## 扩展

- **加风格**：读 `styles/_SCHEMA.md`，复制最接近的 JSON 改 5 处即可，不改代码。
- **加版式**：在 `deck_kit.py` 写一个 `slide_xxx(d, s)` 并注册到 `ARCHETYPES`，签名与原语见 `_SCHEMA.md`。
- **加导出格式**：`preview.py` 里已有 PDF 链路，需要 PNG/PDF 交付时直接复用。

## 依赖

`python-pptx`（必需）、`lxml`（python-pptx 依赖）、`Pillow`+`numpy`（仅环境光 PNG）、
macOS + Keynote（仅预览渲染，缺了不影响出 pptx）。

## 已知边界

- **比例**：默认 16:9（13.333 × 7.5）。要 4:3 就改 spec 的 `width: 10, height: 7.5`——
  版式会自适应，但字号是按 16:9 调的，4:3 建议把 `type` 整体降 10–15%。
- **字体**：样式包用的 PingFang SC / Songti SC / Helvetica Neue 是 macOS 自带。
  目标机器没有时会回落到系统默认字体，版式不塌但气质会变——给非 Mac 用户就用
  Windows 常见字体（微软雅黑 / 等线 / Arial）另存一个样式包。
- **环境光**：需要 Pillow + numpy；缺了会自动跳过，深色页退回纯色。
- **预览**：Keynote 导出 PDF 再光栅化，能忠实还原字体与行距；没有 Keynote 时
  只能出第一页缩略图，此时请以 pptx 打开后的实际效果为准。
  中文路径已处理（内部走临时 ASCII 目录中转）；若预览是 0 张，加 `-v` 看具体错误。
