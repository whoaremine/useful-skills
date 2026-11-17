# 样式包 token 规范（style pack）

一个样式 = `styles/<name>.json` 里的**设计令牌**，不含任何内容。
`scripts/deck_kit.py` 是唯一读者，改 token 就能换风格，不用碰代码。

```jsonc
{
  "name": "my-style",          // 文件名（不带 .json），调用 --style 时用它
  "label": "我的风格",          // 给人看的中文名
  "description": "一句话说明适用场景",
  "voice": "文案语气约束：给规划页面的 AI 看，决定标题长度与句式",

  "fonts": {
    "latin": "Helvetica Neue",       // 数字/英文
    "cjk": "PingFang SC",            // 正文中文
    "cjk_bold": "PingFang SC Semibold", // 加粗中文用的家族（见 weight_mode）
    "weight_mode": "split"           // split | flag（见下）
  },

  "colors": {
    "dark_bg": "#000000", "light_bg": "#FFFFFF", "tint_bg": "#F5F5F7",
    "ink": "#1D1D1F", "ink_dark": "#FFFFFF",
    "muted": "#86868B", "muted_dark": "#A1A1A6",
    "hair": "#D2D2D7", "hair_dark": "#2C2C2E",
    "accent": "#0071E3", "accent_dark": "#299BFF",
    "card": "#F5F5F7", "card_dark": "#161618",
    "card_line": null, "card_line_dark": null,
    "shadow": "#000000"              // 卡片投影颜色（hard/soft 都取它）
  },

  "type": {                          // 全部单位 pt
    "hero": 96, "title": 52, "title_lg": 42,
    "lead": 19, "body": 14, "body_xl": 16.5, "body_sm": 12.5,
    "small": 12, "micro": 10.5, "micro_faint": 10, "eyebrow": 10.5,
    "stat_big": 62, "stat_huge": 124, "card_title": 20
  },

  "tracking": { "hero": -2.4, "title": -1.3, "eyebrow": 1.1 },  // pt，负=收紧

  "geometry": {                      // 单位 inch
    "margin_x": 0.95,                // 左右版心
    "header_y": 0.92,                // eyebrow 的 y
    "content_y": 1.34,               // 标题起始 y
    "cover_y": 1.6, "foot_y": 5.72,
    "line_factor": 1.34,             // 行高推算系数（中文 1.34 稳）
    "lead_line": 1.5,                // 副标题行距倍数
    "stats_y": 3.86, "cards_y": 3.9,
    "list_y": 2.34, "list_step": 1.28
  },

  "decor": {
    "card_shape": "rounded",         // rounded | rect
    "card_radius": 0.05,             // 圆角（占短边比例）
    "card_border": 0,                // pt，0=无描边
    "card_shadow": "none",           // none | soft | hard（hard=贴纸风错位硬投影）
    "shadow_offset": 0.07,
    "hairline": true,
    "glow": true,                    // 深色页环境光
    "glow_color": "#5A78FF", "glow_alpha": 0.42,
    "glow_x": 4.2, "glow_y": -1.6, "glow_w": 10.5, "glow_h": 6.0,
    "bg_pattern": "none",            // none | dots | grid
    "pattern_step": 0.34, "pattern_dot": 0.012,
    "top_bar": { "h": 0.14, "color": "#B3251F" },   // 顶部色条（公文风）
    "table_head_bg": "#FFE3D2",                     // 表头底色
    "table_head_ink": "#FFFFFF",                    // 表头字色
    "table_zebra": "#FFF8F2"                        // 隔行底色
  }
}
```

## weight_mode 为什么重要

中文用 `bold` 标志会被渲染成 PingFang **Bold**，比苹果/现代排版想要的 **Semibold** 重一档，
大字号尤其明显。两种模式：

| 模式 | 行为 | 适合 |
|---|---|---|
| `split`（推荐） | 中文走 `cjk_bold` 家族（如 PingFang SC Semibold），数字/英文走 `bold` 标志 | 极简、官网、正式汇报 |
| `flag` | 全部走 `bold` 标志 | 电商、促销、需要重量感的风格 |

## 新增一个风格（3 步）

1. **复制最接近的**：`cp styles/free.json styles/mystyle.json`
2. **改 5 处**：`name/label`、`colors`（主色 + 底色 + 文字色）、`fonts`（要不要换字体家族）、
   `decor`（圆角/描边/投影/顶部色条）、`type`（整体字号档位：极简调大、公文调小）
3. **试渲染**：
   ```bash
   python scripts/build_deck.py --spec examples/spec.all-archetypes.json --style mystyle -o /tmp/mystyle.pptx --preview
   ```
   打开 `_sheet.png` 逐页看，确认十种版式都成立。

### 两个现成范例（可直接改成你要的风格）

| 想做的风格 | 抄这个 | 关键 token |
|---|---|---|
| 电商/促销/社群 | `taobao.json` | `card_shadow: hard` + `card_border: 2.25` + `card_radius: 0.16` + 暖色 `glow` |
| 行政/公文/汇报 | `gov.json` | `card_shape: rect` + `card_radius: 0` + `cjk_bold: Songti SC` + `top_bar` + `table_head_bg` |

再举两个一句话就能调出来的例子：

- **小红书/种草**：`light_bg:#FFF8F0`、`accent:#FF2E4D`、`card_radius:0.2`、`hero:84`、`weight_mode:flag`
- **科技暗色**：`dark_bg:#0B0F1A`、`tint_bg:#0F1526`、`light_bg:#0B0F1A`（全暗）、`accent:#00E5A0`、`glow:true`

## 版式（archetype）如何扩展

版式 = `scripts/deck_kit.py` 里的一个函数，签名固定：

```python
def slide_mytimeline(d: Deck, s: dict):
    d.new_slide(s.get("tone", "light"))     # 建页 + 底色 + 顶条 + 环境光
    d.eyebrow(s.get("eyebrow"))
    y = d.st.g("content_y", 1.34)
    d.text(d.M, y, d.CW, s.get("title", ""), role="title", bold=True)
    ...
    return None

ARCHETYPES["mytimeline"] = slide_mytimeline       # 注册后即可在 spec 里用 "type": "mytimeline"
```

可用原语：

| 原语 | 用途 |
|---|---|
| `d.new_slide(tone)` | `dark` / `light` / `tint` |
| `d.text(x, y, w, body, role=..., bold=, align=, line=)` | 自动按 token 取字号/字距；`body` 可传字符串数组当多段 |
| `d.disp_bottom(y, size, lines, ls)` | **必用**：推算文本真实占位，返回下一元素的 y |
| `d.card(x, y, w, h)` | 画卡片（圆角/描边/投影由样式包决定） |
| `d.hairline(x, y, w)` / `d.dot(cx, cy, r)` | 发丝线 / 圆点 |
| `d.rect(x, y, w, h, fill, radius, border, border_w)` | 任意形状 |
| `decorate_table(d, tbl, s, cols, rows)` | 表格样式（只用横向发丝线） |
