#!/usr/bin/env python3
"""把「内容 spec + 样式包」渲染成 .pptx。

    python build_deck.py --spec deck.json --style apple -o out.pptx
    python build_deck.py --spec deck.json --style ./my_style.json -o out.pptx --preview

spec 结构见 examples/spec.example.json；版式原型见 SKILL.md。
"""

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from deck_kit import ARCHETYPES, build, load_style  # noqa: E402

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main(argv=None):
    ap = argparse.ArgumentParser(description="style-driven pptx builder")
    ap.add_argument("--spec", help="内容 JSON")
    ap.add_argument("--list", action="store_true", help="列出可用样式与版式")
    ap.add_argument("--style", default="free", help="样式名（styles/ 下）或 .json 绝对路径")
    ap.add_argument("-o", "--out", help="输出 .pptx 路径")
    ap.add_argument("--preview", action="store_true", help="生成完顺手导出 PNG 预览")
    ap.add_argument("--preview-dir", default=None)
    args = ap.parse_args(argv)

    if args.list:
        import glob
        base = os.path.join(SKILL_DIR, "styles")
        print("样式（styles/）：")
        for f in sorted(glob.glob(os.path.join(base, "*.json"))):
            if os.path.basename(f).startswith("_"):
                continue
            st = load_style(f)
            print(f"  - {st.name:10s} {st.label}")
        print("\n版式（archetype）：")
        print("  " + ", ".join(ARCHETYPES))
        return 0

    if not args.spec:
        ap.error("需要 --spec，或用 --list 查看可用样式")
    if not args.out:
        ap.error("需要 -o/--out 指定输出文件")

    with open(args.spec, encoding="utf-8") as f:
        spec = json.load(f)

    style = load_style(args.style, os.path.join(SKILL_DIR, "styles"))
    build(spec, style, args.out)

    n = len(spec.get("slides", []))
    used = {}
    for s in spec.get("slides", []):
        used[s.get("type")] = used.get(s.get("type"), 0) + 1
    summary = " · ".join(f"{k}×{v}" for k, v in used.items())
    print(f"✅ {os.path.basename(args.out)}")
    print(f"   样式：{style.label} ({style.name})　页数：{n}")
    print(f"   版式：{summary}")
    print(f"   可用版式：{', '.join(ARCHETYPES)}")

    if args.preview:
        from preview import render
        out_dir = args.preview_dir or os.path.splitext(args.out)[0] + "_preview"
        files = render(args.out, out_dir)
        print(f"   预览：{len(files)} 张 → {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
