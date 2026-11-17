#!/usr/bin/env python3
"""把 .pptx 渲染成逐页 PNG，用于人工/自动视觉检查。

链路：Keynote 导出 PDF → Quartz 逐页光栅化（macOS 原生，无需额外依赖）。
没有 Keynote 时退化为 qlmanage（只能出第一页缩略图）。

    python preview.py deck.pptx out_dir
    python preview.py deck.pptx out_dir --scale 2 --sheet   # 额外拼一张联系表
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile


def _have(cmd: str) -> bool:
    return shutil.which(cmd) is not None


_STAGE = {"dir": None}


def _stage_dir() -> str:
    if _STAGE["dir"] is None:
        _STAGE["dir"] = tempfile.mkdtemp(prefix="pptpng_")
    return _STAGE["dir"]


def _ascii_only(p: str) -> bool:
    try:
        p.encode("ascii")
        return True
    except UnicodeEncodeError:
        return False


def pptx_to_pdf(pptx: str, pdf: str, verbose: bool = False) -> bool:
    """用 Keynote 导出 PDF（忠实还原字体与间距）。

    注意：Keynote 的 AppleScript 导出对非 ASCII 路径会静默失败，
    调用方需先把文件暂存到纯 ASCII 路径（见 render()）。
    """
    if not os.path.isdir("/Applications/Keynote.app"):
        return False
    script = """
on run argv
    set inFile to POSIX file (item 1 of argv)
    set outFile to POSIX file (item 2 of argv)
    tell application "Keynote"
        activate
        set doc to open inFile
        with timeout of 900 seconds
            export doc to outFile as PDF with properties {PDF image quality:Best}
        end timeout
        close doc saving no
    end tell
end run
"""
    with tempfile.NamedTemporaryFile("w", suffix=".scpt", delete=False) as f:
        f.write(script)
        scpt = f.name
    try:
        proc = subprocess.run(["osascript", scpt, pptx, pdf],
                              capture_output=True, text=True)
        if verbose and proc.returncode != 0:
            print(f"⚠️ Keynote 导出失败：{proc.stderr.strip()[:400]}", file=sys.stderr)
        return proc.returncode == 0 and os.path.isfile(pdf)
    except Exception as e:  # noqa: BLE001
        if verbose:
            print(f"⚠️ Keynote 导出异常：{e}", file=sys.stderr)
        return False
    finally:
        os.unlink(scpt)


def pdf_to_png(pdf: str, out_dir: str, scale: float = 2.0):
    """把 PDF 逐页写成 PNG。

    Quartz 的 CGImageDestination 对非 ASCII 输出路径会静默失败，
    所以这里先写进临时 ASCII 目录，再逐个搬到目标目录。
    """
    import Quartz

    files = []
    staged: list = []
    url = Quartz.CFURLCreateFromFileSystemRepresentation(None, pdf.encode(), len(pdf), False)
    doc = Quartz.CGPDFDocumentCreateWithURL(url)
    n = Quartz.CGPDFDocumentGetNumberOfPages(doc)
    for i in range(1, n + 1):
        page = Quartz.CGPDFDocumentGetPage(doc, i)
        rect = Quartz.CGPDFPageGetBoxRect(page, Quartz.kCGPDFMediaBox)
        w, h = int(rect.size.width * scale), int(rect.size.height * scale)
        cs = Quartz.CGColorSpaceCreateDeviceRGB()
        ctx = Quartz.CGBitmapContextCreate(None, w, h, 8, 0, cs,
                                           Quartz.kCGImageAlphaPremultipliedFirst)
        Quartz.CGContextSetRGBFillColor(ctx, 1, 1, 1, 1)
        Quartz.CGContextFillRect(ctx, Quartz.CGRectMake(0, 0, w, h))
        Quartz.CGContextScaleCTM(ctx, scale, scale)
        Quartz.CGContextDrawPDFPage(ctx, page)
        img = Quartz.CGBitmapContextCreateImage(ctx)
        if _ascii_only(out_dir):
            path = os.path.join(out_dir, f"p{i:02d}.png")
        else:
            path = os.path.join(_stage_dir(), f"p{i:02d}.png")
            staged.append((path, os.path.join(out_dir, f"p{i:02d}.png")))
        dest = Quartz.CGImageDestinationCreateWithURL(
            Quartz.CFURLCreateFromFileSystemRepresentation(None, path.encode(), len(path), False),
            "public.png", 1, None)
        Quartz.CGImageDestinationAddImage(dest, img, None)
        Quartz.CGImageDestinationFinalize(dest)
        files.append(path)
    for src_p, dst_p in staged:
        shutil.move(src_p, dst_p)
    return [dst for _, dst in staged] or files


def contact_sheet(files, path, cols=3, tile_w=760):
    from PIL import Image
    if not files:
        return None
    with Image.open(files[0]) as im0:
        ratio = im0.height / im0.width
    tw, th = tile_w, int(tile_w * ratio)
    rows = (len(files) + cols - 1) // cols
    pad = 14
    sheet = Image.new("RGB", (cols * tw + (cols + 1) * pad, rows * th + (rows + 1) * pad),
                      (242, 242, 244))
    for i, f in enumerate(files):
        im = Image.open(f).resize((tw, th), Image.LANCZOS)
        c, r = i % cols, i // cols
        sheet.paste(im, (pad + c * (tw + pad), pad + r * (th + pad)))
    sheet.save(path)
    return path


def render(pptx: str, out_dir: str, scale: float = 2.0, sheet: bool = False,
           verbose: bool = False):
    os.makedirs(out_dir, exist_ok=True)

    # 关键：Keynote 导出不认非 ASCII 路径，先暂存到临时 ASCII 目录再导出
    with tempfile.TemporaryDirectory(prefix="pptprev_") as td:
        src = pptx
        if not _ascii_only(pptx):
            src = os.path.join(td, "deck.pptx")
            shutil.copy2(pptx, src)
        pdf = os.path.join(td, "deck.pdf")

        if pptx_to_pdf(src, pdf, verbose):
            files = pdf_to_png(pdf, out_dir, scale)
            if sheet:
                contact_sheet(files, os.path.join(out_dir, "_sheet.png"))
            return files

    # 退化方案：Quick Look 只能渲染第一页
    if _have("qlmanage"):
        subprocess.run(["qlmanage", "-t", "-s", "1800", "-o", out_dir, pptx],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        for f in os.listdir(out_dir):
            if f.endswith(".png"):
                return [os.path.join(out_dir, f)]
    print("⚠️ 无法渲染预览（需要 macOS + Keynote）。PPTX 本身已生成。", file=sys.stderr)
    return []


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("out_dir")
    ap.add_argument("--scale", type=float, default=2.0)
    ap.add_argument("--sheet", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    a = ap.parse_args(argv)
    files = render(a.pptx, a.out_dir, a.scale, a.sheet, a.verbose)
    print("\n".join(files))
    return 0 if files else 1


if __name__ == "__main__":
    raise SystemExit(main())
