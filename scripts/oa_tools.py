# -*- coding: utf-8 -*-
"""
OA 报销打包工具集 —— 实战验证过的函数
依赖: pypdf, Pillow, numpy
"""
import os, io, re, shutil, zipfile, subprocess
import numpy as np
from PIL import Image

INVOICE_PT = (595.276, 396.85)   # 发票：横版
SLIP_PT    = (595.0, 842.0)      # 出库单：A4 竖版


# ---------- 1. 从扫描 PDF 无损取每页位图 ----------
def extract_pages(scan_pdf, out_dir):
    """提取每页嵌入的原始位图，不重编码。返回路径列表。"""
    import pypdf
    os.makedirs(out_dir, exist_ok=True)
    r = pypdf.PdfReader(scan_pdf)
    paths = []
    for i, p in enumerate(r.pages):
        img = p.images[0].image
        path = os.path.join(out_dir, "p%d.png" % (i + 1))
        img.save(path)
        paths.append(path)
        print("page %d: %dx%d" % (i + 1, img.width, img.height))
    return paths


# ---------- 2. 亮度归一化（去阴影/不均光） ----------
def norm(g):
    sw, sh = max(1, g.shape[1] // 48), max(1, g.shape[0] // 48)
    small = np.asarray(Image.fromarray(g.astype(np.uint8)).resize((sw, sh), Image.BOX)).astype(np.float32)
    bg = np.asarray(Image.fromarray(np.clip(small, 0, 255).astype(np.uint8))
                    .resize((g.shape[1], g.shape[0]), Image.BILINEAR)).astype(np.float32)
    return np.clip(g / np.maximum(bg, 1) * 235, 0, 255)


def _longest_run(mask, gap):
    """最长连续 True 段 —— 用它能跳过折痕/污渍造成的孤立行"""
    idx = np.where(mask)[0]
    if len(idx) == 0:
        return (0, -1)
    best = cur = (idx[0], idx[0])
    for v in idx[1:]:
        if v - cur[1] <= gap:
            cur = (cur[0], v)
        else:
            if cur[1] - cur[0] > best[1] - best[0]:
                best = cur
            cur = (v, v)
    if cur[1] - cur[0] > best[1] - best[0]:
        best = cur
    return best


# ---------- 3. 票据内容包围盒（抗折痕） ----------
def content_box(im, rowf=0.004, colf=0.002, pad=0.008):
    """返回 (x0,y0,x1,y1)。折痕不会被误当成内容。"""
    g = np.asarray(im.convert("L")).astype(np.float32)
    ink = norm(g) < 168
    H, W = ink.shape
    r0, r1 = _longest_run(ink.mean(axis=1) > rowf, int(H * 0.02))
    band = ink[r0:r1 + 1] if r1 > r0 else ink
    c0, c1 = _longest_run(band.mean(axis=0) > colf, int(W * 0.02))
    if r1 <= r0 or c1 <= c0:
        return (0, 0, W, H)
    px, py = int(W * pad), int(H * pad)
    return (max(0, c0 - px), max(0, r0 - py), min(W, c1 + px), min(H, r1 + py))


# ---------- 4. 摆到指定页面规格 ----------
def make_pdf(img, kind, dst, dpi=300.0):
    """kind: "invoice" | "slip"。发票裁紧横版，出库单保留整张 A4。"""
    if kind == "invoice":
        pt, mg = INVOICE_PT, (0.008, 0.010, 0.008, 0.014)
        img = img.crop(content_box(img))
    elif kind == "slip":
        pt, mg = SLIP_PT, (0.010, 0.012, 0.010, 0.012)
    else:
        raise ValueError("kind must be invoice or slip")
    W = int(round(pt[0] / 72.0 * dpi))
    H = int(round(pt[1] / 72.0 * dpi))
    mw, mt, mr, mb = mg
    aw, ah = W * (1 - mw - mr), H * (1 - mt - mb)
    s = min(aw / img.width, ah / img.height)
    nw, nh = max(1, int(img.width * s)), max(1, int(img.height * s))
    c = Image.new("RGB", (W, H), (255, 255, 255))
    c.paste(img.resize((nw, nh), Image.LANCZOS),
            (int(mw * W + (aw - nw) / 2), int(mt * H)))
    c.save(dst, "PDF", resolution=dpi, quality=95)
    return os.path.getsize(dst)


# ---------- 5. 改 docx 文本（保留全部样式） ----------
def patch_docx_text(src, dst, replacements):
    """replacements: {old: new}。只换文本，表格/字体/页边距不动。"""
    tmp = dst + ".tmp"
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}
    xml = data["word/document.xml"].decode("utf-8")
    for old, new in replacements.items():
        xml = xml.replace(old, new)
    data["word/document.xml"] = xml.encode("utf-8")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, data[n])
    os.replace(tmp, dst)
    return dst


# ---------- 6. 打包 zip ----------
def make_zip(src_dir, zip_path, arc_root="OA"):
    """遍历目录打包，UTF-8 文件名。"""
    entries = []

    def walk(d, prefix):
        for name in sorted(os.listdir(d)):
            full = os.path.join(d, name)
            if os.path.isdir(full):
                walk(full, prefix + name + "/")
            else:
                entries.append((prefix + name, full))

    walk(src_dir, arc_root + "/")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for arc, full in entries:
            z.write(full, arc)
    return len(entries), os.path.getsize(zip_path)


# ---------- 7. 金额校验 ----------
def check_amounts(items, total_str):
    """items: [(name, spec, qty, unit_price, amount), ...]
    total_str: "¥7,720.00"
    返回 (ok, 明细合计, 票面合计, 差额)"""
    def num(s):
        return float(re.sub(r"[^0-9.]", "", s))
    s = 0.0
    for it in items:
        qty, unit, amt = num(it[2]), num(it[3]), num(it[4])
        s += amt
        if abs(qty * unit - amt) > 0.01:
            print("  !! %s: %g x %g = %.2f != %g" % (it[0], qty, unit, qty * unit, amt))
    t = num(total_str)
    d = round(s - t, 2)
    return (abs(d) <= 0.01, round(s, 2), t, d)


# ---------- 8. 自检 ----------
def verify(zip_path):
    ok = True
    with zipfile.ZipFile(zip_path) as z:
        if z.testzip() is not None:
            print("!! zip 损坏")
            ok = False
        for n in z.namelist():
            if not n.lower().endswith(".pdf"):
                continue
            tmp = os.path.join(os.environ.get("TEMP", "."), "_chk.pdf")
            with open(tmp, "wb") as f:
                f.write(z.read(n))
            out = subprocess.run(["pdfinfo", tmp], capture_output=True, text=True,
                                 encoding="utf-8", errors="replace").stdout
            pg = re.search(r"Pages:\s+(\d+)", out)
            sz = re.search(r"Page size:\s+([\d.]+) x ([\d.]+)", out)
            w, h = (float(sz.group(1)), float(sz.group(2))) if sz else (0, 0)
            kind = "landscape" if w > h else "portrait"
            print("  %-34s pages=%s %5.0fx%-5.0f %s" %
                  (n, pg.group(1) if pg else "?", w, h, kind))
    print("verify:", "PASS" if ok else "FAIL")
    return ok


if __name__ == "__main__":
    print("oa_tools ready")
    print("INVOICE_PT", INVOICE_PT)
    print("SLIP_PT", SLIP_PT)
