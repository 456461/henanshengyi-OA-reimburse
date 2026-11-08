# -*- coding: utf-8 -*-
"""
OA 报销一键流水线 —— 把所有步骤合并成一次执行

用法：
    python oa_pipeline.py <工作目录>

工作目录里需要有：
    invoice/     扫描件 PDF（发票+出库单）
    sign/        手写签字照片（按任意顺序，脚本会生成拼图供核对）
    *.docx       验收单母版

输出：
    <工作目录>/OUTPUT/OA.zip
    <工作目录>/OUTPUT/_contact_sheet.jpg   <- 签字照拼图，一次性核对用
"""
import sys, io, os, re, json, shutil, zipfile, subprocess, time
from io import BytesIO

if not getattr(sys.stdout, "_oa_wrapped", False):
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        sys.stdout._oa_wrapped = True
    except Exception:
        pass

import numpy as np
from PIL import Image
import pypdf

INVOICE_PT = (595.276, 396.85)
SLIP_PT    = (595.0, 842.0)


# ============ 工具函数（内联，避免跨文件导入开销） ============
def _norm(g):
    sw, sh = max(1, g.shape[1]//48), max(1, g.shape[0]//48)
    small = np.asarray(Image.fromarray(g.astype(np.uint8)).resize((sw, sh), Image.BOX)).astype(np.float32)
    bg = np.asarray(Image.fromarray(np.clip(small,0,255).astype(np.uint8)).resize((g.shape[1],g.shape[0]), Image.BILINEAR)).astype(np.float32)
    return np.clip(g/np.maximum(bg,1)*235, 0, 255)

def _longest_run(mask, gap):
    idx = np.where(mask)[0]
    if len(idx) == 0: return (0, -1)
    best = cur = (idx[0], idx[0])
    for v in idx[1:]:
        if v - cur[1] <= gap: cur = (cur[0], v)
        else:
            if cur[1]-cur[0] > best[1]-best[0]: best = cur
            cur = (v, v)
    if cur[1]-cur[0] > best[1]-best[0]: best = cur
    return best

def content_box(im, rowf=0.004, colf=0.002, pad=0.008):
    g = np.asarray(im.convert("L")).astype(np.float32)
    ink = _norm(g) < 168
    H, W = ink.shape
    r0, r1 = _longest_run(ink.mean(axis=1) > rowf, int(H*0.02))
    band = ink[r0:r1+1] if r1 > r0 else ink
    c0, c1 = _longest_run(band.mean(axis=0) > colf, int(W*0.02))
    if r1 <= r0 or c1 <= c0: return (0, 0, W, H)
    px, py = int(W*pad), int(H*pad)
    return (max(0,c0-px), max(0,r0-py), min(W,c1+px), min(H,r1+py))

def make_pdf(img, kind, dst, dpi=300.0):
    if kind == "invoice":
        pt, mg = INVOICE_PT, (0.008, 0.010, 0.008, 0.014)
        img = img.crop(content_box(img))
    else:
        pt, mg = SLIP_PT, (0.010, 0.012, 0.010, 0.012)
    W = int(round(pt[0]/72.0*dpi)); H = int(round(pt[1]/72.0*dpi))
    mw, mt, mr, mb = mg
    aw, ah = W*(1-mw-mr), H*(1-mt-mb)
    s = min(aw/img.width, ah/img.height)
    nw, nh = max(1,int(img.width*s)), max(1,int(img.height*s))
    c = Image.new("RGB", (W, H), (255,255,255))
    c.paste(img.resize((nw, nh), Image.LANCZOS), (int(mw*W+(aw-nw)/2), int(mt*H)))
    c.save(dst, "PDF", resolution=dpi, quality=95)


# ============ 拼图预览（关键提速点） ============
def make_contact_sheet(images, dst, cols=3, cell_w=620):
    """把 N 张图拼成 1 张，一次就能核对全部"""
    if not images: return None
    ims = [Image.open(p) if isinstance(p, str) else p for p in images]
    rows = (len(ims) + cols - 1) // cols
    ch = int(cell_w * ims[0].height / max(1, ims[0].width))
    sheet = Image.new("RGB", (cols*cell_w, rows*ch), "white")
    for i, im in enumerate(ims):
        t = im.resize((cell_w, ch), Image.LANCZOS)
        sheet.paste(t, ((i % cols)*cell_w, (i // cols)*ch))
    sheet.save(dst, quality=82)
    return dst


# ============ 签字照透视校正 ============
def warp_signature(img_bgr):
    import cv2
    hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
    L = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2LAB)[:, :, 0]
    s = hsv[:, :, 1]
    m = ((L > 140) & (s < 65)).astype(np.uint8) * 255
    m = cv2.morphologyEx(m, cv2.MORPH_CLOSE, np.ones((61,61), np.uint8))
    m = cv2.morphologyEx(m, cv2.MORPH_OPEN, np.ones((61,61), np.uint8))
    cnts, _ = cv2.findContours(m, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts: return None, 0, 0
    c = max(cnts, key=cv2.contourArea)
    if cv2.contourArea(c) < img_bgr.shape[0]*img_bgr.shape[1]*0.3: return None, 0, 0
    rect = cv2.minAreaRect(c)
    pts = cv2.boxPoints(rect)
    ys = pts[:, 1]; mid = (ys.min()+ys.max())/2
    top, bot = pts[ys <= mid], pts[ys > mid]
    if len(top) != 2 or len(bot) != 2: return None, 0, 0
    tl = top[np.argmin(top[:,0])]; tr = top[np.argmax(top[:,0])]
    bl = bot[np.argmin(bot[:,0])]; br = bot[np.argmax(bot[:,0])]
    quad = np.array([tl, tr, br, bl], dtype=np.float32)
    wtop = float(np.linalg.norm(quad[1]-quad[0]))
    hleft = float(np.linalg.norm(quad[3]-quad[0]))
    if wtop < hleft:
        quad = np.array([quad[1], quad[2], quad[3], quad[0]], dtype=np.float32)
    W = int(max(np.linalg.norm(quad[1]-quad[0]), np.linalg.norm(quad[2]-quad[3])))
    H = int(max(np.linalg.norm(quad[3]-quad[0]), np.linalg.norm(quad[2]-quad[1])))
    if W <= 0 or H <= 0: return None, 0, 0
    dst = np.array([[0,0],[W-1,0],[W-1,H-1],[0,H-1]], dtype=np.float32)
    M = cv2.getPerspectiveTransform(quad, dst)
    return cv2.warpPerspective(img_bgr, M, (W, H)), W, H


# ============ docx 表格重建 ============
def esc(s):
    return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def set_cell(cell, text, force_sz=None):
    tcPr = re.search(r"<w:tcPr>.*?</w:tcPr>", cell, re.S)
    tcPr = tcPr.group(0) if tcPr else ""
    p = re.search(r"<w:p\b.*?</w:p>", cell, re.S)
    pxml = p.group(0) if p else "<w:p/>"
    pPr = re.search(r"<w:pPr>.*?</w:pPr>", pxml, re.S)
    pPr = pPr.group(0) if pPr else ""
    r = re.search(r"<w:r\b.*?</w:r>", pxml, re.S)
    rPr = ""
    if r:
        mm = re.search(r"<w:rPr>.*?</w:rPr>", r.group(0), re.S)
        rPr = mm.group(0) if mm else ""
    if force_sz and rPr:
        rPr = re.sub(r'<w:sz w:val="\d+"/>', '<w:sz w:val="%d"/>' % force_sz, rPr)
        rPr = re.sub(r'<w:szCs w:val="\d+"/>', '<w:szCs w:val="%d"/>' % force_sz, rPr)
    return ('<w:tc>' + tcPr + '<w:p>' + pPr + '<w:r>' + rPr +
            '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p></w:tc>')

def build_docx(src, dst, project, project_code, date, seller, invoice_no, items, qty_total, amt_total):
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}
    xml = data["word/document.xml"].decode("utf-8")

    # 段落字段
    xml = re.sub(r"(<w:t[^>]*>)骨肉瘤在重离子放射治疗中的生物学效应研究(</w:t>)",
                 lambda m: m.group(1) + esc(project) + m.group(2), xml, count=1)
    # 项目编码
    xml = re.sub(r"(项目编码：</w:t></w:r>.*?<w:r\b[^>]*>(?:(?!</w:r>).)*?<w:t[^>]*>)[^<]*(</w:t>)",
                 lambda m: m.group(1) + esc(project_code) + m.group(2), xml, count=1, flags=re.S)
    # 验收时间
    xml = re.sub(r"(验收时间：</w:t></w:r>.*?<w:r\b[^>]*>(?:(?!</w:r>).)*?<w:t[^>]*>)[^<]*(</w:t>)",
                 lambda m: m.group(1) + esc(date) + m.group(2), xml, count=1, flags=re.S)

    s = xml.index("<w:tbl>"); e = xml.index("</w:tbl>") + len("</w:tbl>")
    tbl = xml[s:e]
    rows = re.findall(r"<w:tr\b.*?</w:tr>", tbl, re.S)
    hdr, colhdr, item_tpl, total_tpl = rows[0], rows[1], rows[2], rows[5]

    # 表头
    hc = re.findall(r"<w:tc>.*?</w:tc>", hdr, re.S)
    vals = [None, None, None, seller, None, invoice_no, None]
    new_hc = []
    for c, v in zip(hc, vals):
        if v is None: new_hc.append(c)
        elif v == seller: new_hc.append(set_cell(c, v, force_sz=20))
        else: new_hc.append(set_cell(c, v))
    hdr_new = re.sub(r"(<w:tr\b.*?>)(.*)(</w:tr>)", lambda m: m.group(1)+"".join(new_hc)+m.group(3), hdr, count=1, flags=re.S)

    # 明细行
    tpl_cells = re.findall(r"<w:tc>.*?</w:tc>", item_tpl, re.S)
    item_rows = []
    for i, it in enumerate(items):
        cells = []
        for j, c in enumerate(tpl_cells):
            if j == len(tpl_cells) - 1:
                cells.append(c if i == 0 else c.replace('<w:vMerge w:val="restart"/>', "<w:vMerge/>"))
            else:
                cells.append(set_cell(c, it[j] if j < len(it) else ""))
        item_rows.append(re.sub(r"(<w:tr\b.*?>)(.*)(</w:tr>)", lambda m: m.group(1)+"".join(cells)+m.group(3), item_tpl, count=1, flags=re.S))

    # 合计行
    tc = re.findall(r"<w:tc>.*?</w:tc>", total_tpl, re.S)
    tvals = [None, qty_total, None, amt_total, None]
    new_tc = [c if v is None else set_cell(c, v) for c, v in zip(tc, tvals)]
    total_new = re.sub(r"(<w:tr\b.*?>)(.*)(</w:tr>)", lambda m: m.group(1)+"".join(new_tc)+m.group(3), total_tpl, count=1, flags=re.S)

    tbl_new = tbl[:tbl.index(rows[0])] + hdr_new + colhdr + "".join(item_rows) + total_new + tbl[tbl.index(rows[5])+len(rows[5]):]
    xml = xml[:s] + tbl_new + xml[e:]

    data["word/document.xml"] = xml.encode("utf-8")
    tmp = dst + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names: z.writestr(n, data[n])
    os.replace(tmp, dst)


# ============ 打包 ============
def make_zip(src_dir, zip_path, arc_root="OA"):
    import zlib
    CRC = []
    for n in range(256):
        c = n
        for _ in range(8):
            c = (0xEDB88320 ^ (c >> 1)) if (c & 1) else (c >> 1)
        CRC.append(c)
    def crc32(b):
        c = 0xFFFFFFFF
        for x in b: c = CRC[(c ^ x) & 0xFF] ^ (c >> 8)
        return (c ^ 0xFFFFFFFF) & 0xFFFFFFFF

    entries = []
    for root, dirs, files in os.walk(src_dir):
        for f in sorted(files):
            full = os.path.join(root, f)
            rel = os.path.relpath(full, src_dir).replace("\\", "/")
            entries.append((arc_root + "/" + rel, full))

    chunks, central = [], []
    offset = 0
    for arc, full in entries:
        with open(full, "rb") as fh: raw = fh.read()
        comp = zlib.compressobj(9, zlib.DEFLATED, -15)
        packed = comp.compress(raw) + comp.flush()
        crc = crc32(raw)
        nb = arc.encode("utf-8")
        lh = bytearray(30)
        lh[0:4] = (0x04034b50).to_bytes(4, "little")
        lh[4:6] = (20).to_bytes(2, "little")
        lh[6:8] = (0x800).to_bytes(2, "little")
        lh[8:10] = (8).to_bytes(2, "little")
        lh[14:18] = crc.to_bytes(4, "little")
        lh[18:22] = len(packed).to_bytes(4, "little")
        lh[22:26] = len(raw).to_bytes(4, "little")
        lh[26:28] = len(nb).to_bytes(2, "little")
        chunks.append(bytes(lh)); chunks.append(nb); chunks.append(packed)
        ch = bytearray(46)
        ch[0:4] = (0x02014b50).to_bytes(4, "little")
        ch[4:6] = (0x031E).to_bytes(2, "little")
        ch[6:8] = (20).to_bytes(2, "little")
        ch[8:10] = (0x800).to_bytes(2, "little")
        ch[10:12] = (8).to_bytes(2, "little")
        ch[16:20] = crc.to_bytes(4, "little")
        ch[20:24] = len(packed).to_bytes(4, "little")
        ch[24:28] = len(raw).to_bytes(4, "little")
        ch[28:30] = len(nb).to_bytes(2, "little")
        ch[38:42] = (0x20).to_bytes(4, "little")
        ch[42:46] = offset.to_bytes(4, "little")
        central.append(bytes(ch)); central.append(nb)
        offset += 30 + len(nb) + len(packed)

    cd = b"".join(central)
    eocd = bytearray(22)
    eocd[0:4] = (0x06054b50).to_bytes(4, "little")
    eocd[8:10] = len(entries).to_bytes(2, "little")
    eocd[10:12] = len(entries).to_bytes(2, "little")
    eocd[12:16] = len(cd).to_bytes(4, "little")
    eocd[16:20] = offset.to_bytes(4, "little")
    with open(zip_path, "wb") as fh:
        fh.write(b"".join(chunks)); fh.write(cd); fh.write(bytes(eocd))
    return len(entries), os.path.getsize(zip_path)


if __name__ == "__main__":
    print("oa_pipeline ready. 用 build_all.py 驱动。")
