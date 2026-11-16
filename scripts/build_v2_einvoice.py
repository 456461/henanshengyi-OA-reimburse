# -*- coding: utf-8 -*-
"""
OA 报销 V2 —— 发票一律用电子原件，禁止扫描版

规矩变更（2026-11-16）:
  发票 <尾号>.pdf   = 电子发票原件，原样复制，一个字节都不改
  出库单 <尾号>.1.pdf = 扫描件切页摆 A4 竖版
  验收单              = 干净母版生成
  签字照              = 照片透视校正 -> <尾号>.2.jpg

运行: python build_v2.py
"""
import sys, io, os, re, shutil, zipfile, time, glob

if not getattr(sys.stdout, "_oa_wrapped", False):
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        sys.stdout._oa_wrapped = True
    except Exception:
        pass

import numpy as np
from PIL import Image
import pypdf
import cv2

T0 = time.time()

# ==================================================================
# CONFIG
# ==================================================================

INVOICE_DIRS = [
    r"D:\DeepSeekWorkspace\oa_v2\invoices",   # 专用归档（首选，别动）
    r"C:\Users\lenovo\Desktop",               # 桌面（用户常丢这）
    r"C:\Users\lenovo\Downloads",             # 下载目录
]
SCAN_PDF    = r"D:\微信聊天记录\xwechat_files\wxid_96ekyiryud0n22_117f\msg\file\2026-11\报销相关文件.pdf"
SIGN_DIR    = r"D:\DeepSeekWorkspace\newbatch_src"
TEMPLATE    = r"C:\Users\lenovo\.dsh\skills\henanshengyi-OA-reimburse\templates\验收单母版.docx"
PRINT_SRC   = r"C:\Users\lenovo\.dsh\skills\henanshengyi-OA-reimburse\scripts"
EXTRA_DIR   = r"C:\Users\lenovo\Downloads\OA"     # 2931 / 3821 的原件（电子发票 + iOS 扫描出库单）
OUT_DIR     = r"D:\DeepSeekWorkspace\FINAL_V2"

PROJECT_NAME = "2026年度省级留学人员科研择优资助"
PROJECT_CODE = "豫人社函〔2026〕5号"
HANDLER      = "牛科润"
ACCEPT_DATE  = "2026年11月16日"

# 这些尾号用各自的原始项目名/编码
OVERRIDE = {"1266": ("骨肉瘤在重离子放射治疗中的生物学效应研究", "ZC20260123")}

# 扫描件页面 -> (尾号, 类型, 旋转角)
# p2=1266出库单  p4=3271出库单  p6=4911出库单  p8=7021出库单
SLIP_PAGES = [
    (2, "1266", 0),
    (4, "3271", 0),
    (6, "4911", 0),   # 这张扫描时是横躺的，摆正
    (8, "7021", 0),
]

# 签字照 -> 尾号（按文件名排序，1-based）
# 已肉眼核对：p1=1266 p2=2931 p3=3271 p4=3821 p5=4911 p6=7021
SIGN_MAP = {1: "1266", 2: "2931", 3: "3271", 4: "3821", 5: "4911", 6: "7021"}

# 全部 6 组
CODES = ["1266", "2931", "3271", "3821", "4911", "7021"]

INVOICE_PT = (595.276, 396.85)
SLIP_PT    = (595.0, 842.0)

# ==================================================================
# 工具
# ==================================================================

def _norm(g):
    sw, sh = max(1, g.shape[1] // 48), max(1, g.shape[0] // 48)
    small = np.asarray(Image.fromarray(g.astype(np.uint8)).resize((sw, sh), Image.BOX)).astype(np.float32)
    bg = np.asarray(Image.fromarray(np.clip(small, 0, 255).astype(np.uint8))
                    .resize((g.shape[1], g.shape[0]), Image.BILINEAR)).astype(np.float32)
    return np.clip(g / np.maximum(bg, 1) * 235, 0, 255)


def _longest_run(mask, gap):
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


def content_box(im, rowf=0.004, colf=0.002, pad=0.008):
    g = np.asarray(im.convert("L")).astype(np.float32)
    ink = _norm(g) < 168
    H, W = ink.shape
    r0, r1 = _longest_run(ink.mean(axis=1) > rowf, int(H * 0.02))
    band = ink[r0:r1 + 1] if r1 > r0 else ink
    c0, c1 = _longest_run(band.mean(axis=0) > colf, int(W * 0.02))
    if r1 <= r0 or c1 <= c0:
        return (0, 0, W, H)
    px, py = int(W * pad), int(H * pad)
    return (max(0, c0 - px), max(0, r0 - py), min(W, c1 + px), min(H, r1 + py))


def make_slip_pdf(img, dst, dpi=300.0):
    """出库单 -> A4 竖版，保留整张纸（含下方空白）"""
    W = int(round(SLIP_PT[0] / 72.0 * dpi))
    H = int(round(SLIP_PT[1] / 72.0 * dpi))
    mw, mt, mr, mb = 0.010, 0.012, 0.010, 0.012
    aw, ah = W * (1 - mw - mr), H * (1 - mt - mb)
    s = min(aw / img.width, ah / img.height)
    nw, nh = max(1, int(img.width * s)), max(1, int(img.height * s))
    c = Image.new("RGB", (W, H), (255, 255, 255))
    c.paste(img.resize((nw, nh), Image.LANCZOS), (int(mw * W + (aw - nw) / 2), int(mt * H)))
    c.save(dst, "PDF", resolution=dpi, quality=95)


def warp_signature(img_bgr):
    hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
    L = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2LAB)[:, :, 0]
    s = hsv[:, :, 1]
    m = ((L > 140) & (s < 65)).astype(np.uint8) * 255
    m = cv2.morphologyEx(m, cv2.MORPH_CLOSE, np.ones((61, 61), np.uint8))
    m = cv2.morphologyEx(m, cv2.MORPH_OPEN, np.ones((61, 61), np.uint8))
    cnts, _ = cv2.findContours(m, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts:
        return None, 0, 0
    c = max(cnts, key=cv2.contourArea)
    if cv2.contourArea(c) < img_bgr.shape[0] * img_bgr.shape[1] * 0.3:
        return None, 0, 0
    rect = cv2.minAreaRect(c)
    pts = cv2.boxPoints(rect)
    ys = pts[:, 1]
    mid = (ys.min() + ys.max()) / 2
    top, bot = pts[ys <= mid], pts[ys > mid]
    if len(top) != 2 or len(bot) != 2:
        return None, 0, 0
    tl = top[np.argmin(top[:, 0])]
    tr = top[np.argmax(top[:, 0])]
    bl = bot[np.argmin(bot[:, 0])]
    br = bot[np.argmax(bot[:, 0])]
    quad = np.array([tl, tr, br, bl], dtype=np.float32)
    wtop = float(np.linalg.norm(quad[1] - quad[0]))
    hleft = float(np.linalg.norm(quad[3] - quad[0]))
    if wtop < hleft:
        quad = np.array([quad[1], quad[2], quad[3], quad[0]], dtype=np.float32)
    W = int(max(np.linalg.norm(quad[1] - quad[0]), np.linalg.norm(quad[2] - quad[3])))
    H = int(max(np.linalg.norm(quad[3] - quad[0]), np.linalg.norm(quad[2] - quad[1])))
    if W <= 0 or H <= 0:
        return None, 0, 0
    dst = np.array([[0, 0], [W - 1, 0], [W - 1, H - 1], [0, H - 1]], dtype=np.float32)
    M = cv2.getPerspectiveTransform(quad, dst)
    return cv2.warpPerspective(img_bgr, M, (W, H)), W, H


# ---------- docx ----------

def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def cell_text(cell):
    return "".join(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", cell, re.S))


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


def build_docx(src, dst, project, project_code, date, handler, seller,
               invoice_no, items, qty_total, amt_total):
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}
    xml = data["word/document.xml"].decode("utf-8")

    # ---- 段落字段：靠 label 定位下一个 w:t ----
    def set_after_label(x, label, newval):
        pat = re.compile(
            r"(" + re.escape(label) + r"</w:t></w:r>.*?<w:t(?:\s[^>]*)?>)([^<]*)(</w:t>)",
            re.S)
        return pat.sub(lambda m: m.group(1) + esc(newval) + m.group(3), x, count=1)

    xml = set_after_label(xml, "项目名称：", project)
    xml = set_after_label(xml, "项目编码：", project_code)
    xml = set_after_label(xml, "验收时间：", date)

    # ---- 表格 ----
    s = xml.index("<w:tbl>")
    e = xml.index("</w:tbl>") + len("</w:tbl>")
    tbl = xml[s:e]
    rows = re.findall(r"<w:tr\b.*?</w:tr>", tbl, re.S)
    hdr, colhdr, item_tpl, total_tpl = rows[0], rows[1], rows[2], rows[5]

    # 表头：母版结构 = 经办人|值|供应商|值|发票号码|值
    hc = re.findall(r"<w:tc>.*?</w:tc>", hdr, re.S)
    labels = [cell_text(c).strip() for c in hc]
    new_hc = []
    i = 0
    while i < len(hc):
        lbl = labels[i]
        if lbl in ("经办人", "供应商", "发票号码") and i + 1 < len(hc):
            val = {"经办人": handler, "供应商": seller, "发票号码": invoice_no}[lbl]
            new_hc.append(hc[i])
            new_hc.append(set_cell(hc[i + 1], val, force_sz=20 if lbl == "供应商" else None))
            i += 2
        else:
            new_hc.append(hc[i])
            i += 1
    hdr_new = re.sub(r"(<w:tr\b.*?>)(.*)(</w:tr>)",
                     lambda m: m.group(1) + "".join(new_hc) + m.group(3),
                     hdr, count=1, flags=re.S)

    # 明细行
    tpl_cells = re.findall(r"<w:tc>.*?</w:tc>", item_tpl, re.S)
    item_rows = []
    for idx, it in enumerate(items):
        cells = []
        for j, c in enumerate(tpl_cells):
            if j == len(tpl_cells) - 1:
                cells.append(c if idx == 0 else c.replace('<w:vMerge w:val="restart"/>', "<w:vMerge/>"))
            else:
                cells.append(set_cell(c, it[j] if j < len(it) else ""))
        item_rows.append(re.sub(r"(<w:tr\b.*?>)(.*)(</w:tr>)",
                                lambda m: m.group(1) + "".join(cells) + m.group(3),
                                item_tpl, count=1, flags=re.S))

    # 合计行
    tc = re.findall(r"<w:tc>.*?</w:tc>", total_tpl, re.S)
    tvals = [None, qty_total, None, amt_total, None]
    new_tc = [c if v is None else set_cell(c, v) for c, v in zip(tc, tvals)]
    total_new = re.sub(r"(<w:tr\b.*?>)(.*)(</w:tr>)",
                       lambda m: m.group(1) + "".join(new_tc) + m.group(3),
                       total_tpl, count=1, flags=re.S)

    tbl_new = (tbl[:tbl.index(rows[0])] + hdr_new + colhdr +
               "".join(item_rows) + total_new +
               tbl[tbl.index(rows[5]) + len(rows[5]):])
    xml = xml[:s] + tbl_new + xml[e:]

    data["word/document.xml"] = xml.encode("utf-8")
    tmp = dst + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, data[n])
    os.replace(tmp, dst)


# ---------- zip ----------

def make_zip(src_dir, zip_path, arc_root="OA"):
    import zlib
    entries = []
    for root, dirs, files in os.walk(src_dir):
        for f in sorted(files):
            full = os.path.join(root, f)
            rel = os.path.relpath(full, src_dir).replace("\\", "/")
            entries.append((arc_root + "/" + rel, full))

    chunks, central = [], []
    offset = 0
    for arc, full in entries:
        with open(full, "rb") as fh:
            raw = fh.read()
        comp = zlib.compressobj(9, zlib.DEFLATED, -15)
        packed = comp.compress(raw) + comp.flush()
        crc = zlib.crc32(raw) & 0xFFFFFFFF
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


# ==================================================================
# 数据（从票面 + 上次 docx 还原，已核对）
# ==================================================================

GROUPS = {
    "1266": dict(seller="上海冠路生物医药科技有限公司", invoice="26312000006900371266",
                 items=[["*化学试剂助剂*增强型ECL化学发光液试剂盒", "BF06053 250ml+250ml", "1", "¥900.00", "¥900.00"],
                        ["*化学试剂助剂*BeyoMag™ Protein A+G磁珠", "P2108-5ml 5ml", "1", "¥1,200.00", "¥1,200.00"]],
                 qty="2", total="¥2,100.00"),
    "3271": dict(seller="上海冠路生物医药科技有限公司", invoice="26312000006900453271",
                 items=[["*化学试剂助剂*BCA蛋白浓度测定试剂盒(增强型)", "P0009 5000次", "3", "¥1,726.00", "¥5,178.00"]],
                 qty="3", total="¥5,178.00"),
    "4911": dict(seller="上海懿贝瑞生物医药科技有限公司", invoice="26312000006900254911",
                 items=[["*化学试剂助剂*特级胎牛血清", "C04001-500 500ml", "2", "¥2,500.00", "¥5,000.00"],
                        ["*化学试剂助剂*OptiMEM I Reduced Serum Medium", "31985070 500mL", "2", "¥530.50", "¥1,061.00"]],
                 qty="4", total="¥6,061.00"),
    "7021": dict(seller="上海懿贝瑞生物医药科技有限公司", invoice="26312000006900187021",
                 items=[["*化学试剂助剂*Annexin V-APC / PI Apoptosis Detection Kit[217]", "E-CK-A217 100T", "3", "¥1,680.00", "¥5,040.00"],
                        ["*化学试剂助剂*无内毒素质粒小提中量试剂盒", "DP118-03 200次", "2", "¥1,340.00", "¥2,680.00"]],
                 qty="5", total="¥7,720.00"),
    "2931": dict(seller="青岛经唯生物技术有限公司", invoice="26922000001244622931",
                 items=[["ProteinA/G免疫(共)沉淀试剂盒", "2", "2", "¥2,110.00", "¥4,220.00"],
                        ["Phusion High-Fidelity DNA Polymerase", "1", "1", "¥3,400.00", "¥3,400.00"],
                        ["高纯度质粒小提试剂盒", "2", "2", "¥987.00", "¥1,974.00"]],
                 qty="5", total="¥9,594.00"),
    "3821": dict(seller="唯可（青岛）生物信息科技有限公司", invoice="26922000001244063821",
                 items=[["化学试剂助剂*Timentin", "瓶", "1", "¥3,400.00", "¥3,400.00"],
                        ["化学试剂助剂*卡纳霉素", "瓶", "3", "¥980.00", "¥2,940.00"],
                        ["化学试剂助剂*氨苄青霉素钠盐", "瓶", "2", "¥1,360.00", "¥2,720.00"],
                        ["化学试剂助剂*蛋白酶K", "瓶", "1", "¥940.00", "¥940.00"],
                        ["化学试剂助剂*LB Broth", "瓶", "1", "¥406.00", "¥406.00"]],
                 qty="8", total="¥10,406.00"),
}

# ==================================================================

PKG = os.path.join(OUT_DIR, "OA")
shutil.rmtree(OUT_DIR, ignore_errors=True)
os.makedirs(PKG, exist_ok=True)

print("=" * 62)
print("  OA 报销 V2 —— 发票电子原件版")
print("=" * 62)

# ---------- 1. 发票：电子原件，原样复制 ----------
print("\n[1/6] 放置电子发票原件...")
t = time.time()
inv_files = []
for d in INVOICE_DIRS:
    if os.path.isdir(d):
        inv_files += glob.glob(os.path.join(d, "dzfp_*.pdf"))
inv_by_tail = {}
for f in inv_files:
    m = re.search(r"dzfp_(\d{20})_", os.path.basename(f))
    if not m:
        continue
    tail = m.group(1)[-4:]
    # 同一尾号有多份时，优先取 oa_v2\invoices 下的
    if tail in inv_by_tail and "oa_v2" not in f:
        continue
    inv_by_tail[tail] = f

placed_inv = set()
for code in CODES:
    d = os.path.join(PKG, code)
    os.makedirs(d, exist_ok=True)
    dst = os.path.join(d, "%s.pdf" % code)
    src = inv_by_tail.get(code)
    if src is None:
        # 2931 / 3821 从 EXTRA_DIR 拿
        cand = os.path.join(EXTRA_DIR, "%s.pdf" % code)
        if os.path.exists(cand):
            src = cand
    if src is None:
        print("      !! %s 找不到电子发票原件" % code)
        continue
    shutil.copyfile(src, dst)
    placed_inv.add(code)
    print("      %-6s <- %s" % (code, os.path.basename(src)))
print("      %d/%d 份, %.2f s" % (len(placed_inv), len(CODES), time.time() - t))

# ---------- 2. 出库单：扫描件切页 ----------
print("\n[2/6] 生成出库单 PDF（扫描件）...")
t = time.time()
reader = pypdf.PdfReader(SCAN_PDF)
for page, code, rot in SLIP_PAGES:
    img = reader.pages[page - 1].images[0].image
    # 横躺的页需要摆正：看宽高比，A4 应该是竖的
    if img.width > img.height * 1.1:
        img = img.rotate(90, expand=True)
    d = os.path.join(PKG, code)
    os.makedirs(d, exist_ok=True)
    make_slip_pdf(img, os.path.join(d, "%s.1.pdf" % code))
    print("      %-6s p%d %dx%d" % (code, page, img.width, img.height))

# 2931 / 3821 出库单从 EXTRA_DIR 拷
for code in ["2931", "3821"]:
    src = os.path.join(EXTRA_DIR, "%s.1.pdf" % code)
    if os.path.exists(src):
        shutil.copyfile(src, os.path.join(PKG, code, "%s.1.pdf" % code))
        print("      %-6s <- 外部原件 %s.1.pdf" % (code, code))
    else:
        print("      !! %s 出库单缺失" % code)
print("      %.2f s" % (time.time() - t))

# ---------- 3. 签字照 ----------
print("\n[3/6] 处理签字照...")
t = time.time()
sign_files = sorted([f for f in os.listdir(SIGN_DIR) if f.lower().endswith((".png", ".jpg", ".jpeg"))])
ok_sign = 0
for idx, code in sorted(SIGN_MAP.items()):
    if idx > len(sign_files):
        continue
    src = os.path.join(SIGN_DIR, sign_files[idx - 1])
    img = cv2.imread(src)
    if img is None:
        print("      !! 读不了 %s" % src)
        continue
    warp, W, H = warp_signature(img)
    if warp is None:
        print("      !! 纸张检测失败 p%d" % idx)
        continue
    # 必须横版
    if H > W:
        warp = cv2.rotate(warp, cv2.ROTATE_90_CLOCKWISE)
        W, H = H, W
    d = os.path.join(PKG, code)
    os.makedirs(d, exist_ok=True)
    cv2.imwrite(os.path.join(d, "%s.2.jpg" % code), warp, [int(cv2.IMWRITE_JPEG_QUALITY), 95])
    ok_sign += 1
    print("      %-6s %dx%d" % (code, W, H))
print("      %d 张, %.2f s" % (ok_sign, time.time() - t))

# ---------- 4. 验收单 ----------
print("\n[4/6] 生成验收单 docx...")
t = time.time()
for code in CODES:
    g = GROUPS[code]
    proj, pcode = OVERRIDE.get(code, (PROJECT_NAME, PROJECT_CODE))
    dst = os.path.join(PKG, code, "%s（非统一采购）.docx" % code)
    build_docx(TEMPLATE, dst, proj, pcode, ACCEPT_DATE, HANDLER,
               g["seller"], g["invoice"], g["items"], g["qty"], g["total"])
    print("      %-6s %d 行" % (code, len(g["items"])))
print("      %d 份, %.2f s" % (len(CODES), time.time() - t))

# ---------- 5. 打印脚本 ----------
print("\n[5/6] 放入打印脚本...")

def copy_with_bom(src, dst):
    """PowerShell 5.1 会把无 BOM 的 UTF-8 当 GBK 读，中文全乱。
       .ps1 必须带 BOM，这里自动补上。"""
    with open(src, "rb") as fh:
        raw = fh.read()
    if src.lower().endswith(".ps1") and not raw.startswith(b"\xef\xbb\xbf"):
        raw = b"\xef\xbb\xbf" + raw
    with open(dst, "wb") as fh:
        fh.write(raw)
    return raw.startswith(b"\xef\xbb\xbf")

for f in ["OA-print.ps1", "OA-print.bat", "OA-run.bat", "pdf-print.ps1", "打印耗材验收单.ps1", "打印耗材验收单.bat"]:
    src = os.path.join(PRINT_SRC, f)
    if os.path.exists(src):
        ok = copy_with_bom(src, os.path.join(PKG, f))
        mark = " (BOM)" if ok else ""
        print("      " + f + mark)

# ---------- 6. 打包 ----------
print("\n[6/6] 打包...")
t = time.time()
zip_path = os.path.join(OUT_DIR, "OA.zip")
n, sz = make_zip(PKG, zip_path)
print("      %d 文件, %.2f MB, %.2f s" % (n, sz / 1024 / 1024, time.time() - t))

print()
print("=" * 62)
print("  完成！总耗时 %.2f 秒" % (time.time() - T0))
print("  " + zip_path)
print("=" * 62)
