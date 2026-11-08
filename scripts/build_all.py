# -*- coding: utf-8 -*-
"""
OA 报销一键构建 —— 只改 CONFIG 段，一次跑完

运行: python build_all.py
"""
import sys, io, os, re, shutil, json, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 只在未包装时包装一次，避免 "I/O operation on closed file"
if not getattr(sys.stdout, "_oa_wrapped", False):
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        sys.stdout._oa_wrapped = True
    except Exception:
        pass

from PIL import Image
import cv2
import pypdf
import oa_pipeline as P

T0 = time.time()

# ==================================================================
# CONFIG —— 每批只改这里
# ==================================================================

SCAN_PDF  = r"D:\微信聊天记录\xwechat_files\wxid_96ekyiryud0n22_117f\msg\file\2026-11\报销相关文件.pdf"
SIGN_DIR  = r"D:\DeepSeekWorkspace\newbatch_src"          # 签字照目录（英文路径）
EXTRA_PDF_DIR = r"C:\Users\lenovo\Downloads\OA"   # 扫描件里没有的组，从这里拷原件（留空则跳过）
TEMPLATE  = r"D:\DeepSeekWorkspace\oa_clean\OA\2931（非统一采购）.docx"
OUT_DIR   = r"D:\DeepSeekWorkspace\FINAL"

PROJECT_NAME = "2026年度省级留学人员科研择优资助"
PROJECT_CODE = "豫人社函〔2026〕5号"
ACCEPT_DATE  = "2026年11月8日"
KEEP_ORIGINAL_NAME = {"1266"}          # 这些尾号保留原项目名
KEEP_ORIGINAL_CODE = {"1266": "ZC20260123"}

# 扫描件页面 -> (尾号, 类型, 旋转CCW)
PAGES = [
    (1, "1266", "invoice",  0),
    (2, "1266", "slip",     0),
    (3, "3271", "invoice",  0),
    (4, "3271", "slip",     0),
    (5, "4911", "invoice", 90),
    (6, "4911", "slip",    90),
    (7, "7021", "invoice", 90),
    (8, "7021", "slip",     0),
]

# 签字照序号(1-based, 按文件名排序) -> 尾号
SIGN_MAP = {1: "1266", 2: "2931", 3: "3271", 4: "3821", 5: "4911", 6: "7021"}

# 每组数据
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

print("=" * 60)
print("  OA 报销 一键构建")
print("=" * 60)

# ---------- 1. 提取扫描件 ----------
print("\n[1/6] 提取扫描件...")
t = time.time()
reader = pypdf.PdfReader(SCAN_PDF)
page_imgs = {}
for page, code, kind, rot in PAGES:
    if page not in page_imgs:
        img = reader.pages[page-1].images[0].image
        if rot: img = img.rotate(rot, expand=True)
        page_imgs[page] = img
print("      %d 页, %.2f s" % (len(page_imgs), time.time()-t))

# ---------- 2. 生成 PDF ----------
print("\n[2/6] 生成发票/出库单 PDF...")
t = time.time()
for page, code, kind, rot in PAGES:
    d = os.path.join(PKG, code)
    os.makedirs(d, exist_ok=True)
    name = "%s.pdf" % code if kind == "invoice" else "%s.1.pdf" % code
    P.make_pdf(page_imgs[page], kind, os.path.join(d, name))
print("      %d 个 PDF, %.2f s" % (len(PAGES), time.time()-t))

# ---------- 2.5 补齐外部原件 ----------
missing = [c for c in GROUPS if not os.path.exists(os.path.join(PKG, c, "%s.pdf" % c))]
if missing:
    print("\n[2.5] 补齐外部原件（扫描件里没有）...")
    t = time.time()
    for code in missing:
        d = os.path.join(PKG, code)
        os.makedirs(d, exist_ok=True)
        ok = False
        for suffix in [".pdf", ".1.pdf"]:
            src = os.path.join(EXTRA_PDF_DIR, code + suffix)
            if os.path.exists(src):
                shutil.copyfile(src, os.path.join(d, code + suffix))
                ok = True
        print("      %-6s %s" % (code, "已拷贝" if ok else "!! 源文件不存在: " + EXTRA_PDF_DIR))
    print("      %.2f s" % (time.time()-t))

# ---------- 3. 签字照 ----------
print("\n[3/6] 处理签字照...")
t = time.time()
sign_files = sorted([f for f in os.listdir(SIGN_DIR) if f.lower().endswith((".png",".jpg",".jpeg"))])
warped_for_sheet = []
for idx, code in sorted(SIGN_MAP.items()):
    if idx > len(sign_files): continue
    src = os.path.join(SIGN_DIR, sign_files[idx-1])
    img = cv2.imread(src)
    if img is None:
        print("      !! 读不了 %s" % src); continue
    warp, W, H = P.warp_signature(img)
    if warp is None:
        print("      !! 纸张检测失败 p%d" % idx); continue
    d = os.path.join(PKG, code)
    os.makedirs(d, exist_ok=True)
    dst = os.path.join(d, "%s.2.jpg" % code)
    cv2.imwrite(dst, warp, [int(cv2.IMWRITE_JPEG_QUALITY), 95])
    warped_for_sheet.append((code, dst))
    print("      %-6s %dx%d" % (code, W, H))

# 拼图
sheet = os.path.join(OUT_DIR, "_签字照拼图.jpg")
P.make_contact_sheet([p for _, p in warped_for_sheet], sheet)
print("      拼图: %s" % sheet)
print("      共 %d 张, %.2f s" % (len(warped_for_sheet), time.time()-t))

# ---------- 4. 验收单 ----------
print("\n[4/6] 生成验收单 docx...")
t = time.time()
for code, g in GROUPS.items():
    proj = ("骨肉瘤在重离子放射治疗中的生物学效应研究" if code in KEEP_ORIGINAL_NAME else PROJECT_NAME)
    pcode = KEEP_ORIGINAL_CODE.get(code, PROJECT_CODE)
    dst = os.path.join(PKG, code, "%s（非统一采购）.docx" % code)
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    P.build_docx(TEMPLATE, dst, proj, pcode, ACCEPT_DATE,
                 g["seller"], g["invoice"], g["items"], g["qty"], g["total"])
    print("      %-6s %d 行明细" % (code, len(g["items"])))
print("      %d 份, %.2f s" % (len(GROUPS), time.time()-t))

# ---------- 5. 打印脚本 ----------
print("\n[5/6] 放入打印脚本...")
t = time.time()
sk_dir = os.path.dirname(os.path.abspath(__file__))
for f in ["打印耗材验收单.ps1", "打印耗材验收单.bat"]:
    src = os.path.join(sk_dir, f)
    if os.path.exists(src):
        shutil.copyfile(src, os.path.join(PKG, f))
        print("      " + f)
print("      %.2f s" % (time.time()-t))

# ---------- 6. 打包 ----------
print("\n[6/6] 打包...")
t = time.time()
zip_path = os.path.join(OUT_DIR, "OA.zip")
n, sz = P.make_zip(PKG, zip_path)
print("      %d 文件, %.2f MB, %.2f s" % (n, sz/1024/1024, time.time()-t))

print()
print("=" * 60)
print("  完成！总耗时 %.2f 秒" % (time.time()-T0))
print("  " + zip_path)
print("=" * 60)
