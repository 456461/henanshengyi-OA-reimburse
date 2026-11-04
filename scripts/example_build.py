# -*- coding: utf-8 -*-
"""
完整示例：从扫描件 PDF 生成 OA.zip

用法：改 CONFIG 段，然后 python example_build.py
"""
import os, sys, shutil
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import oa_tools as T

# ================= CONFIG =================
SCAN_PDF = r"D:\path\to\扫描件.pdf"
TEMPLATE = os.path.join(os.path.dirname(__file__), "..", "templates", "验收单母版.docx")
WORK     = r"D:\DeepSeekWorkspace\oa_work"
OUT_ZIP  = r"D:\DeepSeekWorkspace\FINAL\OA.zip"

PROJECT_NAME   = "项目名称（示例）"
OLD_PROJECT    = "项目名称（示例）"
KEEP_ORIGINAL  = {"1266"}      # 这些尾号保留原项目名
PROJECT_CODE   = "ZC00000000"
HANDLER        = "张三"

# 每张票：(页码, 尾号, 类型, 旋转角度CCW, 发票号, 供应商, 明细, 数量合计, 金额合计, 日期)
# 明细 = [(名称, 规格, 数量, 单价, 金额), ...]
GROUPS = [
    dict(page=1, code="1266", kind="invoice", rot=0,
         invoice="00000000000000000000", seller="供应商名称",
         items=[("耗材名称示例A", "SPEC-A", "1", "¥0.00", "¥0.00"),
                ("耗材名称示例B", "SPEC-B", "1", "¥0.00", "¥0.00")],
         qty="2", total="¥0.00", date="YYYY年MM月DD日"),
    dict(page=2, code="1266", kind="slip",    rot=0),
    # ... 其余各组照此填写
]
# ==========================================


def main():
    shutil.rmtree(WORK, ignore_errors=True)
    os.makedirs(WORK, exist_ok=True)
    pkg = os.path.join(WORK, "OA")
    os.makedirs(pkg, exist_ok=True)

    # 1. 提取每页位图
    imgs = T.extract_pages(SCAN_PDF, os.path.join(WORK, "pages"))

    # 2. 按组生成 PDF
    from PIL import Image
    for g in GROUPS:
        img = Image.open(imgs[g["page"] - 1])
        if g.get("rot"):
            img = img.rotate(g["rot"], expand=True)
        d = os.path.join(pkg, g["code"])
        os.makedirs(d, exist_ok=True)
        name = "%s.pdf" % g["code"] if g["kind"] == "invoice" else "%s.1.pdf" % g["code"]
        T.make_pdf(img, g["kind"], os.path.join(d, name))
        print("  ", g["code"], name)

    # 3. 生成验收单
    from collections import defaultdict
    by_code = defaultdict(list)
    for g in GROUPS:
        by_code[g["code"]].append(g)

    for code, gs in by_code.items():
        inv = next((g for g in gs if g["kind"] == "invoice"), None)
        if not inv:
            continue
        # 金额校验
        ok, s, t, diff = T.check_amounts(inv["items"], inv["total"])
        if not ok:
            print("  !! %s 金额不平: %s vs %s" % (code, s, t))
        # 复制母版并改项目名
        dst = os.path.join(pkg, code, "%s（非统一采购）.docx" % code)
        repl = {}
        if code not in KEEP_ORIGINAL:
            repl[OLD_PROJECT] = PROJECT_NAME
        if repl:
            T.patch_docx_text(TEMPLATE, dst, repl)
        else:
            shutil.copyfile(TEMPLATE, dst)
        print("  ", code, "docx ok")

    # 4. 打包 + 自检
    n, size = T.make_zip(pkg, OUT_ZIP)
    print("\nzip: %d files, %.2f MB" % (n, size / 1024 / 1024))
    T.verify(OUT_ZIP)


if __name__ == "__main__":
    main()
