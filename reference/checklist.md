# 交付前自检清单

复制这份清单，逐项打勾。**全部通过才能交付。**

## 一、ZIP 完整性

```python
import zipfile
with zipfile.ZipFile("OA.zip") as z:
    assert z.testzip() is None, "zip damaged"
    print(len(z.namelist()), "files")
```

- [ ] `testzip()` 返回 None
- [ ] 文件数 == 组数 x 3

## 二、PDF 规格

```bash
for f in OA/*/*.pdf; do
  echo "=== $f ==="
  pdfinfo "$f" | grep -E "Pages|Page size"
done
```

- [ ] 每个 PDF `Pages: 1`
- [ ] 发票 `*.pdf`（不带 .1）→ 595 x 397 左右，**横版**
- [ ] 出库单 `*.1.pdf` → 595 x 842 左右，**竖版**

## 三、发票号一致性

每组 docx 里的发票号，必须和文件夹名一致。

```python
import zipfile, re, os
for code in os.listdir("OA"):
    d = os.path.join("OA", code)
    if not os.path.isdir(d):
        continue
    docx = [f for f in os.listdir(d) if f.endswith(".docx")][0]
    with zipfile.ZipFile(os.path.join(d, docx)) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    inv = re.search(r"<w:t[^>]*>(\\d{20})</w:t>", xml)
    assert inv, code + " no invoice number"
    assert inv.group(1).endswith(code), code + " mismatch: " + inv.group(1)
    print(code, "OK")
```

- [ ] 每组发票号尾 4 位 == 文件夹名

## 四、项目名称

- [ ] 该改的组已改成新名称
- [ ] 该保留的组仍是原名称
- [ ] 全文没有残留旧名称（除保留组）

```python
NEW = "项目名称（示例）"
OLD = "项目名称（示例）"
KEEP = {"1266"}   # 保留原名的组

for code in sorted(os.listdir("OA")):
    d = os.path.join("OA", code)
    if not os.path.isdir(d):
        continue
    docx = [f for f in os.listdir(d) if f.endswith(".docx")][0]
    with zipfile.ZipFile(os.path.join(d, docx)) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    has_new, has_old = NEW in xml, OLD in xml
    if code in KEEP:
        assert has_old and not has_new, code + " should keep old"
    else:
        assert has_new and not has_old, code + " should use new"
    print(code, "OK")
```

## 五、金额校验

```python
from oa_tools import check_amounts

items = [
    ("耗材名称示例A", "SPEC-A", "1", "900.00", "900.00"),
    ("BeyoMag Protein A+G磁珠", "SPEC-B", "1", "1200.00", "1200.00"),
]
ok, s, t, diff = check_amounts(items, "2100.00")
assert ok, "amount mismatch: %s vs %s" % (s, t)
```

- [ ] 每行 单价 x 数量 == 金额
- [ ] 各行金额之和 == 票面价税合计

## 六、肉眼复核 docx

**必须做。** 用 Word 打开验收单，渲染成图片看一眼：

```powershell
$w = New-Object -ComObject Word.Application
$w.Visible = $false
Get-ChildItem OA\*\*.docx | ForEach-Object {
  $d = $w.Documents.Open($_.FullName, $false, $true)
  $d.ExportAsFixedFormat($_.FullName + '.pdf', 17)
  $d.Close(0)
}
$w.Quit()
```

- [ ] 表格边框完整，没有断线
- [ ] 字体没变（宋体/仿宋）
- [ ] 项目名称那行没溢出、没换行错乱
- [ ] 明细行数正确（几张票就几行）
- [ ] 签名栏留空
- [ ] 横向 A4 版式，不是竖的

## 七、肉眼复核 PDF

- [ ] 发票是正的，字不是倒的/躺着的
- [ ] 发票号清晰可读
- [ ] 出库单标题在上方，不是侧着
- [ ] 没有桌面背景、没有其他纸张边缘
- [ ] 四角没有大块黑边

## 常见故障对照

| 症状 | 原因 | 修法 |
|---|---|---|
| 发票是竖的 | 旋转角度给错 | 改 rot 参数，正负90 或 180 |
| 边缘有黑框 | 裁切没去干净 | 调大 content_box 的 pad |
| 内容被裁掉 | 包围盒太小 | 调小 rowf/colf |
| 出库单上半截没了 | 用了发票的裁切逻辑 | slip 类型不裁切 |
| docx 打不开 | zip 打包损坏 | 用 patch_docx_text，别手动重打 |
| 中文文件名乱码 | 编码问题 | make_zip 已处理，勿改 |
