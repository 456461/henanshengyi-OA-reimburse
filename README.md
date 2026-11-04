# henanshengyi-OA-reimburse

> 把一沓扫描件票据，自动整理成财务要的 OA.zip。

河南省人民医院（Henan Provincial People's Hospital）报销单据打包工具。
输入发票/出库单扫描件 PDF + 验收单模板，输出按发票尾号分组的完整报销包。

## 这东西解决什么问题

每次报销都要把手上一堆票据整理成固定格式：

```
Oa/
├── 1266/
│   ├── 1266.pdf                 # 发票（横版扫描件）
│   ├── 1266.1.pdf               # 出库单（A4 竖版）
│   └── 1266（非统一采购）.docx   # 验收单
├── 3271/
└── ...
```

手工做一遍要把它裁正、调尺寸、填验收单、按号分组、打包……几十张票能耗掉一下午。这个工具把它变成一条命令。

## 核心设计

### PDF 只从扫描件来

**绝不**把手机照片拼成 PDF。照片有桌面背景、纸张折痕、透视畸变，出来的东西不是扫描件，财务不认。

正确做法是从扫描 PDF 里提取**原始位图**，零重编码，只做旋转和版面摆放：

```python
import pypdf
reader = pypdf.PdfReader(scan_pdf)
img = reader.pages[i].images[0].image   # 原始像素，不损失
```

### 方向必须人工确认

试过自动判定方向：投影方差、自相关、傅里叶……**全都不靠谱**。90° 和 270° 的投影方差完全一样，算法无法区分"正着"和"倒着"。

唯一可靠的办法：把每页按候选角度渲染成 PNG，一张张用眼睛看。确认发票号读得对、标题朝向对、字不是倒的。

### 页面规格对齐原件

| 文件 | 尺寸 | 方向 |
|---|---|---|
| 发票 `.pdf` | 595.276 × 396.85 pt | 横版 |
| 出库单 `.1.pdf` | 595 × 842 pt | A4 竖版 |

跟单位 OA 系统里的原件一模一样，交上去不会被退。

### docx 只改文本

验收单是克隆母版改的。只替换 `word/document.xml` 里的文本节点，表格边框、字体、字号、页边距、横向节属性**一个字不动**。

### 金额必须含税

发票印的是不含税金额，验收单填含税金额。工具内置校验：

```
各行(单价 × 数量) == 各行金额
Σ各行金额 == 票面「价税合计」
```

对不上直接报错，不会悄悄生成错文件。

## 安装

```bash
pip install pypdf pillow numpy
```

## 使用

### 作为 AI Skill 使用

把整个目录放进 `~/.dsh/skills/` 或 `~/.claude/skills/`，然后直接对 AI 说：

> 帮我把这批票据打包成 OA.zip

AI 会自动加载这个 skill 的流程。

### 作为 Python 库使用

```python
import oa_tools as T

# 1. 提取扫描件每页位图
imgs = T.extract_pages("扫描件.pdf", "work/pages")

# 2. 生成发票 PDF（横版）
from PIL import Image
img = Image.open(imgs[0])
T.make_pdf(img, "invoice", "1266.pdf")

# 3. 生成出库单 PDF（A4 竖版）
img2 = Image.open(imgs[1])
T.make_pdf(img2, "slip", "1266.1.pdf")

# 4. 改验收单项目名称
T.patch_docx_text("母版.docx", "1266（非统一采购）.docx", {
    "旧项目名称": "2026年度省级留学人员科研择优资助"
})

# 5. 打包
T.make_zip("OA", "OA.zip")

# 6. 自检
T.verify("OA.zip")
```

完整示例见 `scripts/example_build.py`。

## API

| 函数 | 作用 |
|---|---|
| `extract_pages(pdf, dir)` | 无损提取每页位图 |
| `content_box(img)` | 检测票据内容包围盒，抗折痕干扰 |
| `make_pdf(img, kind, dst)` | 摆成发票/出库单规格 |
| `patch_docx_text(src, dst, {old:new})` | 改 docx 文本，保留样式 |
| `make_zip(src_dir, zip)` | 打包，UTF-8 文件名 |
| `check_amounts(items, total)` | 金额校验 |
| `verify(zip)` | 页面尺寸自检 |

## 技术细节

### 抗折痕的内容包围盒

直接算墨迹包围盒会被纸张折痕、污渍拉长。这里用**最长连续段**：

```python
def _longest_run(mask, gap):
    # 只保留最长的一段连续 True，跳过孤立干扰
```

折痕通常只占几行，构不成"最长连续段"，自然被忽略。

### 光照归一化

手机扫描往往中间亮边缘暗。用大核高斯模糊估背景，再做除法：

```python
bg = GaussianBlur(gray, sigma=W/22)
normalized = clip(gray / bg * 235, 0, 255)
```

### UTF-8 文件名打包

Windows 下 zip 中文名容易乱码。`make_zip` 内部用 `zipfile` 写入 UTF-8 标记，解压不乱码。

## 注意事项

- **只处理未压缩的 PDF**：扫描件里的图必须是 JPEG/PNG，不能是 JBIG2/CCITT
- **一页一张票**：一页多票需要先切分
- **验收单母版不可删**：所有验收单都从 `templates/验收单母版.docx` 克隆

## License

MIT
