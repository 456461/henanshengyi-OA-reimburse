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


## 批量打印

打开 OA 文件夹，**双击 `打印耗材验收单.bat`**，就会把每个子文件夹里的验收单全部送进打印队列，不用一个个点。

- **路径无关**：换电脑、换盘符、换目录名都能用，脚本以自己所在位置为根
- **自动跳过锁文件**：Word 的 `~$xxx.docx` 临时文件不会被打印
- **打印前确认**：列出所有文档 + 显示默认打印机，按 Y 才打

命令行细化控制：

```powershell
# 只列清单，不打印（先核对）
.\打印耗材验收单.bat        # 双击后手动加参数不方便，推荐直接跑 ps1
powershell -ExecutionPolicy Bypass -File .\打印耗材验收单.ps1 -DryRun

# 跳过确认直接打
powershell -ExecutionPolicy Bypass -File .\打印耗材验收单.ps1 -Force
```

**兼容性**：自动适配 Microsoft Word 和 WPS，不需要手动选。

探测顺序（逐级降级）：

1. `Word.Application` —— 新版 WPS 也注册了这个 ProgID，先试它
2. `KWPS.Application` / `WPS.Application` —— WPS 原生接口
3. **Shell 兜底** —— 交给系统 `.docx` 文件关联，装的是什么就用什么

三种都不可用时会明确报错，不会静默失败。

强制指定引擎（一般用不上）：

```powershell
# 只在特殊排查时用
powershell -ExecutionPolicy Bypass -File .\打印耗材验收单.ps1 -Engine WPS
```


## 自动填写的字段

| 字段 | 规则 |
|---|---|
| 验收时间 | **默认当天日期**（`2026年11月8日` 这种格式，月日不补零） |
| 供应商 | 从发票「销售方」读 |
| 发票号码 | 从发票读 |
| 明细（名称/规格/数量/单价/金额） | 从发票逐行读，**含税价** |
| 数量合计 / 金额合计 | 自动求和校验 |

以下字段**不恒定**，每次由用户指定：项目名称、项目编码、经办人。


## 出问题时：日志

每次运行都会在 OA 根目录生成 `打印日志_YYYYMMDD_HHMMSS.txt`。

**报错时把这个文件发出来就能定位问题**，里面记录了：

- 环境快照（PowerShell 版本、Office 位数、执行策略、用户名、计算机名）
- 默认打印机
- 找到的每个 docx 文件
- 每种打印引擎的探测结果（Word / WPS / Shell 哪个能用、哪个不能用、为什么）
- 逐份打印结果
- 失败时的完整异常（类型、消息、行号、调用栈）

### 先跑诊断模式

遇到报错，先执行这条（不打印，只出日志）：

```
打印耗材验收单.bat -Diagnose
```

然后把生成的 `打印日志_*.txt` 发出去。

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

## 更新记录

- 2026-11-05：首次发布

## License

MIT
