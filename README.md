# henanshengyi-OA-reimburse

> 把一沓票据，自动整理成财务要的 OA.zip。

河南省人民医院（Henan Provincial People's Hospital）报销单据打包工具。
输入**电子发票原件** + 出库单扫描件 + 验收单模板，输出按发票尾号分组的完整报销包。

## 这东西解决什么问题

每次报销都要把手上一堆票据整理成固定格式：

```
Oa/
├── 1266/
│   ├── 1266.pdf                 # 发票（电子原件，原字节复制）
│   ├── 1266.1.pdf               # 出库单 / 清单（A4 竖版扫描件）
│   ├── 1266.2.jpg               # 验收单签字照
│   └── 1266（非统一采购）.docx   # 验收单
├── 3271/
└── ...
```

手工做一遍要把它裁正、调尺寸、填验收单、按号分组、打包……几十张票能耗掉一下午。这个工具把它变成一条命令，**1.6 秒出包**。

## 核心设计

### 发票必须是电子原件，扫描版会被打回

> **2026-11-16 规矩变更**

审核不认扫描版发票。发票 `<尾号>.pdf` 的**唯一合法来源是税局电子发票原件**（`dzfp_*.pdf`），找到就**原字节复制**——不裁边、不转图、不重编码。

```python
# 对：原样复制，零加工
shutil.copyfile(electronic_invoice_src, dst_invoice_pdf)

# 错：扫描件切页当发票 —— 审核直接拒绝
img = reader.pages[i].images[0].image   # 这只能用于出库单
```

判定是不是电子原件（三看）：

| 检查项 | 电子原件 | 扫描件 |
|---|---|---|
| `metadata /Producer` | `Suwell OFD convertor` | 无 |
| `metadata /Author` | `China Tax` | 无 |
| 页面尺寸 | 正好 595.276 × 396.85 pt | 尺寸各异 |

```python
md = dict(pypdf.PdfReader(f).metadata or {})
is_electronic = md.get("/Producer") == "Suwell OFD convertor"
```

### 出库单 / 清单从扫描件来

**绝不**把手机照片拼成 PDF。照片有桌面背景、纸张折痕、透视畸变，出来的东西不是扫描件，财务不认。

正确做法是从扫描 PDF 里提取**原始位图**，零重编码，只做旋转和版面摆放：

```python
import pypdf
reader = pypdf.PdfReader(scan_pdf)
img = reader.pages[i].images[0].image   # 原始像素，不损失
```

方向也不用逐页目测——出库单是 A4，宽高比就能判：

```python
if img.width > img.height * 1.1:   # 躺着 -> 转正
    img = img.rotate(90, expand=True)
```

### 页面规格对齐原件

| 文件 | 尺寸 | 方向 |
|---|---|---|
| 发票 `.pdf` | 595.276 × 396.85 pt | 横版（电子原件自带） |
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

## 脚本清单

| 文件 | 作用 |
|---|---|
| `build_v2_einvoice.py` | **一键构建（主力）**，发票走电子原件，1.6 秒出包 |
| `build_all.py` | 旧版，发票走扫描件（仅当拿不到电子原件时用） |
| `OA-print.bat` | **唯一打印入口**，双击进菜单 / 带参数走命令行 |
| `OA-print.ps1` | 统一打印逻辑 |
| `pdf-print.ps1` | 零依赖 PDF 打印模块（WinRT） |

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


## 打印

### 一个 bat，两种用法

`OA-print.bat` 是**唯一入口**，不带参数双击进菜单，带参数直接执行：

```powershell
# 双击 = 进菜单
# 或命令行带参数：
.\OA-print.bat -Mode Step2 -DryRun
```

### 按真实工作流分两步，别一次全打

```
第一步  打印验收单 -> 手写签名 -> 拍照发回打包
第二步  打包完成后再打印发票 + 出库单/清单
```

**双击 `OA-print.bat`**，进菜单：

```
  第一步  打印验收单 -> 手写签名 -> 拍照回传打包
  第二步  打包完成后再打印发票 + 出库单/清单

   [1] 第一步：打印全部验收单（docx）
   [2] 第二步：打印全部发票 + 出库单
   [3] 单独重打某一组（按尾号）
   [4] 全部一次打完（一般不用）
   [D] 诊断模式
   [Q] 退出
```

选 `[3]` 可以按尾号单独重打——打废了、漏签了用这个。

### 命令行参数

```powershell
cd <OA目录>

.\OA-print.bat -Mode Step2 -DryRun              # 预览发票+出库单
.\OA-print.bat -Mode Step1                      # 真打验收单
.\OA-print.bat -Mode Step2                      # 真打发票+出库单
.\OA-print.bat -Mode Group -Group 1266 -DryRun  # 预览某一组
.\OA-print.bat -Mode Diagnose                   # 环境诊断
```

### 为什么只能是 .bat，不能是 .ps1

本机执行策略会拦未签名脚本，直接 `.\OA-print.ps1` 会报：

```
无法加载文件 ... 未对文件进行数字签名。
```

**必须走 bat**，因为 bat 里带了 `-ExecutionPolicy Bypass`：

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File "%PSFILE%" %*
```

### 打印顺序：先进后出

打印机出纸是后打的落上面：

- **组内**：先送出库单，后送发票 → 拿到手最上面是发票
- **组间**：尾号倒序（7021 → ... → 1266）→ 最上面是 1266 那组

### PDF 打印零依赖

**很多机器没装 PDF 阅读器**，`.pdf` 连文件关联都没有，`Start-Process -Verb Print` 直接失败。

这里用 Windows 自带的 `Windows.Data.Pdf`（WinRT）把 PDF 渲染成位图，再走 `System.Drawing.Printing` 送打印队列——**不装任何东西**。

探测顺序：SumatraPDF → Acrobat → Reader → Foxit → Edge/Chrome → WinRT 渲染 → Shell。

### 兼容性

- **Word/WPS 自适应**：`Word.Application` → `KWPS.Application` → `WPS.Application` → Shell 兜底
- **路径无关**：换电脑、换盘符、换目录名都能用，脚本以自己所在位置为根
- **自动跳过锁文件**：Word 的 `~$xxx.docx` 不会被打印

### 出问题时：日志

每次运行都会在 OA 根目录生成 `打印日志_YYYYMMDD_HHMMSS.txt`。报错时把这个文件发出来就能定位。

里面记录了环境快照、默认打印机、找到的文件清单、每种引擎探测结果、逐份耗时和失败原因。

先跑诊断（不打印，只出日志）：

```powershell
.\OA-run.bat -Mode Diagnose
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


## 速度

整条流水线 **1.6 秒**跑完：

```
[1/6] 放置电子发票原件      0.01 s
[2/6] 生成出库单 PDF        1.0 s
[3/6] 处理签字照            0.35 s
[4/6] 生成验收单 docx       0.02 s
[5/6] 放入打印脚本          0.0 s
[6/6] 打包                  0.25 s
```

**瓶颈从来不在计算，在来回确认。** 所以流程设计成：

1. 开工前一次性问齐所有信息（项目名/编码/经办人/路径）
2. 看图用拼图（6 张合成 1 张），不要逐张
3. **能代码跑的一律代码跑**——字段核对、金额校验、发票验真、方向判断、最终自检，全走程序断言
4. 一次跑完，中间不打断

### 一键构建

```bash
cd scripts
# 改 build_v2_einvoice.py 里的 CONFIG 段
python build_v2_einvoice.py
```

输出 `OA.zip` + 完整控制台日志，**49 项自检全绿才算完**。

## 技术细节

### bat 内容必须纯 ASCII

**中文注释会把 bat 拆坏。** 实测在 bat 里写中文 `REM` 注释后，cmd 会把 UTF-8 字节按 GBK 切开，报出 `'un.bat' is not recognized` 这种诡异错误。

```bat
REM 对：注释全英文
REM  Print docx invoices in reverse order

REM 错：中文注释，cmd 解析会崩
REM  打印验收单
```

`chcp 65001` 解决不了。文件名可以是中文，**内容必须纯 ASCII**。

### PowerShell 脚本必须带 UTF-8 BOM

`powershell.exe` 5.1 遇到无 BOM 的 UTF-8 文件会按 GBK 解码，中文字符串全乱、语法直接报错。

构建器里自动补 BOM：

```python
def copy_with_bom(src, dst):
    with open(src, "rb") as fh:
        raw = fh.read()
    if src.lower().endswith(".ps1") and not raw.startswith(b"\xef\xbb\xbf"):
        raw = b"\xef\xbb\xbf" + raw
    with open(dst, "wb") as fh:
        fh.write(raw)
```

### WinRT 打印 PDF 的两个坑

```powershell
# 坑 1：IAsyncAction 的 AsTask 是【非泛型】方法
#   ❌ $m.MakeGenericMethod()  -> "is not a GenericMethodDefinition"
#   ✅ 直接 Invoke
$asTaskAction = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object { $_.Name -eq 'AsTask' -and -not $_.IsGenericMethod -and
                   $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' }

# 坑 2：PdfDocument 没有 Dispose 方法
#   ❌ $pdf.Dispose()  -> "does not contain a method named 'Dispose'"
#   ✅ 不用管，交给 GC
```

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

Windows 下 zip 中文名容易乱码。`make_zip` 内部手写 zip 结构，所有条目强制设 UTF-8 标志位（0x800），解压不乱码。

## 注意事项

- **发票必须是电子原件**：没有就停下来要，别拿扫描版凑合
- **只处理未压缩的 PDF**：扫描件里的图必须是 JPEG/PNG，不能是 JBIG2/CCITT
- **一页一张票**：一页多票需要先切分
- **验收单母版不可删**：所有验收单都从 `templates/验收单母版.docx` 克隆

## 更新记录

- 2026-11-05：首次发布
- 2026-11-08：一键流水线（3.7s）、签字照拼图预览、开工前问齐纪律
- 2026-11-16：**发票改用电子原件**（扫描版审核不通过）、**能代码跑的一律代码跑**、两步式打印工作流、零依赖 PDF 打印（WinRT）、执行策略与 BOM/ASCII 规则

## License

MIT
