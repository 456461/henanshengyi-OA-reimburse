---
name: henanshengyi-OA-reimburse
description: 河南省人民医院 OA 报销单据打包。把扫描件发票/出库单 PDF + 验收单 docx 整理成按发票尾号分组的 OA.zip。触发词：OA报销、报销打包、验收单、非统一采购、耗材验收单、发票出库单整理。
---

# 河南省人民医院 OA 报销打包

把一堆扫描件票据整理成交付给财务的 `OA.zip`。每组按发票尾号一个文件夹，内含发票、出库单、验收单三件套。

## ⚡ 执行纪律：先问齐，再动手，一次跑完

**这个 skill 最大的耗时不在计算，在来回确认。**

实测：整条流水线跑完只要 **3.7 秒**（提取 0.6s + 出 PDF 1.4s + 签字照 0.4s + 验收单 0.02s + 打包 1.2s）。
但人工一张张看图、反复渲染验证，能耗掉几十分钟。

### 硬规则 1：开工前一次性问齐

**收到材料后，先问清楚下面全部内容，问完再动手。不许边做边问。**

| # | 必问项 | 说明 |
|---|---|---|
| 1 | **扫描件 PDF 路径** | 发票+出库单的扫描件（不是照片） |
| 2 | **签字照目录** | 打印签字后拍的照片 |
| 3 | **项目名称** | 不恒定，每批可能不同 |
| 4 | **项目编码** | 不恒定 |
| 5 | **经办人** | 不恒定 |
| 6 | **是否有多个项目混批** | 有的话要逐份指定名称+编码 |

**验收时间不用问** —— 默认取当天。

问完立刻整理成 CONFIG，一次跑完，中间不打断用户。

### 硬规则 2：看图用拼图，不要一张张看

**这是提速的关键。** 6 张签字照逐张读 = 6 次往返，每次几十秒。
拼成一张 = 1 次往返，看清全部。

```python
def make_contact_sheet(images, dst, cols=3, cell_w=620):
    ims = [Image.open(p) if isinstance(p, str) else p for p in images]
    rows = (len(ims) + cols - 1) // cols
    ch = int(cell_w * ims[0].height / max(1, ims[0].width))
    sheet = Image.new('RGB', (cols*cell_w, rows*ch), 'white')
    for i, im in enumerate(ims):
        sheet.paste(im.resize((cell_w, ch), Image.LANCZOS), ((i%cols)*cell_w, (i//cols)*ch))
    sheet.save(dst, quality=82)
```

6 张拼成 1860x754 的图，174 KB，**一次看清全部发票号和签名**。

同理：**验收单 docx 也不要点开逐个检查**，用程序校验（读 `word/document.xml` 比对字段），
只有全部通过后才考虑抽样渲染 1-2 张肉眼确认。

### 硬规则 3：用一键流水线，不要分步执行

`scripts/build_all.py` 已把全部步骤串好，改完 CONFIG 直接跑：

```bash
cd scripts
python build_all.py
```

输出：

```
[1/6] 提取扫描件...        8 页, 0.6 s
[2/6] 生成发票/出库单...    8 个 PDF, 1.4 s
[2.5] 补齐外部原件...       （扫描件里没有的组）
[3/6] 处理签字照...        6 张 + 拼图, 0.4 s
[4/6] 生成验收单...        6 份, 0.02 s
[5/6] 放入打印脚本...      0.0 s
[6/6] 打包...              26 文件, 11 MB, 1.2 s

完成！总耗时 3.69 秒
```

### 硬规则 4：OpenCV 读不了中文路径

`cv2.imread()` 在 Windows 上遇到中文路径直接返回 None，**且不报错**，只会静默失败。

```python
# 错：静默返回 None
img = cv2.imread(r'c:\Users\xxx\Desktop\第1组\0001.png')   # None!

# 对：先复制到纯英文路径
import shutil
shutil.copyfile(src_cn_path, r'D:\work\p1.png')
img = cv2.imread(r'D:\work\p1.png')
```

### 时间预算

| 环节 | 应耗时 |
|---|---|
| 问齐信息 | 1 次往返 |
| 看图核对 | 1-2 次往返（拼图） |
| 跑流水线 | 4 秒 |
| 程序校验 | 1 次 |
| **合计** | **3-5 次往返** |

如果发现自己要读第 3 张以上的单张图，**停下来，改用拼图**。

## 何时用这个 Skill

用户提到以下任一情况：

- 要整理/打包 OA 报销材料
- 手上有发票扫描件 PDF + 出库单 + 需要填验收单
- 提到「非统一采购」「耗材验收单」「河南省人民医院」

## 铁律（每条都是踩坑换来的，不许违反）

### 1. PDF 必须是扫描件

**唯一合法来源是扫描件 PDF。** 从 PDF 里提取每页的原始位图，只做旋转和版面摆放。

**绝对禁止**把手机照片（桌面背景、纸张折痕、斜角拍摄）拼成 PDF。那样出来的不是扫描件，是照片集，财务不认。

```python
import pypdf
reader = pypdf.PdfReader(scan_pdf)
img = reader.pages[i].images[0].image   # 原始位图，零重编码
```

### 2. 逐页人工确认方向

**不要**用投影方差、自相关之类的算法自动判方向——实测会翻车（90° 和 270° 的方差完全一样，无法区分正反）。

做法：把每页按候选角度渲染成 PNG，**一张张看**，确认：

- 发票号读得对不对
- 标题「电子发票」/「出库单」朝向对不对
- 内容是不是正的（字不是倒的、不是躺着的）

### 3. 页面规格对齐 OA 原件

| 类型 | 尺寸 (pt) | 方向 |
|---|---|---|
| 发票 `.pdf` | 595.276 × 396.85 | 横版 |
| 出库单 `.1.pdf` | 595 × 842 | A4 竖版 |

发票要裁紧（去多余白边），出库单保留整张 A4（含下方空白），跟原件一致。

### 4. docx 只改文本，样式一个字不动

验收单是从母版克隆的。改内容是替换 `word/document.xml` 里的文本节点，**不要**重建文档。

必须保留：表格边框、字体（宋体/仿宋）、字号、页边距、横向节属性。

```python
with zipfile.ZipFile(src) as z:
    data = {n: z.read(n) for n in z.namelist()}
xml = data["word/document.xml"].decode("utf-8")
xml = xml.replace(OLD_NAME, NEW_NAME)
data["word/document.xml"] = xml.encode("utf-8")
```

### 5. 金额用含税口径

发票上印的是**不含税**金额，验收单要填**含税**单价/金额。

校验公式：
```
各行(单价 × 数量) = 各行金额
Σ各行金额 = 票面「价税合计」
```

对不上就别交，回去重算。

## 开工前必问（三项，缺一不可）

收到票据后，**先问清这三项再动手**，不许猜、不许用默认值：

| # | 要问的 | 为什么 |
|---|---|---|
| 1 | **项目名称** | 不恒定，每批不同 |
| 2 | **项目编码** | 不恒定，可能有多个项目混在一批 |
| 3 | **经办人** | 不恒定，不是固定某个人 |

**验收时间不问** —— 默认取当天日期（见下节）。

补充确认：

- 同一批里是否存在**多个项目**？（有的话要逐份指定名称和编码）
- 科室信息是否需要预填？（默认全部留空手写）

**如果用户说「按上次的来」，也要逐项念一遍确认**，因为上次的可能和这批不同。

## 命名与目录结构

```
OA/
├── 打印耗材验收单.bat           双击批量打印（路径无关）
├── 打印耗材验收单.ps1           实际执行脚本
├── <发票尾号>/
│   ├── <发票尾号>.pdf                 发票（横版扫描件）
│   ├── <发票尾号>.1.pdf               出库单（A4 竖版扫描件）
│   ├── <发票尾号>.2.jpg               验收单手写签字照片
│   └── <发票尾号>（非统一采购）.docx   验收单
```

- **发票尾号** = 发票号码后 4 位
- 一个尾号一个文件夹
- 每组固定 4 个文件（发票 / 出库单 / 验收单 / 签字照）

## 验收单字段

| 字段 | 来源 |
|---|---|
| 项目名称 | **不恒定** —— 用户每次指定 |
| 项目编码 | **不恒定** —— 用户每次指定，可能多套并存 |
| 验收时间 | **默认当天日期** —— 用户另指定则用指定的 |
| 经办人 | **不恒定** —— 用户指定 |
| 供应商 | 票面销售方全称（从发票读） |
| 发票号码 | 票面 20 位（从发票读） |
| 明细 | 逐行：名称 / 规格 / 数量 / 单价 / 金额（从发票读） |
| 签名·工号·科室 | **留空**，手填 |


### 验收时间：默认当天

**不必每次询问，默认填执行当天的日期。**

```
格式：YYYY年M月D日     例：2026年11月8日
注意：月、日不补零（写 11月8日，不写 11月08日）
```

取当天日期的正确方式 —— **以系统时间为准**：

```python
from datetime import date
d = date.today()
DATE = "%d年%d月%d日" % (d.year, d.month, d.day)
```

例外（此时才需要问）：

- 用户明确说「按某月某日算」→ 用用户给的
- 票据是补录历史单，验收日与打印日不同 → 问清楚

**不要**用发票上的开票日期充数，那是开票时间，不是验收时间。

### 不恒定字段的规则（重要）

**项目名称、项目编码、验收时间、经办人 —— 这四项每批都可能不同，不许写死默认值。**

```
绝对禁止：
  ❌ 经办人默认填「牛科润」
  ❌ 项目编码默认填某个固定值
  ❌ 验收时间用票面日期（用当天，不是开票日）
```

正确做法：

- 用户明确给了 → 照填
- 用户没给 → **留空并主动问**，不要自己编
- 同一批里不同票可能属于不同项目 → 允许逐份不同

**科室同理**：签名栏三处「科室」全部留空手填，脚本不预填任何值。


### 6. 手写签字照片必须按 `<尾号>.2.jpg` 命名

打印签字后拍的照片，每张对应一张验收单。处理步骤：

1. **逐张看发票号**，确认属于哪个尾号（绝对不能靠文件名或顺序猜）
2. 透视校正：纸张在木桌上斜拍，用四点校正展平
3. **保持横版**：验收单是 A4 横向，长边必须是宽
4. 命名为 `{尾号}.2.jpg` 放进对应文件夹

```python
# 关键：四点排序必须保证长边水平
wtop = norm(tr - tl); hleft = norm(bl - tl)
if wtop < hleft:   # 竖了，旋转 90 度重新映射
    quad = [tr, br, bl, tl]
```

**坑**：一开始按 y 坐标分上下组排序，遇到纸张有旋转就会把长短边搞反，
出来是竖版歪图。必须显式检查 `wtop vs hleft`。

### 7. zip 必须强制 UTF-8 标志

Python `zipfile.write()` 对纯 ASCII 文件名**不会**设 UTF-8 位（0x800），
中文名才设。混在一起难保不出乱码。自己写写入器，所有条目强制 `flag_bits |= 0x800`。

```python
lh.writeUInt16LE(0x800, 6)   # local header UTF-8 flag
ch.writeUInt16LE(0x800, 8)   # central dir UTF-8 flag
```

### 8. 校验 docx 时记得它是 zip

```python
# 错：直接 decode docx 字节
xml = zipf.read('xxx.docx').decode('utf-8')   # UnicodeDecodeError!

# 对：先当 zip 打开，再读 document.xml
blob = zipf.read('xxx.docx')
inner = zipfile.ZipFile(BytesIO(blob))
xml = inner.read('word/document.xml').decode('utf-8')
```


### 9. OA 根目录必须放批量打印脚本

每组文件夹里的 docx 要能**一次全打印**，不用一个个点开。根目录放两个文件：

| 文件 | 作用 |
|---|---|
| `打印耗材验收单.bat` | 双击启动（纯 ASCII，避免编码问题） |
| `打印耗材验收单.ps1` | 实际逻辑（UTF-8 BOM 保存） |

#### 路径无关是硬要求

这个包会被拷到**各种电脑**上，盘符、目录名全都不一样。**绝对不许写死路径。**

```powershell
# 正确：以脚本自身位置为根
if ($PSScriptRoot) { $Root = $PSScriptRoot }
else { $Root = Split-Path -Parent $MyInvocation.MyCommand.Path }
```

bat 同理，用 `%~dp0` 拿自身目录，再让 PowerShell 去那里找 .ps1：

```bat
set "HERE=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s = Get-ChildItem -LiteralPath $env:HERE -Filter '*.ps1' | Select-Object -First 1; & $s.FullName"
```

#### 必须跳过 Word 锁文件

Word 打开文档时会在同目录生成 `~$xxx.docx`。不过滤的话会多打印空文件：

```powershell
Where-Object { $_.Name -notlike '~$*' }
```

#### 提供 -DryRun 与 -Force

- `-DryRun`：只列清单不打印，让你先核对
- `-Force`：跳过确认直接打印
- `-Engine Auto|Word|WPS|Shell`：强制指定打印引擎（默认 Auto 自动探测）
- 默认：列出清单 + 显示默认打印机 + 询问 Y/n

#### 必须带日志（报错排查的唯一线索）

**这是硬要求。** 用户的电脑环境千差万别——PowerShell 版本、Office 位数、执行策略、默认打印机、权限……都会导致脚本报错。没有日志就等于闭眼修 bug。

**日志必须在最开头建立**，不能等出错才想起写日志（那时可能已经崩了）。

```powershell
# 优先级：用户指定 > 脚本所在目录 > TEMP
$candidates = @()
if ($LogDir) { $candidates += $LogDir }
$candidates += $Root
$candidates += $env:TEMP

foreach ($dir in $candidates) {
    try {
        # 探测可写性（写个临时文件试试）
        $probe = Join-Path $dir ('.wt_' + $stamp + '.tmp')
        Set-Content -LiteralPath $probe -Value 'x' -Encoding ASCII -ErrorAction Stop
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        $script:LogPath = Join-Path $dir $logName
        break
    } catch { continue }
}
```

**为什么要三级兜底**：根目录可能只读（U 盘写保护、网络盘权限），TEMP 几乎总可写。

#### 日志必须记录的内容

| 类别 | 具体项 |
|---|---|
| **环境快照** | PowerShell 版本、CLR、进程位数、OS 版本、用户名、计算机名、执行策略 |
| **路径** | 脚本路径、根目录、当前目录 |
| **打印机** | 默认打印机名称（读不到也要记 WARN） |
| **扫描结果** | 找到几个 docx、每个的相对路径和大小 |
| **引擎探测** | 每个 ProgID 试了没、可用与否、失败原因 |
| **文件关联** | `.docx` 关联到哪个 ProgID、print 动词存不存在 |
| **逐份结果** | 每份的开始、成功/失败、失败原因 |
| **失败汇总** | 失败清单 + 总数统计 |

#### 异常必须记录完整上下文

光记 `$_.Exception.Message` 不够，那只告诉你「失败了」，不告诉你「为什么、在哪」。

```powershell
function Format-Error {
    param($Err)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('  异常类型 : ' + $Err.Exception.GetType().FullName)
    [void]$sb.AppendLine('  消息     : ' + $Err.Exception.Message)
    if ($Err.Exception.InnerException) {
        [void]$sb.AppendLine('  内部异常 : ' + ...)   # COM 错误常藏在这里
    }
    if ($Err.CategoryInfo) {
        [void]$sb.AppendLine('  错误类别 : ' + $Err.CategoryInfo.Category)
    }
    if ($Err.InvocationInfo.PositionMessage) {
        [void]$sb.AppendLine('  出错位置 : ' + ...)    # 行号 + 源码片段
    }
    if ($Err.ScriptStackTrace) {
        [void]$sb.AppendLine('  调用栈   :')          # 完整调用链
    }
    return $sb.ToString()
}
```

实测效果（损坏 docx 的真实日志）：

```
[ERROR]       -> 失败: 文件可能已经损坏。
[ERROR]   异常类型 : System.Runtime.InteropServices.COMException
  消息     : 文件可能已经损坏。
  错误类别 : OperationStopped /
  出错位置 : At ...\打印耗材验收单.ps1:368 char:17
+ ...  $doc = $app.Documents.Open($d.FullName, $false, $true)
  调用栈   :
    at <ScriptBlock>, ...\打印耗材验收单.ps1: line 368
```

#### 控制台编码必须先设

不设的话，`Write-Host` 输出中文在某些环境下直接抛异常：

```powershell
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch { }   # 老系统设不了也不能崩
```

#### 提供 -Diagnose 模式

不打印，只跑环境检查并出日志。**这是用户报错时的第一句话**：「先跑一下 -Diagnose 把日志发我」。

```powershell
if ($Diagnose) {
    Write-Log '===== 诊断完成 =====' 'STEP'
    Write-Log '本次为诊断模式，未打印任何文件。' 'WARN'
    Save-Log
    exit 0
}
```

#### 日志文件名带时间戳

`打印日志_yyyyMMdd_HHmmss.txt` —— 不覆盖历史，多次运行可以对比。用户直接把这个文件发过来就能定位。

#### bat 里要给出日志提示

bat 是纯 ASCII（避免编码炸弹），所以提示用英文：

```bat
if not "%RC%"=="0" (
    echo [FAILED] exit code: %RC%
    echo A log file was written to this folder.
    echo Please send that log file to support for diagnosis.
)
```

`chcp 65001` 可以试着切 UTF-8，但**不要依赖它**——老系统上不生效。

#### 打印顺序：必须倒序

打印机出纸是**先进后出**（纸张叠放，后打的在上面）。所以必须**按文件夹名倒序**打印，
这样最上面才是第一份（尾号最小的）。

```powershell
# 倒序：7021 -> 4911 -> 3821 -> 3271 -> 2931 -> 1266
# 出纸后最上面是 1266，符合翻阅顺序
$docs = Get-ChildItem -LiteralPath $Root -Recurse -File -Filter '*.docx' |
        Where-Object { $_.Name -notlike '~$*' } |
        Sort-Object @{ Expression = { $_.Directory.Name }; Descending = $true },
                    @{ Expression = { $_.Name }; Descending = $true }
```

**不要**用正序，否则拿到手是反的，还得自己重新排。

#### 打印机「设置有问题」的多策略降级

真实故障（住院骨肿瘤病区那台）：

```
默认打印机 : \\胃肠外科一D07\HP LaserJet 1020
[ERROR] -> 失败: 由于打印机的当前设置有问题，Windows 无法打印。
  出错位置 : $doc.PrintOut()
  每份耗时 57 秒才报错
```

**这类网络共享的 GDI 打印机（HP LaserJet 1020 是典型）很容易出这问题。**
单靠 `$doc.PrintOut()` 一种方法不够，必须多策略降级：

| 顺序 | 策略 | 说明 |
|---|---|---|
| 1 | **ActivePrinter + 同步打印** | 先设 `$App.ActivePrinter`，再 `PrintOut([ref]$false,[ref]$false)` |
| 2 | PrintOut 默认（后台） | 最常见写法 |
| 3 | **转 PDF 后打印** | 绕过 Word 打印通道，对驱动不兼容最有效 |

```powershell
# 策略 1 —— 对网络打印机最有效
$App.ActivePrinter = $PrinterName
$doc = $App.Documents.Open($DocPath, $false, $true)
$doc.PrintOut([ref]$false, [ref]$false)   # 第二个参数 Background=false，同步等待
```

**同步打印的好处**：能立刻知道成功还是失败，不会把任务堆在队列里拖 57 秒。
实测从 57 秒降到 8.7 秒。

#### 必须记录打印耗时

用 `Stopwatch` 包住每次打印，日志里记下耗时：

```powershell
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$res = Invoke-PrintOne -App $app -DocPath $d.FullName -PrinterName $printerName
$sw.Stop()
# 日志: -> 已送打印 [ActivePrinter+同步] 用时 8752ms
```

**为什么重要**：耗时 57 秒说明超时，耗时 400ms 说明立即拒绝。
从耗时能判断是「打印机无响应」还是「参数不对」。

#### 必须自动适配 Word / WPS

**每台电脑装的 Office 不一样**：有的装 Microsoft Word，有的装 WPS。脚本不能只认一种。

三级探测，逐级降级：

```powershell
# 1. Word.Application —— 新版 WPS 也注册了这个 ProgID，先试它最省事
# 2. KWPS.Application / WPS.Application —— WPS 原生 ProgID
# 3. Shell —— 交给系统 .docx 文件关联（Start-Process -Verb Print）
```

探测函数：

```powershell
function Test-ComProgId {
    param([string]$ProgId)
    try {
        $obj = New-Object -ComObject $ProgId -ErrorAction Stop
        try { $obj.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch { }
        return $true
    } catch { return $false }
}
```

COM 打印：

```powershell
$app = New-Object -ComObject $progId
$app.Visible = $false
$app.DisplayAlerts = 0
$doc = $app.Documents.Open($path, $false, $true)   # 只读打开，防误改
$doc.PrintOut()
$doc.Close(0)
```

兜底（Shell）：

```powershell
Start-Process -FilePath $docxPath -Verb Print
```

**坑**：
- `DisplayAlerts = 0` 必须设，否则打印机配置异常时弹框会卡死批量流程。
- 强制 `-Engine WPS` 但机器上没有 WPS 时，**必须给黄色警告**再降级，不能静默切走。
- Shell 模式会短暂弹出 Word 窗口，属正常现象（走文件关联的固有行为）。

## 交付前自检

全部通过才算完成：

- [ ] zip `testzip()` 无损坏
- [ ] 每个 PDF 页数 = 1
- [ ] 发票 595×397 横版
- [ ] 出库单 595×842 竖版
- [ ] 每份 docx 发票号 == 所在文件夹名
- [ ] 每组有 4 个文件（含 `{尾号}.2.jpg` 签字照）
- [ ] OA 根目录有 `打印耗材验收单.bat` 和 `.ps1`
- [ ] 打印脚本 `-DryRun` 能正确列出全部 docx（且不含 `~$` 锁文件）
- [ ] 打印脚本 `-Diagnose` 能输出环境日志
- [ ] 跑一次确认生成了 `打印日志_*.txt`，内容含环境/引擎/文件清单
- [ ] 打印顺序为**倒序**（7021 → ... → 1266）
- [ ] 日志里记录了每份的打印耗时和所用策略
- [ ] 项目名称逐份核对（该改的改了，该留的留着）
- [ ] 明细金额求和 == 票面价税合计
- [ ] Word 渲染一遍验收单，确认没错位

## 输出

`D:\DeepSeekWorkspace\FINAL\OA.zip`，每次覆盖重建。

## 配套脚本

| 文件 | 作用 |
|---|---|
| `scripts/build_all.py` | **一键构建（主入口）**，改 CONFIG 后直接跑，3.7 秒出包 |
| `scripts/oa_pipeline.py` | 流水线核心：PDF/签字照/docx/打包，全部函数内联 |
| `scripts/oa_tools.py` | 通用工具函数（供单独调用） |
| `scripts/打印耗材验收单.ps1` | 批量打印（Word/WPS 自适应 + 日志 + 倒序） |
| `scripts/打印耗材验收单.bat` | 双击启动器 |
| `templates/验收单母版.docx` | 验收单母版（已净化，无真实数据） |

### 常用函数

| 函数 | 作用 |
|---|---|
| `P.make_pdf(img, kind, dst)` | 摆成发票/出库单规格 |
| `P.warp_signature(img_bgr)` | 签字照透视校正 |
| `P.make_contact_sheet(images, dst)` | **拼图预览（提速关键）** |
| `P.build_docx(...)` | 重建验收单表格 |
| `P.make_zip(src, zip)` | 打包（UTF-8 名） |

## 常见问题

**Q: 扫描 PDF 一页里有多张票怎么办？**  
先按内容切分，或让用户重扫。不要硬裁。

**Q: 某页方向拿不准？**  
宁可多渲染几个角度对比。看标题文字方向最可靠。

**Q: 发票号和出库单上的号对不上？**  
以发票为准，同时提醒用户核对。

**Q: 用户只给照片没给扫描件？**  
告诉他照片不合格，要扫描全能王那类真扫描 PDF。
