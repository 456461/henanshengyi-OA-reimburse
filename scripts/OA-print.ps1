#Requires -Version 5.1
<#
  OA 报销材料 一键打印（统一入口）

  ★ 按真实工作流分两步，别一次全打：

    第一步  打印验收单 -> 手写签名 -> 拍照发回给 agent 打包
    第二步  打印发票 + 出库单/清单（.1.pdf）

  菜单：
    [1] 第一步：打印全部验收单（docx）
    [2] 第二步：打印全部发票 + 出库单
    [3] 单独重打某一组（按尾号）<- 打废了、漏签了用这个
    [4] 全部一次打完（一般不用）
    [D] 诊断模式
    [Q] 退出

  文件命名约定：
    <尾号>.pdf      发票
    <尾号>.1.pdf    出库单 / 清单

  打印顺序（先进后出，后打的落在最上面）：
    组内：先出库单，后发票  -> 拿到手最上面是发票，下面是出库单
    组间：尾号倒序（7021 -> ... -> 1266），最上面是 1266 那组

  参数（非交互，方便脚本调用）：
    -Mode Menu|Step1|Step2|Group|All|Diagnose
    -Group 1266          配合 -Mode Group
    -Items All|Pdf|Docx  配合 -Mode Group
    -DryRun / -Force / -Engine Auto|Word|WPS|Shell / -LogDir <path>
#>

param(
    [ValidateSet('Menu','Step1','Step2','Group','All','Diagnose','Docs','Materials','Pdf')]
    [string]$Mode = 'Menu',

    [string]$Group = '',

    [ValidateSet('All','Pdf','Docx')]
    [string]$Items = 'All',

    [switch]$DryRun,
    [switch]$Force,

    [ValidateSet('Auto','Word','WPS','Shell')]
    [string]$Engine = 'Auto',

    [string]$LogDir = ''
)

$ErrorActionPreference = 'Continue'

# ================= 日志基础设施 =================

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch { }

$script:LogPath = $null
$script:LogBuffer = New-Object System.Collections.ArrayList
$script:PdfHandler = $null

function Write-Log {
    param(
        [string]$Msg = '',
        [ValidateSet('INFO','OK','WARN','ERROR','STEP','DRY')]
        [string]$Level = 'INFO'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = '[' + $ts + '][' + $Level.PadRight(5) + '] ' + $Msg
    $color = switch ($Level) {
        'ERROR' { 'Red' }
        'WARN'  { 'Yellow' }
        'OK'    { 'Green' }
        'STEP'  { 'Cyan' }
        'DRY'   { 'Magenta' }
        default { 'Gray' }
    }
    try { Write-Host $line -ForegroundColor $color } catch { Write-Host $line }
    [void]$script:LogBuffer.Add($line)
}

function Save-Log {
    if (-not $script:LogPath) { return }
    try {
        $script:LogBuffer | Set-Content -LiteralPath $script:LogPath -Encoding UTF8 -ErrorAction Stop
    } catch {
        try {
            $alt = Join-Path $env:TEMP ([System.IO.Path]::GetFileName($script:LogPath))
            $script:LogBuffer | Set-Content -LiteralPath $alt -Encoding UTF8 -ErrorAction Stop
            $script:LogPath = $alt
        } catch { }
    }
}

function Format-Error {
    param($Err)
    if (-not $Err) { return '  (no error object)' }
    $sb = New-Object System.Text.StringBuilder
    try {
        [void]$sb.AppendLine('  异常类型 : ' + $Err.Exception.GetType().FullName)
        [void]$sb.AppendLine('  消息     : ' + $Err.Exception.Message)
        if ($Err.Exception.InnerException) {
            [void]$sb.AppendLine('  内部异常 : ' + $Err.Exception.InnerException.GetType().Name + ' -> ' + $Err.Exception.InnerException.Message)
        }
        if ($Err.InvocationInfo -and $Err.InvocationInfo.PositionMessage) {
            [void]$sb.AppendLine('  出错位置 : ' + $Err.InvocationInfo.PositionMessage)
        }
    } catch {
        [void]$sb.AppendLine('  (格式化异常失败: ' + $_.Exception.Message + ')')
    }
    return $sb.ToString()
}

# ================= 1. 定位根目录 =================

if ($PSScriptRoot) {
    $Root = $PSScriptRoot
} else {
    $Root = Split-Path -Parent $MyInvocation.MyCommand.Path
}
try { $Root = (Resolve-Path -LiteralPath $Root).Path } catch { }

# ================= 2. 建立日志 =================

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logName = '打印日志_' + $stamp + '.txt'

$candidates = @()
if ($LogDir) { $candidates += $LogDir }
$candidates += $Root
$candidates += $env:TEMP

foreach ($dir in $candidates) {
    if (-not $dir) { continue }
    try {
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        $probe = Join-Path $dir ('.wt_' + $stamp + '.tmp')
        Set-Content -LiteralPath $probe -Value 'x' -Encoding ASCII -ErrorAction Stop
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        $script:LogPath = Join-Path $dir $logName
        break
    } catch { continue }
}

# ================= 3. 默认打印机 =================

$printerName = '(未检测到)'
try {
    $pr = Get-CimInstance -ClassName Win32_Printer -Filter 'Default=TRUE' -ErrorAction Stop
    if ($pr) { $printerName = $pr.Name }
} catch {
    try {
        $pr2 = Get-WmiObject -Class Win32_Printer -Filter 'Default=TRUE' -ErrorAction Stop
        if ($pr2) { $printerName = $pr2.Name }
    } catch { }
}

# ================= 4. 扫描分组 =================

function Get-Groups {
    $dirs = @()
    try {
        $dirs = Get-ChildItem -LiteralPath $Root -Directory -ErrorAction Stop |
                Where-Object { $_.Name -notlike '_*' } |
                Sort-Object Name -Descending
    } catch {
        return @()
    }

    $list = @()
    foreach ($d in $dirs) {
        $code = $d.Name
        $dirFull = $d.FullName

        $slipPath    = Join-Path $dirFull ($code + '.1.pdf')
        $invoicePath = Join-Path $dirFull ($code + '.pdf')

        $docxPath = $null
        try {
            $cand = Get-ChildItem -LiteralPath $dirFull -File -Filter '*.docx' -ErrorAction Stop |
                    Where-Object { $_.Name -notlike '~$*' } |
                    Select-Object -First 1
            if ($cand) { $docxPath = $cand.FullName }
        } catch { }

        $hasSlip    = Test-Path -LiteralPath $slipPath
        $hasInvoice = Test-Path -LiteralPath $invoicePath
        $hasDocx    = [bool]$docxPath

        if (-not ($hasSlip -or $hasInvoice -or $hasDocx)) { continue }

        $list += [pscustomobject]@{
            Code       = $code
            Dir        = $dirFull
            Slip       = $(if ($hasSlip)    { $slipPath }    else { $null })
            Invoice    = $(if ($hasInvoice) { $invoicePath } else { $null })
            Docx       = $docxPath
            HasSlip    = $hasSlip
            HasInvoice = $hasInvoice
            HasDocx    = $hasDocx
        }
    }
    return $list
}

$Groups = @(Get-Groups)
$allCodes = @($Groups | ForEach-Object { $_.Code })

# ================= 5. 打印计划 =================

function Build-Plan {
    param(
        [string[]]$Codes,
        [string]$What   # docs | pdf | all
    )
    $plan = New-Object System.Collections.ArrayList
    foreach ($code in $Codes) {
        $g = $Groups | Where-Object { $_.Code -eq $code } | Select-Object -First 1
        if (-not $g) { continue }

        if ($What -eq 'pdf' -or $What -eq 'all') {
            # 先进后出：先送出库单，后送发票 => 拿到手最上面是发票
            if ($g.Slip) {
                [void]$plan.Add([pscustomobject]@{ Path = $g.Slip; Kind = 'pdf'; Label = ($code + '.1.pdf   出库单/清单') })
            }
            if ($g.Invoice) {
                [void]$plan.Add([pscustomobject]@{ Path = $g.Invoice; Kind = 'pdf'; Label = ($code + '.pdf    发票') })
            }
        }
        if ($What -eq 'docs' -or $What -eq 'all') {
            if ($g.Docx) {
                [void]$plan.Add([pscustomobject]@{ Path = $g.Docx; Kind = 'docx'; Label = (Split-Path $g.Docx -Leaf) })
            }
        }
    }
    return @($plan)
}

# ================= 6. 引擎探测 =================

function Test-ComProgId {
    param([string]$ProgId)
    try {
        $obj = New-Object -ComObject $ProgId -ErrorAction Stop
        try { $obj.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch { }
        return $true
    } catch { return $false }
}

function Resolve-DocxEngine {
    $order = @()
    switch ($Engine) {
        'Word'  { $order = @('Word') }
        'WPS'   { $order = @('WPS') }
        'Shell' { $order = @('Shell') }
        default { $order = @('Word','WPS','Shell') }
    }
    foreach ($e in $order) {
        if ($e -eq 'Word') {
            if (Test-ComProgId 'Word.Application') { return [pscustomobject]@{ Kind='Word'; ProgId='Word.Application' } }
        } elseif ($e -eq 'WPS') {
            foreach ($k in @('KWPS.Application','WPS.Application')) {
                if (Test-ComProgId $k) { return [pscustomobject]@{ Kind='WPS'; ProgId=$k } }
            }
        } else {
            return [pscustomobject]@{ Kind='Shell'; ProgId=$null }
        }
    }
    return [pscustomobject]@{ Kind='Shell'; ProgId=$null }
}

function Resolve-PdfHandler {
    if ($null -ne $script:PdfHandler) { return $script:PdfHandler }

    $cands = @(
        @{ Kind='Sumatra'; Paths=@(
            (Join-Path $env:LOCALAPPDATA 'SumatraPDF\SumatraPDF.exe'),
            'C:\Program Files\SumatraPDF\SumatraPDF.exe',
            'C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe') },
        @{ Kind='Acrobat'; Paths=@(
            'C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe',
            'C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe',
            'C:\Program Files\Adobe\Acrobat\Acrobat.exe') },
        @{ Kind='Reader'; Paths=@(
            'C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe',
            'C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe',
            'C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe') },
        @{ Kind='Foxit'; Paths=@(
            'C:\Program Files (x86)\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe',
            'C:\Program Files\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe',
            'C:\Program Files (x86)\Foxit Software\Foxit Reader\FoxitReader.exe') }
    )

    foreach ($c in $cands) {
        foreach ($p in $c.Paths) {
            if ($p -and (Test-Path -LiteralPath $p)) {
                $script:PdfHandler = [pscustomobject]@{ Kind=$c.Kind; Exe=$p }
                return $script:PdfHandler
            }
        }
    }

    # 没装专用阅读器：退到 Edge / Chrome 的 kiosk 静默打印
    $browserCands = @(
        @{ Kind='Edge'; Paths=@(
            'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
            'C:\Program Files\Microsoft\Edge\Application\msedge.exe') },
        @{ Kind='Chrome'; Paths=@(
            'C:\Program Files\Google\Chrome\Application\chrome.exe',
            'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
            (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')) }
    )
    foreach ($c in $browserCands) {
        foreach ($p in $c.Paths) {
            if ($p -and (Test-Path -LiteralPath $p)) {
                $script:PdfHandler = [pscustomobject]@{ Kind=$c.Kind; Exe=$p }
                return $script:PdfHandler
            }
        }
    }

    $script:PdfHandler = [pscustomobject]@{ Kind='Shell'; Exe=$null }
    return $script:PdfHandler
}

# ================= 7. 单文件打印 =================

# 零依赖 PDF 打印模块（WinRT 渲染 + System.Drawing）
$pdfModule = Join-Path $PSScriptRoot 'pdf-print.ps1'
if (Test-Path -LiteralPath $pdfModule) {
    try { . $pdfModule } catch { Write-Log ('加载 PDF 打印模块失败: ' + $_.Exception.Message) 'WARN' }
}

function Invoke-PrintPdfOne {
    param([string]$Path, [string]$Printer, $Handler)

    $lastErr = $null
    $hasPrinter = ($Printer -and $Printer -notlike '(*')

    # ---- 策略 1：专用 PDF 阅读器命令行（若装了）----
    if ($Handler -and $Handler.Kind -notin @('Shell','Edge','Chrome') -and $Handler.Exe) {
        try {
            $argLine = ''
            switch ($Handler.Kind) {
                'Acrobat' {
                    $argLine = '/t "' + $Path + '"'
                    if ($hasPrinter) { $argLine += ' "' + $Printer + '"' }
                }
                'Reader' {
                    $argLine = '/t "' + $Path + '"'
                    if ($hasPrinter) { $argLine += ' "' + $Printer + '"' }
                }
                'Foxit' {
                    $argLine = '/t "' + $Path + '"'
                    if ($hasPrinter) { $argLine += ' "' + $Printer + '"' }
                }
                'Sumatra' {
                    if ($hasPrinter) { $argLine = '-silent -print-to "' + $Printer + '" "' + $Path + '"' }
                    else { $argLine = '-silent -print-to-default "' + $Path + '"' }
                }
            }
            Start-Process -FilePath $Handler.Exe -ArgumentList $argLine -ErrorAction Stop
            Start-Sleep -Milliseconds 1800
            return [pscustomobject]@{ Ok=$true; Strategy=($Handler.Kind + ' 命令行'); Error=$null }
        } catch {
            $lastErr = $_
        }
    }

    # ---- 策略 2：WinRT 渲染 + System.Drawing（零依赖，最稳）----
    if (Get-Command Invoke-PrintPdfNative -ErrorAction SilentlyContinue) {
        $r = Invoke-PrintPdfNative -Path $Path -Printer $Printer
        if ($r.Ok) { return $r }
        if ($r.Error) { $lastErr = $r.Error }
    }

    # ---- 策略 3：Shell Print 动词 ----
    try {
        Start-Process -FilePath $Path -Verb Print -ErrorAction Stop
        Start-Sleep -Milliseconds 1800
        return [pscustomobject]@{ Ok=$true; Strategy='Shell Print 动词'; Error=$null }
    } catch {
        $lastErr = $_
    }

    return [pscustomobject]@{ Ok=$false; Strategy='全部失败'; Error=$lastErr }
}

function Invoke-PrintDocxOne {
    param($App, [string]$DocPath, [string]$Printer)

    $doc = $null
    $lastErr = $null
    $hasPrinter = ($Printer -and $Printer -notlike '(*')

    if ($hasPrinter) {
        try {
            try { $App.ActivePrinter = $Printer } catch { }
            $doc = $App.Documents.Open($DocPath, $false, $true)
            $doc.PrintOut([ref]$false, [ref]$false)
            Start-Sleep -Milliseconds 600
            try { $doc.Close(0) } catch { }
            return [pscustomobject]@{ Ok=$true; Strategy='ActivePrinter+同步'; Error=$null }
        } catch {
            $lastErr = $_
            if ($doc) { try { $doc.Close(0) } catch { }; $doc = $null }
        }
    }

    try {
        $doc = $App.Documents.Open($DocPath, $false, $true)
        $doc.PrintOut()
        Start-Sleep -Milliseconds 600
        try { $doc.Close(0) } catch { }
        return [pscustomobject]@{ Ok=$true; Strategy='PrintOut 默认'; Error=$null }
    } catch {
        $lastErr = $_
        if ($doc) { try { $doc.Close(0) } catch { }; $doc = $null }
    }

    try {
        $tmpPdf = Join-Path $env:TEMP ('_oa_' + $stamp + '_' + [System.IO.Path]::GetFileNameWithoutExtension($DocPath) + '.pdf')
        $doc = $App.Documents.Open($DocPath, $false, $true)
        $doc.ExportAsFixedFormat($tmpPdf, 17)
        try { $doc.Close(0) } catch { }
        $doc = $null
        Start-Process -FilePath $tmpPdf -Verb Print -ErrorAction Stop
        Start-Sleep -Milliseconds 1500
        return [pscustomobject]@{ Ok=$true; Strategy='转PDF后打印'; Error=$null }
    } catch {
        $lastErr = $_
        if ($doc) { try { $doc.Close(0) } catch { } }
    }

    return [pscustomobject]@{ Ok=$false; Strategy='全部失败'; Error=$lastErr }
}

function Invoke-PrintPlan {
    param([array]$Plan)

    $ok = 0
    $fail = 0
    $failList = @()

    if ($Plan.Count -eq 0) {
        Write-Log '没有可打印的文件。' 'WARN'
        return [pscustomobject]@{ Ok=0; Fail=0; FailList=@() }
    }

    $pdfHandler = Resolve-PdfHandler
    $pdfDesc = $pdfHandler.Kind
    if ($pdfHandler.Exe) { $pdfDesc += '  ' + $pdfHandler.Exe }
    Write-Log ('PDF 打印器 : ' + $pdfDesc)

    $docxEngine = Resolve-DocxEngine
    $docxDesc = $docxEngine.Kind
    if ($docxEngine.ProgId) { $docxDesc += '  ' + $docxEngine.ProgId }
    Write-Log ('DOCX 引擎  : ' + $docxDesc)
    Write-Log ''

    $app = $null
    try {
        $i = 0
        foreach ($item in $Plan) {
            $i++
            Write-Log ('[' + $i + '/' + $Plan.Count + '] ' + $item.Label)
            $sw = [System.Diagnostics.Stopwatch]::StartNew()

            $res = $null
            try {
                if ($item.Kind -eq 'pdf') {
                    $res = Invoke-PrintPdfOne -Path $item.Path -Printer $printerName -Handler $pdfHandler
                } else {
                    if ($docxEngine.Kind -eq 'Shell') {
                        Start-Process -FilePath $item.Path -Verb Print -ErrorAction Stop
                        Start-Sleep -Milliseconds 1200
                        $res = [pscustomobject]@{ Ok=$true; Strategy='Shell'; Error=$null }
                    } else {
                        if (-not $app) {
                            $app = New-Object -ComObject $docxEngine.ProgId -ErrorAction Stop
                            try { $app.Visible = $false } catch { }
                            try { $app.DisplayAlerts = 0 } catch { }
                        }
                        $res = Invoke-PrintDocxOne -App $app -DocPath $item.Path -Printer $printerName
                    }
                }
            } catch {
                $res = [pscustomobject]@{ Ok=$false; Strategy='异常'; Error=$_ }
            }

            $sw.Stop()
            if ($res -and $res.Ok) {
                Write-Log ('      -> 已送打印 [' + $res.Strategy + '] ' + $sw.ElapsedMilliseconds + 'ms') 'OK'
                $ok++
            } else {
                $msg = '未知错误'
                if ($res -and $res.Error) { $msg = $res.Error.Exception.Message }
                Write-Log ('      -> 失败: ' + $msg) 'ERROR'
                if ($res -and $res.Error) { Write-Log (Format-Error $res.Error) 'ERROR' }
                $failList += ($item.Label + ' :: ' + $msg)
                $fail++
            }
        }
    } finally {
        if ($app) {
            try { $app.Quit() } catch { }
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
        }
    }

    return [pscustomobject]@{ Ok=$ok; Fail=$fail; FailList=$failList }
}

# ================= 8. 屏幕输出 =================

Write-Host ''
Write-Host ('=' * 62) -ForegroundColor Cyan
Write-Host '   OA 报销材料   一键打印' -ForegroundColor Cyan
Write-Host ('=' * 62) -ForegroundColor Cyan
Write-Host ''

Write-Log '===== 环境信息 =====' 'STEP'
Write-Log ('根目录     : ' + $Root)
try { Write-Log ('操作系统   : ' + [System.Environment]::OSVersion.VersionString) } catch { }
try { Write-Log ('PowerShell : ' + $PSVersionTable.PSVersion.ToString()) } catch { }
Write-Log ('默认打印机 : ' + $printerName)
Write-Log ''

Write-Log '===== 扫描材料 =====' 'STEP'

if ($Groups.Count -eq 0) {
    Write-Log ('在 ' + $Root + ' 下没有找到任何材料文件夹') 'ERROR'
    Write-Log '请确认本脚本与各尾号文件夹（1266 / 2931 / ...）在同一层。' 'WARN'
    Save-Log
    Write-Host ''
    Write-Host ('日志已保存: ' + $script:LogPath) -ForegroundColor Yellow
    # 只有交互模式才等按键，否则命令行调用会永久挂住
    if ($Mode -eq 'Menu') {
        Write-Host '按任意键退出...' -ForegroundColor DarkGray
        try { [void]$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') } catch { }
    }
    exit 1
}

Write-Log ('共 ' + $Groups.Count + ' 组（尾号倒序）：') 'OK'
foreach ($g in $Groups) {
    $parts = @()
    if ($g.HasInvoice) { $parts += '发票' }
    if ($g.HasSlip)    { $parts += '出库单' }
    if ($g.HasDocx)    { $parts += '验收单' }
    Write-Log ('  ' + $g.Code.PadRight(6) + ' ' + ($parts -join ' + '))
}
Write-Log ''

function Show-Diagnose {
    Write-Log '===== 诊断信息 =====' 'STEP'
    try {
        $pd = Get-CimInstance -ClassName Win32_Printer -Filter 'Default=TRUE' -ErrorAction SilentlyContinue
        if ($pd) {
            Write-Log ('  打印机驱动 : ' + $pd.DriverName)
            Write-Log ('  打印机端口 : ' + $pd.PortName)
            Write-Log ('  共享/本地  : ' + $pd.Shared + ' / ' + $pd.Local)
            Write-Log ('  脱机状态   : ' + $pd.WorkOffline)
            if (-not $pd.Local) { Write-Log '  提示：网络共享打印机，驱动不匹配易报「设置有问题」' 'WARN' }
        }
    } catch { }
    try {
        $sp = Get-Service -Name Spooler -ErrorAction Stop
        Write-Log ('  打印服务   : ' + $sp.Status)
    } catch { }

    $ph = Resolve-PdfHandler
    Write-Log ('  PDF 打印器 : ' + $ph.Kind + $(if ($ph.Exe) { '  ' + $ph.Exe } else { '  (未找到专用阅读器)' }))
    $de = Resolve-DocxEngine
    Write-Log ('  DOCX 引擎  : ' + $de.Kind + $(if ($de.ProgId) { '  ' + $de.ProgId } else { '' }))

    try {
        $assoc = (Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\.pdf' -ErrorAction SilentlyContinue).'(default)'
        Write-Log ('  .pdf 关联  : ' + $assoc)
        if ($assoc) {
            $pc = (Get-ItemProperty ('Registry::HKEY_CLASSES_ROOT\' + $assoc + '\shell\print\command') -ErrorAction SilentlyContinue).'(default)'
            if ($pc) { Write-Log '  .pdf print : 存在' } else { Write-Log '  .pdf print : 不存在' 'WARN' }
        }
    } catch { }

    Write-Log '诊断完成，未打印任何文件。' 'WARN'
}

function Show-PlanAndMaybePrint {
    param([array]$Plan, [string]$Title)

    if ($Plan.Count -eq 0) {
        Write-Log '没有匹配的文件。' 'WARN'
        return
    }

    Write-Log ('===== ' + $Title + ' =====') 'STEP'
    $i = 0
    foreach ($p in $Plan) {
        $i++
        Write-Log ('  ' + $i + '. ' + $p.Label)
    }
    Write-Log ''

    if ($DryRun) {
        Write-Log '试运行：只列清单，未实际打印。' 'DRY'
        return
    }

    if (-not $Force) {
        $ans = 'y'
        try { $ans = Read-Host ('确认打印以上 ' + $Plan.Count + ' 个文件? [Y/n]') } catch { $ans = 'y' }
        if ($ans -ne '' -and $ans.Trim().ToLower() -notmatch '^(y|yes)$') {
            Write-Log '用户取消。' 'WARN'
            return
        }
    }

    Write-Log '===== 开始打印 =====' 'STEP'
    $r = Invoke-PrintPlan -Plan $Plan
    Write-Log ''
    $lvl = 'OK'
    if ($r.Fail -gt 0) { $lvl = 'ERROR' }
    Write-Log ('完成：成功 ' + $r.Ok + ' 个，失败 ' + $r.Fail + ' 个') $lvl

    if ($r.FailList.Count -gt 0) {
        Write-Log '失败清单：' 'ERROR'
        foreach ($f in $r.FailList) { Write-Log ('  - ' + $f) 'ERROR' }
        Write-Log ''
        Write-Log '排查建议：' 'STEP'
        Write-Log '  1. 手动打开该文件按 Ctrl+P 试打，看是否同样报错' 'WARN'
        Write-Log '  2. 检查打印机是否脱机 / 缺纸 / 暂停' 'WARN'
        Write-Log '  3. 网络共享打印机：确认共享主机开机且打印服务运行' 'WARN'
        Write-Log '  4. 清空打印队列后重试（卡住的任务会阻塞后续）' 'WARN'
        Write-Log '  5. PDF 打不出：装 SumatraPDF（体积小、命令行稳）' 'WARN'
    }
}

# ================= 9. 非交互模式 =================

if ($Mode -ne 'Menu') {
    switch ($Mode) {
        'Diagnose'  { Show-Diagnose }
        'Step1'     { Show-PlanAndMaybePrint (Build-Plan -Codes $allCodes -What 'docs') '[第一步] 全部验收单' }
        'Step2'     { Show-PlanAndMaybePrint (Build-Plan -Codes $allCodes -What 'pdf')  '[第二步] 全部发票 + 出库单' }
        'Docs'      { Show-PlanAndMaybePrint (Build-Plan -Codes $allCodes -What 'docs') '全部验收单' }
        'Pdf'       { Show-PlanAndMaybePrint (Build-Plan -Codes $allCodes -What 'pdf')  '全部发票 + 出库单' }
        'Materials' { Show-PlanAndMaybePrint (Build-Plan -Codes $allCodes -What 'all')  '全部材料' }
        'All'       { Show-PlanAndMaybePrint (Build-Plan -Codes $allCodes -What 'all')  '全部材料' }
        'Group'     {
            if (-not $Group) {
                Write-Log '需要 -Group 参数，例如 -Group 1266' 'ERROR'
            } elseif ($allCodes -notcontains $Group) {
                Write-Log ('没有这个尾号: ' + $Group) 'ERROR'
                Write-Log ('可选: ' + ($allCodes -join ', ')) 'WARN'
            } else {
                $what = switch ($Items) { 'Pdf' { 'pdf' } 'Docx' { 'docs' } default { 'all' } }
                Show-PlanAndMaybePrint (Build-Plan -Codes @($Group) -What $what) ('尾号 ' + $Group)
            }
        }
    }
    Save-Log
    Write-Host ''
    Write-Host ('日志已保存: ' + $script:LogPath) -ForegroundColor Yellow
    exit 0
}

# ================= 10. 交互菜单 =================

function Show-Menu {
    Write-Host ''
    Write-Host ('-' * 64) -ForegroundColor DarkCyan
    Write-Host '  报销流程走两步，别一次全打：' -ForegroundColor White
    Write-Host ''
    Write-Host '   第一步  打印验收单 -> 手写签名 -> 拍照回传打包' -ForegroundColor DarkYellow
    Write-Host '   第二步  打包完成后再打印发票 + 出库单/清单' -ForegroundColor DarkYellow
    Write-Host ''
    Write-Host ('-' * 64) -ForegroundColor DarkCyan
    Write-Host ''
    Write-Host '   [1] 打印全部验收单（docx）' -ForegroundColor Green
    Write-Host '   [2] 打印某一组：发票 + 出库单   <- 输入尾号' -ForegroundColor Green
    Write-Host ''
    Write-Host '   [3] 打印全部发票 + 出库单' -ForegroundColor Gray
    Write-Host '   [4] 打印某一组：验收单 + 发票 + 出库单' -ForegroundColor Gray
    Write-Host '   [5] 打印全部（三样一起）' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '   [D] 诊断模式（不打印）' -ForegroundColor DarkGray
    Write-Host '   [Q] 退出' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host ('-' * 64) -ForegroundColor DarkCyan
}

function Select-GroupInteractive {
    $ordered = @($Groups)
    Write-Host ''
    Write-Host '  可选尾号：' -ForegroundColor White
    $i = 0
    foreach ($g in $ordered) {
        $i++
        $parts = @()
        if ($g.HasInvoice) { $parts += '发票' }
        if ($g.HasSlip)    { $parts += '出库单' }
        if ($g.HasDocx)    { $parts += '验收单' }
        Write-Host ('   [' + $i + '] ' + $g.Code.PadRight(6) + ' ' + ($parts -join ' + ')) -ForegroundColor Gray
    }
    Write-Host ''
    $sel = ''
    try { $sel = Read-Host '输入编号或尾号（直接回车取消）' } catch { return $null }
    if ([string]::IsNullOrWhiteSpace($sel)) { return $null }
    $sel = $sel.Trim()

    if ($sel -match '^\d+$' -and $sel.Length -le 2) {
        $idx = [int]$sel
        if ($idx -ge 1 -and $idx -le $ordered.Count) { return $ordered[$idx-1].Code }
    }
    if ($allCodes -contains $sel) { return $sel }

    Write-Host ('  没有匹配的尾号: ' + $sel) -ForegroundColor Red
    return $null
}

:menuLoop while ($true) {
    Show-Menu
    $choice = ''
    try { $choice = Read-Host '请输入选项' } catch { break menuLoop }
    if ($null -eq $choice) { break menuLoop }
    $choice = $choice.Trim().ToUpper()

    if ($choice -eq 'Q') { break menuLoop }

    switch ($choice) {
        '1' { Show-PlanAndMaybePrint (Build-Plan -Codes $allCodes -What 'docs') '全部验收单' }

        '2' {
            # 高频操作：选一组，直接打发票 + 出库单，不再多问一层
            $code = Select-GroupInteractive
            if ($code) {
                Show-PlanAndMaybePrint (Build-Plan -Codes @($code) -What 'pdf') ('尾号 ' + $code + '：发票 + 出库单')
            }
        }

        '3' { Show-PlanAndMaybePrint (Build-Plan -Codes $allCodes -What 'pdf')  '全部发票 + 出库单' }

        '4' {
            $code = Select-GroupInteractive
            if ($code) {
                Show-PlanAndMaybePrint (Build-Plan -Codes @($code) -What 'all') ('尾号 ' + $code + '：三样全打')
            }
        }

        '5' { Show-PlanAndMaybePrint (Build-Plan -Codes $allCodes -What 'all')  '全部材料（三样一起）' }

        'D' { Show-Diagnose }
        ''  { }
        default { Write-Host '  无效选项' -ForegroundColor Red }
    }
}

Save-Log
Write-Host ''
Write-Host ('日志已保存: ' + $script:LogPath) -ForegroundColor Yellow
Write-Host '按任意键退出...' -ForegroundColor DarkGray
try { [void]$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') } catch { }
exit 0
