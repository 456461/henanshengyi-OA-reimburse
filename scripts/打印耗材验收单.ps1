#Requires -Version 5.1
<#
  耗材验收单 批量打印（自动适配 Microsoft Word / WPS / 系统默认程序）

  用法：放在 OA 文件夹根目录，双击同目录的 .bat 即可。
  路径无关：自动以【本脚本所在目录】为根，换电脑换盘符都不用改。

  参数：
    -DryRun          只列出将打印的文件，不实际打印
    -Force           跳过确认，直接打印
    -Engine Auto|Word|WPS|Shell   强制指定打印引擎（默认 Auto 自动探测）
    -Diagnose        只做环境诊断并输出报告，不打印
    -LogDir <path>   指定日志目录（默认放在本脚本所在目录）

  日志：每次运行都会生成 打印日志_yyyyMMdd_HHmmss.txt
        报错时把这个文件发出来就能定位问题。
#>

param(
    [switch]$DryRun,
    [switch]$Force,
    [ValidateSet('Auto','Word','WPS','Shell')]
    [string]$Engine = 'Auto',
    [switch]$Diagnose,
    [string]$LogDir = ''
)

$ErrorActionPreference = 'Continue'

# ================= 日志基础设施（最先建立） =================

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch { }

$script:LogPath = $null
$script:LogBuffer = New-Object System.Collections.ArrayList

function Write-Log {
    param(
        [string]$Msg = '',
        [ValidateSet('INFO','OK','WARN','ERROR','STEP')]
        [string]$Level = 'INFO'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = '[' + $ts + '][' + $Level.PadRight(5) + '] ' + $Msg
    $color = switch ($Level) {
        'ERROR' { 'Red' }
        'WARN'  { 'Yellow' }
        'OK'    { 'Green' }
        'STEP'  { 'Cyan' }
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
        if ($Err.CategoryInfo) {
            [void]$sb.AppendLine('  错误类别 : ' + $Err.CategoryInfo.Category + ' / ' + $Err.CategoryInfo.Activity)
        }
        if ($Err.InvocationInfo -and $Err.InvocationInfo.PositionMessage) {
            [void]$sb.AppendLine('  出错位置 : ' + $Err.InvocationInfo.PositionMessage)
        }
        if ($Err.ScriptStackTrace) {
            [void]$sb.AppendLine('  调用栈   :')
            $stackLines = $Err.ScriptStackTrace -split [Environment]::NewLine
            foreach ($sline in $stackLines) { [void]$sb.AppendLine('    ' + $sline.Trim()) }
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

# ================= 2. 建立日志文件 =================

$script:stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$stamp = $script:stamp
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

Write-Host ''
Write-Host ('=' * 56) -ForegroundColor Cyan
Write-Host '   耗材验收单  批量打印' -ForegroundColor Cyan
Write-Host ('=' * 56) -ForegroundColor Cyan
Write-Host ''

# ================= 3. 环境快照 =================

Write-Log '===== 环境信息 =====' 'STEP'
Write-Log ('脚本路径   : ' + $PSCommandPath)
Write-Log ('根目录     : ' + $Root)
try { Write-Log ('当前目录   : ' + (Get-Location).Path) } catch { }
try { Write-Log ('操作系统   : ' + [System.Environment]::OSVersion.VersionString) } catch { }
try {
    $psv = $PSVersionTable.PSVersion.ToString()
    $ed = 'Desktop'
    try { if ($PSVersionTable.PSEdition) { $ed = $PSVersionTable.PSEdition } } catch { }
    Write-Log ('PowerShell : ' + $psv + ' (' + $ed + ')')
} catch { Write-Log 'PowerShell 版本读取失败' 'WARN' }
try { Write-Log ('CLR 版本   : ' + $PSVersionTable.CLRVersion) } catch { }
try {
    $bits = '32-bit'
    if ([System.Environment]::Is64BitProcess) { $bits = '64-bit' }
    Write-Log ('进程位数   : ' + $bits)
} catch { }
try { Write-Log ('用户名     : ' + [System.Environment]::UserName) } catch { }
try { Write-Log ('计算机名   : ' + [System.Environment]::MachineName) } catch { }
try { Write-Log ('执行策略   : ' + (Get-ExecutionPolicy)) } catch { Write-Log '执行策略读取失败' 'WARN' }
Write-Log ''

# 默认打印机
$printerName = '(未检测到)'
try {
    $pr = Get-CimInstance -ClassName Win32_Printer -Filter 'Default=TRUE' -ErrorAction Stop
    if ($pr) { $printerName = $pr.Name }
    else { Write-Log '系统未设置默认打印机' 'WARN' }
} catch {
    Write-Log ('读取默认打印机失败: ' + $_.Exception.Message) 'WARN'
    try {
        $pr2 = Get-WmiObject -Class Win32_Printer -Filter 'Default=TRUE' -ErrorAction Stop
        if ($pr2) { $printerName = $pr2.Name }
    } catch { }
}
Write-Log ('默认打印机 : ' + $printerName)
try {
    $pd = Get-CimInstance -ClassName Win32_Printer -Filter 'Default=TRUE' -ErrorAction SilentlyContinue
    if ($pd) {
        Write-Log ('  驱动     : ' + $pd.DriverName)
        Write-Log ('  端口     : ' + $pd.PortName)
        Write-Log ('  共享     : ' + $pd.Shared + ' / 本地: ' + $pd.Local)
        Write-Log ('  脱机     : ' + $pd.WorkOffline)
        if (-not $pd.Local) { Write-Log '  注意：这是网络共享打印机，驱动不匹配会导致「设置有问题」' 'WARN' }
    }
} catch { }
try {
    $sp = Get-Service -Name Spooler -ErrorAction Stop
    Write-Log ('打印服务   : ' + $sp.Status)
    if ($sp.Status -ne 'Running') { Write-Log '  打印后台服务未运行！' 'ERROR' }
} catch { }
Write-Log ''

# ================= 4. 扫描 docx =================

Write-Log '===== 扫描文档 =====' 'STEP'

$docs = @()
try {
    # 倒序：打印机出纸是先进后出，倒着打最后出来的才是第一份
    $docs = Get-ChildItem -LiteralPath $Root -Recurse -File -Filter '*.docx' -ErrorAction Stop |
            Where-Object { $_.Name -notlike '~$*' } |
            Sort-Object @{ Expression = { $_.Directory.Name }; Descending = $true },
                        @{ Expression = { $_.Name }; Descending = $true }
} catch {
    Write-Log ('扫描目录失败: ' + $_.Exception.Message) 'ERROR'
    Write-Log (Format-Error $_) 'ERROR'
    Save-Log
    Write-Host ''
    Write-Host ('日志已保存: ' + $script:LogPath) -ForegroundColor Yellow
    exit 3
}

if ($docs.Count -eq 0) {
    Write-Log '未找到任何 .docx 文件' 'WARN'
    Write-Log ('请确认本脚本与各尾号文件夹在同一层: ' + $Root) 'WARN'
    Save-Log
    Write-Host ''
    Write-Host ('日志已保存: ' + $script:LogPath) -ForegroundColor Yellow
    exit 1
}

Write-Log ('共找到 ' + $docs.Count + ' 份文档（倒序打印）') 'OK'
$n = 0
foreach ($d in $docs) {
    $n++
    $rel = $d.FullName.Substring($Root.Length).TrimStart([char]92)
    Write-Log ('  ' + $n + '. ' + $rel + '   (' + [math]::Round($d.Length/1024,1) + ' KB)')
}
Write-Log ''

# ================= 5. 探测打印引擎 =================

Write-Log '===== 打印引擎探测 =====' 'STEP'

function Test-ComProgId {
    param([string]$ProgId)
    try {
        $obj = New-Object -ComObject $ProgId -ErrorAction Stop
        try { $obj.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch { }
        Write-Log ('  ' + $ProgId + ' -> 可用') 'OK'
        return $true
    } catch {
        $msg = $_.Exception.Message
        if ($msg.Length -gt 120) { $msg = $msg.Substring(0,120) + '...' }
        Write-Log ('  ' + $ProgId + ' -> 不可用 (' + $msg + ')')
        return $false
    }
}

$engineOrder = @()
switch ($Engine) {
    'Word'  { $engineOrder = @('Word') }
    'WPS'   { $engineOrder = @('WPS') }
    'Shell' { $engineOrder = @('Shell') }
    default { $engineOrder = @('Word', 'WPS', 'Shell') }
}

$chosen = $null
$progId = $null
foreach ($e in $engineOrder) {
    if ($e -eq 'Word') {
        if (Test-ComProgId 'Word.Application') { $chosen = 'Word'; $progId = 'Word.Application'; break }
    }
    elseif ($e -eq 'WPS') {
        foreach ($k in @('KWPS.Application', 'WPS.Application')) {
            if (Test-ComProgId $k) { $chosen = 'WPS'; $progId = $k; break }
        }
        if ($chosen) { break }
    }
    else {
        $chosen = 'Shell'; break
    }
}

if (-not $chosen) { $chosen = 'Shell' }

if ($Engine -ne 'Auto' -and $Engine -ne 'Shell' -and $chosen -ne $Engine) {
    Write-Log ('未找到 ' + $Engine + '，已降级为系统默认程序') 'WARN'
}

$engineLabel = switch ($chosen) {
    'Word'  { 'Microsoft Word (' + $progId + ')' }
    'WPS'   { 'WPS (' + $progId + ')' }
    'Shell' { '系统默认程序（文件关联）' }
}
Write-Log ('最终引擎 : ' + $engineLabel) 'OK'

# .docx 文件关联检查（Shell 模式依赖）
try {
    $assoc = (Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\.docx' -ErrorAction SilentlyContinue).'(default)'
    if ($assoc) {
        Write-Log ('  .docx 关联 : ' + $assoc)
        $pc = (Get-ItemProperty ('Registry::HKEY_CLASSES_ROOT\' + $assoc + '\shell\print\command') -ErrorAction SilentlyContinue).'(default)'
        if ($pc) { Write-Log '  print 动词 : 存在' }
        else { Write-Log '  print 动词 : 不存在（Shell 模式会失败）' 'WARN' }
    } else {
        Write-Log '  .docx 未关联任何程序' 'WARN'
    }
} catch {
    Write-Log ('  读取文件关联失败: ' + $_.Exception.Message) 'WARN'
}
Write-Log ''

# ================= 6. 诊断模式 =================

if ($Diagnose) {
    Write-Log '===== 诊断完成 =====' 'STEP'
    Write-Log '本次为诊断模式，未打印任何文件。' 'WARN'
    Save-Log
    Write-Host ''
    Write-Host ('日志已保存: ' + $script:LogPath) -ForegroundColor Yellow
    exit 0
}

# ================= 7. 试运行 =================

if ($DryRun) {
    Write-Log '试运行：只列清单，未实际打印。' 'WARN'
    Save-Log
    Write-Host ''
    Write-Host ('日志已保存: ' + $script:LogPath) -ForegroundColor Yellow
    exit 0
}

# ================= 8. 确认 =================

if (-not $Force) {
    $ans = 'y'
    try {
        $ans = Read-Host ('确认打印以上 ' + $docs.Count + ' 份? [Y/n]')
    } catch {
        Write-Log '无法读取输入（非交互环境？），按确认处理' 'WARN'
        $ans = 'y'
    }
    if ($ans -ne '' -and $ans.Trim().ToLower() -notmatch '^(y|yes)$') {
        Write-Log '用户取消，未打印任何文件。' 'WARN'
        Save-Log
        Write-Host ''
        Write-Host ('日志已保存: ' + $script:LogPath) -ForegroundColor Yellow
        exit 0
    }
}

# ================= 9. 打印（多策略降级） =================

# 关键：网络共享的 GDI 打印机（如 HP LaserJet 1020）常见报错
#   「由于打印机的当前设置有问题，Windows 无法打印」
# 原因通常是 ActivePrinter 未正确设置、驱动不匹配、或后台队列卡死。
# 这里按三种策略依次尝试，任一成功即继续下一份。

function Invoke-PrintOne {
    param($App, $DocPath, $PrinterName)
    
    $doc = $null
    $lastErr = $null

    # ---- 策略 1：先显式设置 ActivePrinter，再同步打印 ----
    # 同步（Background = false）能避免任务堆在队列里报错，也能拿到真实结果
    if ($PrinterName) {
        try {
            try {
                $App.ActivePrinter = $PrinterName
            } catch {
                Write-Log ('        设置 ActivePrinter 失败: ' + $_.Exception.Message) 'WARN'
            }
            $doc = $App.Documents.Open($DocPath, $false, $true)
            $doc.PrintOut([ref]$false, [ref]$false)
            Start-Sleep -Milliseconds 600
            try { $doc.Close(0) } catch { }
            return @{ Ok = $true; Strategy = 'ActivePrinter+同步'; Error = $null }
        } catch {
            $lastErr = $_
            if ($doc) { try { $doc.Close(0) } catch { }; $doc = $null }
        }
    }

    # ---- 策略 2：普通 PrintOut（默认后台） ----
    try {
        $doc = $App.Documents.Open($DocPath, $false, $true)
        $doc.PrintOut()
        Start-Sleep -Milliseconds 600
        try { $doc.Close(0) } catch { }
        return @{ Ok = $true; Strategy = 'PrintOut 默认'; Error = $null }
    } catch {
        $lastErr = $_
        if ($doc) { try { $doc.Close(0) } catch { }; $doc = $null }
    }

    # ---- 策略 3：导出 PDF 后交给系统打印 ----
    # 绕过 Word 的打印通道，对驱动兼容性问题最有效
    try {
        $tmpPdf = Join-Path $env:TEMP ('_oa_' + $script:stamp + '_' + [System.IO.Path]::GetFileNameWithoutExtension($DocPath) + '.pdf')
        $doc = $App.Documents.Open($DocPath, $false, $true)
        $doc.ExportAsFixedFormat($tmpPdf, 17)
        try { $doc.Close(0) } catch { }
        $doc = $null
        Start-Process -FilePath $tmpPdf -Verb Print -ErrorAction Stop
        Start-Sleep -Milliseconds 1500
        return @{ Ok = $true; Strategy = '转PDF后打印'; Error = $null }
    } catch {
        $lastErr = $_
        if ($doc) { try { $doc.Close(0) } catch { } }
    }

    return @{ Ok = $false; Strategy = '全部失败'; Error = $lastErr }
}

# ================= 10. 打印 =================

Write-Log '===== 开始打印 =====' 'STEP'

$ok = 0
$fail = 0
$failList = @()

if ($chosen -eq 'Shell') {
    $n = 0
    foreach ($d in $docs) {
        $n++
        $rel = $d.FullName.Substring($Root.Length).TrimStart([char]92)
        Write-Log ('[' + $n + '/' + $docs.Count + '] ' + $rel)
        try {
            Start-Process -FilePath $d.FullName -Verb Print -ErrorAction Stop
            Start-Sleep -Milliseconds 1200
            Write-Log '      -> 已送打印 (Shell)' 'OK'
            $ok++
        } catch {
            Write-Log ('      -> 失败: ' + $_.Exception.Message) 'ERROR'
            Write-Log (Format-Error $_) 'ERROR'
            $failList += ($rel + ' :: ' + $_.Exception.Message)
            $fail++
        }
    }
}
else {
    $app = $null
    try {
        Write-Log ('启动 ' + $progId + ' ...')
        $app = New-Object -ComObject $progId -ErrorAction Stop
        try { $app.Visible = $false } catch { Write-Log '  设置 Visible 失败（不影响打印）' 'WARN' }
        try { $app.DisplayAlerts = 0 } catch { Write-Log '  设置 DisplayAlerts 失败（可能弹窗）' 'WARN' }

        $n = 0
        foreach ($d in $docs) {
            $n++
            $rel = $d.FullName.Substring($Root.Length).TrimStart([char]92)
            Write-Log ('[' + $n + '/' + $docs.Count + '] ' + $rel)

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $res = Invoke-PrintOne -App $app -DocPath $d.FullName -PrinterName $printerName
            $sw.Stop()

            if ($res.Ok) {
                Write-Log ('      -> 已送打印 [' + $res.Strategy + '] 用时 ' + $sw.ElapsedMilliseconds + 'ms') 'OK'
                $ok++
            } else {
                Write-Log ('      -> 失败（用时 ' + [math]::Round($sw.Elapsed.TotalSeconds,1) + ' 秒）') 'ERROR'
                if ($res.Error) {
                    Write-Log ('      ' + $res.Error.Exception.Message) 'ERROR'
                    Write-Log (Format-Error $res.Error) 'ERROR'
                }
                $reason = '未知错误'
                if ($res.Error) { $reason = $res.Error.Exception.Message }
                $failList += ($rel + ' :: ' + $reason)
                $fail++
            }
        }
    }
    catch {
        Write-Log ('无法启动 ' + $chosen + ': ' + $_.Exception.Message) 'ERROR'
        Write-Log (Format-Error $_) 'ERROR'
        $fail = $docs.Count - $ok
    }
    finally {
        if ($app) {
            try { $app.Quit() } catch { }
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
        }
    }
}

# ================= 10. 汇总 =================

Write-Log ''
if ($fail -gt 0) {
    Write-Log ('完成：成功 ' + $ok + ' 份，失败 ' + $fail + ' 份') 'ERROR'
} else {
    Write-Log ('完成：成功 ' + $ok + ' 份，失败 ' + $fail + ' 份') 'OK'
}

if ($failList.Count -gt 0) {
    Write-Log '失败清单：' 'ERROR'
    foreach ($fitem in $failList) { Write-Log ('  - ' + $fitem) 'ERROR' }
    Write-Log ''
    Write-Log '===== 排查建议 =====' 'STEP'
    Write-Log '若报「打印机的当前设置有问题，Windows 无法打印」：' 'WARN'
    Write-Log '  1. 手动打开任意一份 docx，按 Ctrl+P 试打，看是否同样报错' 'WARN'
    Write-Log '  2. 检查打印机是否脱机（设置 > 蓝牙和设备 > 打印机和扫描仪）' 'WARN'
    Write-Log '  3. 网络共享打印机：确认共享主机开机且打印服务运行' 'WARN'
    Write-Log '  4. HP LaserJet 1020 属 GDI 打印机，共享端与本机驱动须同版本' 'WARN'
    Write-Log '  5. 清空打印队列后重试（卡住的任务会阻塞后续打印）' 'WARN'
    Write-Log '  6. 终极方案：在共享主机上共享时勾选「其他驱动程序」安装 x64 驱动' 'WARN'
}

Save-Log
Write-Host ''
Write-Host ('日志已保存: ' + $script:LogPath) -ForegroundColor Yellow
Write-Host '（如遇报错，请把这个日志文件发给技术支持）' -ForegroundColor DarkGray
Write-Host ''

if ($fail -gt 0) { exit 2 } else { exit 0 }
