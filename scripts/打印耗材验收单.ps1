#Requires -Version 5.1
<#
  耗材验收单 批量打印（自动适配 Microsoft Word / WPS / 系统默认程序）

  用法：放在 OA 文件夹根目录，双击同目录的 .bat 即可。
  路径无关：自动以【本脚本所在目录】为根，换电脑换盘符都不用改。

  参数：
    -DryRun          只列出将打印的文件，不实际打印
    -Force           跳过确认，直接打印
    -Engine Auto|Word|WPS|Shell   强制指定打印引擎（默认 Auto 自动探测）
#>

param(
    [switch]$DryRun,
    [switch]$Force,
    [ValidateSet('Auto','Word','WPS','Shell')]
    [string]$Engine = 'Auto'
)

$ErrorActionPreference = 'Continue'

# ---------- 1. 定位根目录（脚本自身所在目录） ----------
if ($PSScriptRoot) {
    $Root = $PSScriptRoot
} else {
    $Root = Split-Path -Parent $MyInvocation.MyCommand.Path
}
try { $Root = (Resolve-Path -LiteralPath $Root).Path } catch { }

Write-Host ''
Write-Host ('=' * 52) -ForegroundColor Cyan
Write-Host '   耗材验收单  批量打印' -ForegroundColor Cyan
Write-Host ('=' * 52) -ForegroundColor Cyan
Write-Host ('  根目录: ' + $Root)
Write-Host ''

# ---------- 2. 收集 docx ----------
$docs = @()
try {
    $docs = Get-ChildItem -LiteralPath $Root -Recurse -File -Filter '*.docx' -ErrorAction Stop |
            Where-Object { $_.Name -notlike '~$*' } |
            Sort-Object @{ Expression = { $_.Directory.Name } }, @{ Expression = { $_.Name } }
} catch {
    Write-Host ('扫描目录失败: ' + $_.Exception.Message) -ForegroundColor Red
    exit 3
}

if ($docs.Count -eq 0) {
    Write-Host '未找到任何 .docx 文件。' -ForegroundColor Yellow
    Write-Host '请确认本脚本与各尾号文件夹放在同一层。' -ForegroundColor Yellow
    exit 1
}

Write-Host ('共找到 ' + $docs.Count + ' 份文档：') -ForegroundColor Green
$n = 0
foreach ($d in $docs) {
    $n++
    $rel = $d.FullName.Substring($Root.Length).TrimStart('\')
    Write-Host ('  {0,2}. {1}' -f $n, $rel)
}
Write-Host ''

# ---------- 3. 探测可用打印引擎 ----------
# Word.Application 在新版 WPS 里也会注册，所以先试它最省事；
# 不行再试 WPS 原生 ProgID；再不行退回系统文件关联的 print 动词。
function Test-ComProgId {
    param([string]$ProgId)
    try {
        $obj = New-Object -ComObject $ProgId -ErrorAction Stop
        try { $obj.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch { }
        return $true
    } catch {
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

# 强制指定却不可用时给出明确提示
if ($Engine -ne 'Auto' -and $Engine -ne 'Shell') {
    $wanted = $Engine
    if ($chosen -ne $wanted) {
        Write-Host ('警告: 未找到 ' + $wanted + '，已自动降级为系统默认程序。') -ForegroundColor Yellow
        Write-Host '      如需指定引擎，请先安装对应软件，或改用 -Engine Auto' -ForegroundColor Yellow
    }
}

if (-not $chosen) {
    $chosen = 'Shell'
}

$engineLabel = switch ($chosen) {
    'Word'  { 'Microsoft Word (' + $progId + ')' }
    'WPS'   { 'WPS (' + $progId + ')' }
    'Shell' { '系统默认程序（文件关联）' }
}
Write-Host ('打印引擎: ' + $engineLabel) -ForegroundColor Green

# 默认打印机
try {
    $pr = Get-CimInstance -ClassName Win32_Printer -Filter 'Default=TRUE' -ErrorAction Stop
    if ($pr) { Write-Host ('默认打印机: ' + $pr.Name) -ForegroundColor Gray }
    else { Write-Host '警告: 系统未设置默认打印机。' -ForegroundColor Yellow }
} catch {
    Write-Host '警告: 读取打印机信息失败。' -ForegroundColor Yellow
}
Write-Host ''

if ($DryRun) {
    Write-Host '[试运行] 只列出清单，未实际发送打印。' -ForegroundColor Yellow
    exit 0
}

# ---------- 4. 确认 ----------
if (-not $Force) {
    $ans = Read-Host ('确认打印以上 ' + $docs.Count + ' 份? [Y/n]')
    if ($ans -ne '' -and $ans.Trim().ToLower() -notmatch '^(y|yes)$') {
        Write-Host '已取消，未打印任何文件。' -ForegroundColor Yellow
        exit 0
    }
}

# ---------- 5. 打印 ----------
$ok = 0
$fail = 0

if ($chosen -eq 'Shell') {
    # 兜底：交给系统里 .docx 的默认处理程序（Word 或 WPS 都行）
    $n = 0
    foreach ($d in $docs) {
        $n++
        $rel = $d.FullName.Substring($Root.Length).TrimStart('\')
        Write-Host ('[{0}/{1}] {2} ... ' -f $n, $docs.Count, $rel) -NoNewline
        try {
            Start-Process -FilePath $d.FullName -Verb Print -ErrorAction Stop
            Start-Sleep -Milliseconds 1200
            Write-Host '已送打印' -ForegroundColor Green
            $ok++
        } catch {
            Write-Host ('失败: ' + $_.Exception.Message) -ForegroundColor Red
            $fail++
        }
    }
}
else {
    $app = $null
    try {
        $app = New-Object -ComObject $progId -ErrorAction Stop
        try { $app.Visible = $false } catch { }
        try { $app.DisplayAlerts = 0 } catch { }

        $n = 0
        foreach ($d in $docs) {
            $n++
            $rel = $d.FullName.Substring($Root.Length).TrimStart('\')
            Write-Host ('[{0}/{1}] {2} ... ' -f $n, $docs.Count, $rel) -NoNewline
            try {
                # 只读打开，避免误改原文件
                $doc = $app.Documents.Open($d.FullName, $false, $true)
                $doc.PrintOut()
                Start-Sleep -Milliseconds 400
                $doc.Close(0)
                Write-Host '已送打印' -ForegroundColor Green
                $ok++
            } catch {
                Write-Host ('失败: ' + $_.Exception.Message) -ForegroundColor Red
                $fail++
            }
        }
    }
    catch {
        Write-Host ('无法启动 ' + $chosen + ': ' + $_.Exception.Message) -ForegroundColor Red
        Write-Host '请确认已安装 Microsoft Word 或 WPS。' -ForegroundColor Yellow
        $fail = $docs.Count - $ok
    }
    finally {
        if ($app) {
            try { $app.Quit() } catch { }
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
        }
    }
}

Write-Host ''
Write-Host ('完成：成功 ' + $ok + ' 份，失败 ' + $fail + ' 份') -ForegroundColor Cyan

if ($fail -gt 0) { exit 2 } else { exit 0 }
