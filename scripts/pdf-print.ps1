# ============================================================
#  PDF 打印模块 —— 零依赖（只用 Windows 自带组件）
#
#  策略链：
#    1. SumatraPDF / Acrobat / Reader / Foxit 命令行（若装了）
#    2. WinRT Windows.Data.Pdf 渲染成位图 + System.Drawing.Printing 打印
#       <- 系统自带，不装任何东西就能用
#    3. Shell Print 动词
#
#  这个文件被 OA-print.ps1 dot-source 引入。
# ============================================================

# ---------- WinRT 异步辅助 ----------

$script:WinRTAvailable = $null
$script:WinRTAsTask = $null

function Initialize-WinRT {
    if ($null -ne $script:WinRTAvailable) { return $script:WinRTAvailable }

    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop

        $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() |
            Where-Object {
                $_.Name -eq 'AsTask' -and
                $_.GetParameters().Count -eq 1 -and
                $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
            } | Select-Object -First 1)

        if (-not $asTaskGeneric) {
            $script:WinRTAvailable = $false
            return $false
        }

        # IAsyncAction（无返回值）用的是非泛型 AsTask，单独缓存
        $asTaskAction = ([System.WindowsRuntimeSystemExtensions].GetMethods() |
            Where-Object {
                $_.Name -eq 'AsTask' -and
                (-not $_.IsGenericMethod) -and
                $_.GetParameters().Count -eq 1 -and
                $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'
            } | Select-Object -First 1)

        $script:WinRTAsTask = $asTaskGeneric
        $script:WinRTAsTaskAction = $asTaskAction
        $script:WinRTAvailable = $true
        return $true
    } catch {
        $script:WinRTAvailable = $false
        return $false
    }
}

function Await-WinRT {
    param($AsyncOp, [Type]$ResultType)

    if (-not (Initialize-WinRT)) { throw 'WinRT 不可用' }

    $m = $script:WinRTAsTask.MakeGenericMethod($ResultType)
    $task = $m.Invoke($null, @($AsyncOp))
    $task.Wait(-1) | Out-Null
    return $task.Result
}

# 非泛型版：等 IAsyncAction（无返回值）
function Await-WinRTAction {
    param($AsyncAction)

    if (-not (Initialize-WinRT)) { throw 'WinRT 不可用' }
    if (-not $script:WinRTAsTaskAction) { throw 'AsTask(IAsyncAction) 不可用' }

    $task = $script:WinRTAsTaskAction.Invoke($null, @($AsyncAction))
    $task.Wait(-1) | Out-Null
    return
}

# ---------- PDF -> 位图（逐页） ----------

function Convert-PdfToBitmaps {
    param(
        [string]$PdfPath,
        [int]$Dpi = 200,
        [int]$MaxPages = 50
    )

    if (-not (Initialize-WinRT)) {
        throw 'WinRT PDF 渲染不可用'
    }

    # 载入 PDF（WinRT 需要 StorageFile）
    Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction SilentlyContinue
    $storageFileType = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
    $file = Await-WinRT ($storageFileType::GetFileFromPathAsync($PdfPath)) ([Windows.Storage.StorageFile])

    $pdfDocType = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType=WindowsRuntime]
    $pdf = Await-WinRT ($pdfDocType::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])

    $count = [int]$pdf.PageCount
    if ($count -gt $MaxPages) { $count = $MaxPages }

    $list = New-Object System.Collections.ArrayList

    for ($i = 0; $i -lt $count; $i++) {
        $page = $pdf.GetPage($i)
        try {
            $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
            $scale = $Dpi / 96.0
            $opts.DestinationWidth  = [uint32]([math]::Round($page.Size.Width  * $scale))
            $opts.DestinationHeight = [uint32]([math]::Round($page.Size.Height * $scale))

            # 渲染到内存流
            $ms = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream

            $op = $page.RenderToStreamAsync($ms, $opts)
            Await-WinRTAction $op

            # 读回字节
            $size = [int]$ms.Size
            $ms.Seek(0)
            $reader = New-Object Windows.Storage.Streams.DataReader($ms.GetInputStreamAt(0))
            $readOp = $reader.LoadAsync([uint32]$size)
            $n = Await-WinRT $readOp ([uint32])

            $buffer = New-Object byte[] $n
            $reader.ReadBytes($buffer)
            $reader.Dispose()
            $ms.Dispose()

            $msImg = New-Object System.IO.MemoryStream(,$buffer)
            $img = [System.Drawing.Image]::FromStream($msImg)
            [void]$list.Add($img)
        } finally {
            $page.Dispose()
        }
    }

    return @($list)
}

# ---------- 用 System.Drawing 打印位图 ----------

function Get-PrinterPaperSize {
    param([string]$PrinterName)
    try {
        $ps = New-Object System.Drawing.Printing.PrinterSettings
        if ($PrinterName -and $PrinterName -notlike '(*') { $ps.PrinterName = $PrinterName }
        return $ps.DefaultPageSettings.PaperSize
    } catch {
        return $null
    }
}

function Print-Bitmaps {
    param(
        [array]$Bitmaps,
        [string]$PrinterName,
        [switch]$Landscape
    )

    if (-not $Bitmaps -or $Bitmaps.Count -eq 0) {
        throw '没有可打印的图像'
    }

    Add-Type -AssemblyName System.Drawing

    $doc = New-Object System.Drawing.Printing.PrintDocument
    if ($PrinterName -and $PrinterName -notlike '(*') {
        $doc.PrinterSettings.PrinterName = $PrinterName
    }
    $doc.DocumentName = 'OA-PDF'

    $script:__imgs = $Bitmaps
    $script:__idx = 0
    $script:__landscape = [bool]$Landscape

    $doc.add_PrintPage({
        param($sender, $e)

        $img = $script:__imgs[$script:__idx]
        $bounds = $e.MarginBounds
        $areaW = $bounds.Width
        $areaH = $bounds.Height

        $iw = [double]$img.Width
        $ih = [double]$img.Height

        # 按比例缩放到可打印区域
        $scale = [Math]::Min($areaW / $iw, $areaH / $ih)
        $dw = [int]($iw * $scale)
        $dh = [int]($ih * $scale)
        $x = $bounds.X + [int](($areaW - $dw) / 2)
        $y = $bounds.Y + [int](($areaH - $dh) / 2)

        $e.Graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $e.Graphics.DrawImage($img, $x, $y, $dw, $dh)

        $script:__idx++
        $e.HasMorePages = ($script:__idx -lt $script:__imgs.Count)
    })

    try {
        if ($script:__landscape) {
            $doc.DefaultPageSettings.Landscape = $true
        }
        $doc.Print()
    } finally {
        try { $doc.Dispose() } catch { }
        foreach ($im in $Bitmaps) { try { $im.Dispose() } catch { } }
    }
}

# ---------- 判断 PDF 是否横版 ----------

function Test-PdfLandscape {
    param([string]$PdfPath)

    try {
        $bytes = [System.IO.File]::ReadAllBytes($PdfPath)
        $txt = [System.Text.Encoding]::ASCII.GetString($bytes)
        # 找 MediaBox [0 0 W H]
        $m = [regex]::Matches($txt, 'MediaBox\s*\[\s*([\d\.\-]+)\s+([\d\.\-]+)\s+([\d\.\-]+)\s+([\d\.\-]+)\s*\]')
        if ($m.Count -gt 0) {
            $last = $m[$m.Count - 1]
            $w = [double]$last.Groups[3].Value - [double]$last.Groups[1].Value
            $h = [double]$last.Groups[4].Value - [double]$last.Groups[2].Value
            return ($w -gt $h)
        }
    } catch { }
    return $false
}

# ---------- 主入口：打印一个 PDF ----------

function Invoke-PrintPdfNative {
    param(
        [string]$Path,
        [string]$Printer,
        [int]$Dpi = 200
    )

    if (-not (Initialize-WinRT)) {
        return [pscustomobject]@{ Ok=$false; Strategy='WinRT 不可用'; Error=$null }
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        return [pscustomobject]@{ Ok=$false; Strategy='文件不存在'; Error=$null }
    }

    $imgs = $null
    try {
        $imgs = Convert-PdfToBitmaps -PdfPath $Path -Dpi $Dpi
        if (-not $imgs -or $imgs.Count -eq 0) {
            return [pscustomobject]@{ Ok=$false; Strategy='渲染失败'; Error=$null }
        }

        $landscape = Test-PdfLandscape -PdfPath $Path
        Print-Bitmaps -Bitmaps $imgs -PrinterName $Printer -Landscape:$landscape
        Start-Sleep -Milliseconds 400
        return [pscustomobject]@{ Ok=$true; Strategy='WinRT渲染+System.Drawing'; Error=$null }
    } catch {
        return [pscustomobject]@{ Ok=$false; Strategy='WinRT 异常'; Error=$_ }
    } finally {
        if ($imgs) {
            foreach ($im in $imgs) { try { $im.Dispose() } catch { } }
        }
    }
}
