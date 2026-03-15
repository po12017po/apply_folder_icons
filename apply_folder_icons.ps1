Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore

$root = $PSScriptRoot
$dirs = Get-ChildItem -LiteralPath $root -Directory

Write-Host "Processing folders (Root images only)..." -ForegroundColor Cyan

foreach ($dir in $dirs) {
    # 移除 -Recurse，只抓取第一層的圖片
    $images = Get-ChildItem -LiteralPath $dir.FullName -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png|webp|bmp)$' } | Sort-Object Name

    if (-not $images) {
        Write-Host "Skipping $($dir.Name) (no image found in root)" -ForegroundColor Yellow
        continue
    }

    $success = $false
    foreach ($firstImage in $images) {
        try {
            if ($firstImage.Extension -match '(?i)\.webp$') {
                # 針對 WebP，使用 WPF 的 BitmapDecoder 來讀取並轉換為 PNG，然後再讓 System.Drawing 讀取
                $stream = New-Object System.IO.FileStream($firstImage.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
                $decoder = [System.Windows.Media.Imaging.BitmapDecoder]::Create($stream, [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat, [System.Windows.Media.Imaging.BitmapCacheOption]::Default)
                $bitmapFrame = $decoder.Frames[0]

                $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
                $encoder.Frames.Add($bitmapFrame)

                $memoryStream = New-Object System.IO.MemoryStream
                $encoder.Save($memoryStream)
                
                # 讓 stream 不關閉，以免後續 System.Drawing 讀取不到
                $memoryStream.Position = 0
                $img = [System.Drawing.Image]::FromStream($memoryStream)
                
                $stream.Dispose()
                # 注意：$memoryStream 需要等 $img.Dispose() 後再關閉
            } else {
                $img = [System.Drawing.Image]::FromFile($firstImage.FullName)
            }
        } catch {
            $errMsg = $_.Exception.Message
            
            # 如果是傳統讀取失敗，但它是 WebP (可能是因為我們不在 WPF 環境或解碼器失敗)
            if ($firstImage.Extension -match '(?i)\.webp$') {
                 Write-Host "  -> 略過 $($firstImage.Name) (WebP 解碼失敗，請確認系統是否有安裝 WebP 影像延伸模組: $errMsg)" -ForegroundColor DarkYellow
            } elseif ($_.Exception -is [System.OutOfMemoryException]) {
                Write-Host "  -> 略過 $($firstImage.Name) (不支援的格式或記憶體不足: $errMsg)" -ForegroundColor DarkYellow
            } else {
                Write-Host "  -> 略過 $($firstImage.Name) (無法讀取: $errMsg)" -ForegroundColor DarkYellow
            }
            
            if ($null -ne $stream) { $stream.Dispose() }
            if ($null -ne $memoryStream) { $memoryStream.Dispose() }
            
            continue
        }

        Write-Host "Processing: $($dir.Name) -> using $($firstImage.Name)"
        $iniPath = Join-Path $dir.FullName "desktop.ini"

        try {
            # 依圖片內容產生 icon 檔名，圖片內容變更時檔名也會跟著變
            $iconHash = (Get-FileHash -LiteralPath $firstImage.FullName -Algorithm SHA256).Hash.Substring(0, 16).ToLowerInvariant()
            $iconFileName = "foldericon_$iconHash.ico"
            $iconPath = Join-Path $dir.FullName $iconFileName

            # 如果存在目前要覆寫的檔案，先移除隱藏與系統屬性，以免無法覆寫
            if (Test-Path -LiteralPath $iconPath) {
                $icoItem = Get-Item -LiteralPath $iconPath -Force
                $icoItem.Attributes = [System.IO.FileAttributes]::Normal
            }

            if (Test-Path -LiteralPath $iniPath) {
                $iniItem = Get-Item -LiteralPath $iniPath -Force
                $iniItem.Attributes = [System.IO.FileAttributes]::Normal
            }

            # 自動刪除舊的 foldericon_*.ico，也順便清理舊版固定檔名 foldericon.ico
            Get-ChildItem -LiteralPath $dir.FullName -File -Force |
                Where-Object { $_.Name -like 'foldericon_*.ico' -or $_.Name -eq 'foldericon.ico' } |
                ForEach-Object {
                    if ($_.Name -ne $iconFileName) {
                        $_.Attributes = [System.IO.FileAttributes]::Normal
                        Remove-Item -LiteralPath $_.FullName -Force
                    }
                }

            # 圖片處理
            $size = 256
            $bmp = New-Object System.Drawing.Bitmap($size, $size)
            $bmp.SetResolution($img.HorizontalResolution, $img.VerticalResolution)
            $graphics = [System.Drawing.Graphics]::FromImage($bmp)
            
            $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $graphics.Clear([System.Drawing.Color]::Transparent)

            $ratio = $img.Width / $img.Height
            if ($ratio -gt 1) {
                $w = $size
                $h = [math]::Round($size / $ratio)
                $x = 0
                $y = [math]::Round(($size - $h) / 2)
            } else {
             $h = $size
                $w = [math]::Round($size * $ratio)
                $y = 0
                $x = [math]::Round(($size - $w) / 2)
            }

            $graphics.DrawImage($img, $x, $y, $w, $h)
            
            $ms = New-Object System.IO.MemoryStream
            $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
            $pngData = $ms.ToArray()
            
            $graphics.Dispose()
            $bmp.Dispose()
            $img.Dispose()
            $ms.Dispose()
            if ($null -ne $memoryStream) { $memoryStream.Dispose() }

            # 寫入 ico
            $fs = [System.IO.File]::Create($iconPath)
            $bw = New-Object System.IO.BinaryWriter($fs)
            $bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]1)
            $bw.Write([byte]0); $bw.Write([byte]0); $bw.Write([byte]0); $bw.Write([byte]0)
            $bw.Write([uint16]1); $bw.Write([uint16]32)
            $bw.Write([uint32]$pngData.Length); $bw.Write([uint32]22)
            $bw.Write($pngData)
            $bw.Dispose(); $fs.Dispose()

        $icoItem = Get-Item -LiteralPath $iconPath -Force
        $icoItem.Attributes = [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System

        # 寫入 desktop.ini (使用 ASCII 避免編碼 BOM 干擾 Windows 讀取，並指向新的 icon 檔名)
        $iniContent = "[.ShellClassInfo]`r`nIconResource=$iconFileName,0`r`n"
        [System.IO.File]::WriteAllText($iniPath, $iniContent, [System.Text.Encoding]::ASCII)

        $iniItem = Get-Item -LiteralPath $iniPath -Force
        $iniItem.Attributes = [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System

        # 設定資料夾屬性，並且更新修改時間來強制 Windows 重新整理快取
        try {
            $dirItem = Get-Item -LiteralPath $dir.FullName -Force
            $dirItem.Attributes = $dirItem.Attributes -bor [System.IO.FileAttributes]::ReadOnly
            $dirItem.LastWriteTime = (Get-Date)
        } catch {
            Write-Host "  -> 警告: 無法更新資料夾時間 (可能被其他程式鎖定)，但不影響圖示套用" -ForegroundColor Yellow
        }

        $success = $true
        break # 成功處理，跳出圖片迴圈

        } catch {
            Write-Host "  -> 處理圖片時發生錯誤 ($($firstImage.Name)): $_" -ForegroundColor Red
            if ($null -ne $img) { $img.Dispose() }
        }
    }

    if (-not $success) {
        Write-Host "Skipping $($dir.Name) (找不到可用或支援的圖片)" -ForegroundColor Yellow
    }
}

Write-Host "All done! Restarting Explorer to refresh icon cache..." -ForegroundColor Green

# 紀錄目前開啟的檔案總管視窗路徑
$openPaths = @()
try {
    $shell = New-Object -ComObject Shell.Application
    foreach ($window in $shell.Windows()) {
        if ($window.FullName -match 'explorer\.exe$') {
            $path = $window.Document.Folder.Self.Path
            if ($path -and (Test-Path -LiteralPath $path)) {
                $openPaths += $path
            }
        }
    }
} catch {
    Write-Host "Warning: Failed to get open explorer windows. $($_.Exception.Message)" -ForegroundColor Yellow
}

# 重新啟動 explorer
Stop-Process -Name explorer -Force
Start-Sleep -Seconds 2

# 重新開啟剛剛的資料夾
if ($openPaths.Count -gt 0) {
    $openPaths = $openPaths | Select-Object -Unique
    foreach ($path in $openPaths) {
        Invoke-Item -LiteralPath $path
    }
}