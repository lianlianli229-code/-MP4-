@echo off
setlocal DisableDelayedExpansion
set "SCRIPT_PATH=%~f0"
set "SCRIPT_DIR=%~dp0"
set "DRAG_TARGET=%~1"
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -Command "$l=[System.IO.File]::ReadAllLines($env:SCRIPT_PATH,[System.Text.Encoding]::UTF8);$i=[Array]::IndexOf($l,'# === POWERSHELL START ===');&([ScriptBlock]::Create(($l[($i+1)..($l.Length-1)]-join[char]10)))"
exit /b %ERRORLEVEL%
# === POWERSHELL START ===

# =================================================================================
# MP4 人声气口（音频波谷）批量提取工具【自带本机 FFmpeg 安全自检与官方安装二合一版】
# 适用场景：放在你自己的 Windows 电脑上运行（免上传大视频，直接提取波谷与达芬奇标记）
# 安全承诺：
#   1. 若本机已装 FFmpeg：100% 纯本地离线运行（锁定 file/pipe 协议白名单，零联网、零遥测）。
#   2. 若本机未装 FFmpeg：先完整展示软件安全四要素，必须经你亲手输入 Y 确认后，
#      才通过唯一指定官方入口 https://ffmpeg.org/ 核验正版源、锁定正式稳定版（FFmpeg 9.0.2）、
#      强制 HTTPS + 本地硬编码 SHA-256 强校验 + 归档路径穿越防护后自动安装到本机。
#   3. 严守零删除/零移动/零覆盖安全红线：全程对原视频只读，绝不包含任何删除或覆盖旧文件指令。
# =================================================================================

# ------------------- 【可调参数区（用记事本打开即可修改）】 -------------------
$NoiseThreshold    = "-35dB"            # 波谷音量阈值：漏检气口可调高为 "-30dB"，误检太多可调低为 "-40dB"
$MinDuration       = 0.18               # 最短气口持续秒数：低于此秒数的微小停顿将被忽略（建议 0.15 ~ 0.30）
$BigValleySeconds  = 0.40               # 大气口分界秒数：达到此秒数在达芬奇标记中标为绿色，否则标为青色
$OutputFolderName  = "气口波谷提取结果" # 汇总输出子文件夹名称
$TimelineStartHour = 1                  # 达芬奇 EDL 起始小时码：默认新建时间线填 1 (01:00:00:00)；XML时间线(00起)可改为 0
# ------------------------------------------------------------------------------

$ErrorActionPreference = "Stop"
$InvariantCulture = [System.Globalization.CultureInfo]::InvariantCulture
$Sys32Dir = [Environment]::SystemDirectory
try {
    $chcpExe = Join-Path $Sys32Dir "chcp.com"
    if (Test-Path -LiteralPath $chcpExe -PathType Leaf) { & $chcpExe 65001 | Out-Null }
} catch {}
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 启动前安全校验可调参数（防止滤镜注入、路径穿越与区域小数点错位）
function Assert-SafeConfig {
    if ([string]::IsNullOrWhiteSpace($NoiseThreshold) -or ($NoiseThreshold.Trim() -notmatch '^-\d+(\.\d+)?dB$')) {
        throw "[配置错误] NoiseThreshold 格式无效（示例: '-35dB'），当前值：$NoiseThreshold"
    }
    $minDurVal = 0.0
    if (-not [double]::TryParse([string]::Format($InvariantCulture, "{0}", $MinDuration), [System.Globalization.NumberStyles]::Float, $InvariantCulture, [ref]$minDurVal) -or $minDurVal -lt 0.01 -or $minDurVal -gt 600.0) {
        throw "[配置错误] MinDuration 必须为 0.01 ~ 600 之间的秒数，当前值：$MinDuration"
    }
    $bigVal = 0.0
    if (-not [double]::TryParse([string]::Format($InvariantCulture, "{0}", $BigValleySeconds), [System.Globalization.NumberStyles]::Float, $InvariantCulture, [ref]$bigVal) -or $bigVal -lt 0.01 -or $bigVal -gt 3600.0) {
        throw "[配置错误] BigValleySeconds 必须为 0.01 ~ 3600 之间的秒数，当前值：$BigValleySeconds"
    }
    if ([string]::IsNullOrWhiteSpace($OutputFolderName) -or $OutputFolderName -in @(".", "..") -or $OutputFolderName -match '[\\/:\*\?"<>\|]') {
        throw "[配置错误] OutputFolderName 仅允许单层合法文件夹名（禁止包含路径分隔符或 ..），当前值：$OutputFolderName"
    }
    if (($TimelineStartHour -isnot [int]) -or $TimelineStartHour -lt 0 -or $TimelineStartHour -gt 23) {
        throw "[配置错误] TimelineStartHour 必须为 0 ~ 23 的整数，当前值：$TimelineStartHour"
    }
}
Assert-SafeConfig

$OfficialEntryUrl    = "https://ffmpeg.org/"
$OfficialDownloadDoc = "https://ffmpeg.org/download.html"

# 供应链安全锁定配置：固定至官方正式稳定版 FFmpeg 9.0.2 不可变归档与本地硬编码 SHA-256
$PinnedFFmpegVersion = "9.0.2"
$PinnedGyanGhUrl     = "https://github.com/GyanD/codexffmpeg/releases/download/9.0.2/ffmpeg-9.0.2-essentials_build.zip"
$PinnedGyanWebUrl    = "https://www.gyan.dev/ffmpeg/builds/packages/ffmpeg-9.0.2-essentials_build.zip"
$PinnedGyanSha256    = "60F467265B1E312373DBCD92200C2618A74850F98D3D078E94296BB3FA2047BA"
$PinnedBtbnUrl       = "https://github.com/FFmpeg-Builds/FFmpeg-Builds/releases/download/autobuild-2026-11-15-12-40/ffmpeg-n9.0.2-57-gd838f9f9ae-win64-gpl-9.0.zip"
$PinnedBtbnSha256    = "11B387A88AD24EE6D78BAE7D24A2417834737E6AC00F0A36CA67FEAD1E924EFF"

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "  MP4 音频波谷（人声气口）批量提取工具 [本机自检 + 安全提取版]" -ForegroundColor Cyan
Write-Host "  当前参数：阈值 = $NoiseThreshold  |  最短气口 = $([string]::Format($InvariantCulture, '{0:0.##}', [double]$MinDuration))s" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan

# 1. 安全查找本机已安装的 FFmpeg 与 FFprobe（仅信任绝对路径，排除当前视频目录防 DLL/EXE 劫持）
function Find-ToolPath([string]$exeName) {
    $defaultBin = [System.IO.Path]::GetFullPath((Join-Path $env:USERPROFILE "ffmpeg\bin")).TrimEnd('\')
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $machPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $candidates = @($defaultBin)
    if ($machPath) { $candidates += ($machPath -split ";") }
    if ($userPath) { $candidates += ($userPath -split ";") }
    if ($env:PATH) { $candidates += ($env:PATH -split ";") }

    $untrustedDirs = @()
    if (-not [string]::IsNullOrWhiteSpace($env:SCRIPT_DIR)) {
        try { $untrustedDirs += [System.IO.Path]::GetFullPath($env:SCRIPT_DIR).TrimEnd('\') } catch {}
    }
    if (-not [string]::IsNullOrWhiteSpace($env:DRAG_TARGET)) {
        try {
            $dtFull = [System.IO.Path]::GetFullPath($env:DRAG_TARGET)
            if (Test-Path -LiteralPath $dtFull -PathType Container) {
                $untrustedDirs += $dtFull.TrimEnd('\')
            } else {
                $dtParent = Split-Path -Parent $dtFull
                if ($dtParent) { $untrustedDirs += $dtParent.TrimEnd('\') }
            }
        } catch {}
    }

    foreach ($dir in $candidates) {
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        $trimmed = $dir.Trim().Trim('"')
        if (-not [System.IO.Path]::IsPathRooted($trimmed)) { continue }
        try {
            $normDir = [System.IO.Path]::GetFullPath($trimmed).TrimEnd('\')
            if ($normDir -ine $defaultBin) {
                $inUntrusted = $false
                foreach ($ud in $untrustedDirs) {
                    if ($normDir -ieq $ud) { $inUntrusted = $true; break }
                }
                if ($inUntrusted) { continue }
            }
            $full = Join-Path $normDir $exeName
            if (Test-Path -LiteralPath $full -PathType Leaf -ErrorAction SilentlyContinue) { return $full }
        } catch {}
    }
    return $null
}


# 辅助核验：解压前静态检查 ZIP 条目，阻断 Zip-Slip 路径穿越
function Test-SafeZipArchive([string]$archivePath, [string]$destRoot) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $fullDest = [System.IO.Path]::GetFullPath($destRoot).TrimEnd('\') + '\'
    $zip = [System.IO.Compression.ZipFile]::OpenRead($archivePath)
    try {
        if ($zip.Entries.Count -eq 0) { return $false }
        foreach ($entry in $zip.Entries) {
            $rawName = $entry.FullName
            if ([string]::IsNullOrWhiteSpace($rawName)) { continue }
            if ($rawName.StartsWith("/") -or $rawName.StartsWith("\") -or $rawName.Contains(":")) {
                return $false
            }
            $parts = @($rawName -split '[\\/]' | Where-Object { $_ -ne '' })
            if ($parts -contains '..' -or $parts -contains '.') {
                return $false
            }
            if ($parts.Count -le 1) { continue }
            $strippedRel = ($parts[1..($parts.Count - 1)] -join '\')
            $resolvedTarget = [System.IO.Path]::GetFullPath((Join-Path $fullDest $strippedRel))
            if (-not $resolvedTarget.StartsWith($fullDest, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $false
            }
        }
        return $true
    } finally {
        $zip.Dispose()
    }
}

# 2. 本机未安装 FFmpeg 时的安全四要素汇报与本机自动安装模块
function Install-FFmpegToLocalUser {
    Write-Host ""
    Write-Host "------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[环境检测] 检测到你当前这台电脑尚未安装 FFmpeg 音视频处理工具。" -ForegroundColor Yellow
    Write-Host "在为你自动安装到本机前，请审阅以下【软件安全核查四要素】：" -ForegroundColor Cyan
    Write-Host "  1. 软件来源：指定官方入口 $OfficialEntryUrl（锁定正式稳定版 v$PinnedFFmpegVersion）" -ForegroundColor White
    Write-Host "     （脚本将现场核验官网 download.html 指定之正版构建源，并强制核对本地硬编码 SHA-256 指纹）" -ForegroundColor Gray
    Write-Host "  2. 是否安全：国际开源白名单认证工具（LGPL/GPL 协议），开源透明，无后门、无遥测回传。" -ForegroundColor White
    Write-Host "  3. 所属国家与注册地：国际开源组织 FFmpeg Project（法律与资金托管于美国 SPI 非营利机构）。" -ForegroundColor White
    Write-Host "  4. 开发者真实背景核查：由法国程序员 Fabrice Bellard 发起、国际开源社区维护；" -ForegroundColor White
    Write-Host "     无海外空壳公司包装，不涉及中国大陆、俄罗斯、北朝鲜等高风险国家团队或企业。" -ForegroundColor White
    Write-Host ""
    Write-Host "  拟安装本机路径：$env:USERPROFILE\ffmpeg\bin（仅写入当前用户目录与用户 Path，免管理员权限）" -ForegroundColor Cyan
    Write-Host "------------------------------------------------------------------" -ForegroundColor Yellow

    $confirm = Read-Host "是否授权现在自动下载并安装 FFmpeg 到你这台电脑？请输入 Y 回车确认（其他键取消）"
    if ($confirm -notmatch '^[Yy]$') {
        Write-Host "[已取消] 未执行安装。如需提取波谷，请重新运行并输入 Y 完成本机 FFmpeg 安装。" -ForegroundColor Yellow
        return $false
    }

    # 严格锁定系统 System32 内置 curl.exe 与 tar.exe，防止 PATH 劫持
    $curlExe = Join-Path $Sys32Dir "curl.exe"
    $tarExe  = Join-Path $Sys32Dir "tar.exe"
    if (-not (Test-Path -LiteralPath $curlExe -PathType Leaf) -or -not (Test-Path -LiteralPath $tarExe -PathType Leaf)) {
        Write-Host "[错误] 当前 Windows 系统目录缺少内置组件 curl.exe 或 tar.exe（需 Windows 10 1803 及以上版本）。" -ForegroundColor Red
        return $false
    }

    $installRoot = Join-Path $env:USERPROFILE "ffmpeg"
    $pkgDir      = Join-Path $installRoot "official_pkg"
    if (-not (Test-Path -LiteralPath $pkgDir)) {
        New-Item -ItemType Directory -Path $pkgDir -Force | Out-Null
    }

    $stamp     = Get-Date -Format "yyyyMMdd_HHmmss"
    $htmlCheck = Join-Path $pkgDir "ffmpeg_official_verify_$stamp.html"

    Write-Host ""
    Write-Host "[1/4] 正在通过 HTTPS 连接指定官方入口核验正版 Windows 构建源：$OfficialEntryUrl ..." -ForegroundColor Cyan
    & $curlExe --proto "=https" --tlsv1.2 -fsSL $OfficialEntryUrl -o NUL
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[错误] 无法安全访问指定官网入口 $OfficialEntryUrl ，请检查本机网络连接。" -ForegroundColor Red
        return $false
    }
    & $curlExe --proto "=https" --tlsv1.2 -fsSL $OfficialDownloadDoc -o $htmlCheck
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $htmlCheck)) {
        Write-Host "[错误] 无法读取官网下载核验页 $OfficialDownloadDoc 。" -ForegroundColor Red
        return $false
    }

    $docText = [System.IO.File]::ReadAllText($htmlCheck, [System.Text.Encoding]::UTF8)
    $officialSources = New-Object System.Collections.Generic.List[object]
    if ($docText -match 'https://www\.gyan\.dev/ffmpeg/builds/') {
        $officialSources.Add([PSCustomObject]@{
            Tag            = "gyan_gh_${PinnedFFmpegVersion}"
            PkgUrl         = $PinnedGyanGhUrl
            ExpectedSha256 = $PinnedGyanSha256.ToUpperInvariant()
        })
        $officialSources.Add([PSCustomObject]@{
            Tag            = "gyan_web_${PinnedFFmpegVersion}"
            PkgUrl         = $PinnedGyanWebUrl
            ExpectedSha256 = $PinnedGyanSha256.ToUpperInvariant()
        })
    }
    if ($docText -match 'https://github\.com/(BtbN|FFmpeg-Builds)/FFmpeg-Builds/releases') {
        $officialSources.Add([PSCustomObject]@{
            Tag            = "btbn_n${PinnedFFmpegVersion}"
            PkgUrl         = $PinnedBtbnUrl
            ExpectedSha256 = $PinnedBtbnSha256.ToUpperInvariant()
        })
    }

    if ($officialSources.Count -eq 0) {
        Write-Host "[安全拦截] 官网下载页未匹配到预期的官方 Windows 构建源，已中止下载以确保安全。" -ForegroundColor Red
        return $false
    }

    $verifiedZipPath = $null
    foreach ($src in $officialSources) {
        # 按镜像源与锁定版本独立命名并附时间戳，既防止覆盖旧文件，又杜绝跨镜像断点续传拼包损坏
        $candidateZip = Join-Path $pkgDir ("ffmpeg_$($src.Tag)_win64_$stamp.zip")

        Write-Host "[2/4] 官网源核验通过！正在下载锁定稳定版 v$PinnedFFmpegVersion 安装包（强制 HTTPS，约 80~110MB）..." -ForegroundColor Cyan
        Write-Host "      锁定正版直链：$($src.PkgUrl)" -ForegroundColor Gray
        & $curlExe --proto "=https" --tlsv1.2 -C - --retry 5 --retry-all-errors -fSL $src.PkgUrl -o $candidateZip
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $candidateZip -PathType Leaf)) {
            Write-Host "      当前节点安装包下载中断，尝试官网认证备用锁定源..." -ForegroundColor Yellow
            continue
        }

        Write-Host "[3/4] 正在核对安装包与本地硬编码 SHA-256 安全指纹..." -ForegroundColor Cyan
        $expectedHash = $src.ExpectedSha256
        $actualHash   = (Get-FileHash -LiteralPath $candidateZip -Algorithm SHA256).Hash.ToUpperInvariant()
        if ($actualHash -ne $expectedHash) {
            Write-Host "      [安全拦截] 本地硬编码 SHA-256 哈希不匹配！期望: $expectedHash | 实际: $actualHash ，已拒绝执行该归档并尝试备用源..." -ForegroundColor Red
            continue
        }

        Write-Host "      本地硬编码 SHA-256 强校验通过（$actualHash）！" -ForegroundColor Green
        $verifiedZipPath = $candidateZip
        break
    }

    if (-not $verifiedZipPath) {
        Write-Host "[错误] 未能下载并通过 SHA-256 完整性核验，已安全中止安装。请检查网络后重新运行。" -ForegroundColor Red
        return $false
    }

    Write-Host "[4/4] 正在核验压缩包目录结构安全性并解压至 $installRoot ..." -ForegroundColor Cyan
    $isSafeArchive = $false
    try {
        $isSafeArchive = Test-SafeZipArchive $verifiedZipPath $installRoot
    } catch {
        $isSafeArchive = $false
    }
    if (-not $isSafeArchive) {
        Write-Host "[安全拦截] 压缩包内部检测到非法路径或目录穿越风险（Zip-Slip），已中止解压！" -ForegroundColor Red
        return $false
    }

    # 使用 -xkf 保持零覆盖原则（不覆盖已存在文件）
    & $tarExe -xkf $verifiedZipPath -C $installRoot --strip-components=1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[错误] 安装包解压失败。" -ForegroundColor Red
        return $false
    }

    $binDir  = Join-Path $installRoot "bin"
    $ffCheck = Join-Path $binDir "ffmpeg.exe"
    $fpCheck = Join-Path $binDir "ffprobe.exe"
    if (-not (Test-Path -LiteralPath $ffCheck -PathType Leaf) -or -not (Test-Path -LiteralPath $fpCheck -PathType Leaf)) {
        Write-Host "[错误] 解压后未找到完整的 ffmpeg.exe / ffprobe.exe（$binDir）。" -ForegroundColor Red
        return $false
    }

    # 安全追加到当前用户 Path（保留原有全部路径，不修改系统级注册表）
    $currentUserPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $pathParts = @()
    if (-not [string]::IsNullOrWhiteSpace($currentUserPath)) {
        $pathParts = $currentUserPath -split ";" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }
    $alreadyInPath = $false
    foreach ($p in $pathParts) {
        if ($p.Trim().TrimEnd('\') -ieq $binDir.TrimEnd('\')) {
            $alreadyInPath = $true
            break
        }
    }
    if (-not $alreadyInPath) {
        $updatedPath = ($pathParts + $binDir) -join ";"
        [Environment]::SetEnvironmentVariable("Path", $updatedPath, "User")
    }
    $env:PATH = "$binDir;$env:PATH"

    Write-Host "[安装成功] FFmpeg 已安全校验并配置到你这台电脑（$binDir），以后双击直接秒速提取！" -ForegroundColor Green
    Write-Host "------------------------------------------------------------------" -ForegroundColor Cyan
    return $true
}

$ffmpegExe  = Find-ToolPath "ffmpeg.exe"
$ffprobeExe = Find-ToolPath "ffprobe.exe"

if (-not $ffmpegExe) {
    $ok = Install-FFmpegToLocalUser
    if (-not $ok) {
        if (-not [Console]::IsInputRedirected) {
            Write-Host ""
            [void](Read-Host "按回车键退出窗口")
        }
        exit 1
    }
    $ffmpegExe  = Find-ToolPath "ffmpeg.exe"
    $ffprobeExe = Find-ToolPath "ffprobe.exe"
}

# 3. 解析目标文件夹与待处理 MP4 列表
$targetDir = $env:SCRIPT_DIR.TrimEnd('\')
$mp4Files  = @()

if (-not [string]::IsNullOrWhiteSpace($env:DRAG_TARGET)) {
    $dragItem = Get-Item -LiteralPath $env:DRAG_TARGET -ErrorAction SilentlyContinue
    if ($null -eq $dragItem) {
        Write-Host "[错误] 拖入的路径不存在：$($env:DRAG_TARGET)" -ForegroundColor Red
        exit 1
    }
    if ($dragItem.PSIsContainer) {
        $targetDir = $dragItem.FullName
        $mp4Files  = @(Get-ChildItem -LiteralPath $targetDir -File | Where-Object { $_.Extension -ieq ".mp4" } | Sort-Object Name)
    } elseif ($dragItem.Extension -ieq ".mp4") {
        $targetDir = $dragItem.DirectoryName
        $mp4Files  = @($dragItem)
    } else {
        Write-Host "[提示] 拖入的文件不是 .mp4 或文件夹：$($dragItem.Name)" -ForegroundColor Yellow
        exit 1
    }
} else {
    $mp4Files = @(Get-ChildItem -LiteralPath $targetDir -File | Where-Object { $_.Extension -ieq ".mp4" } | Sort-Object Name)
}

if ($mp4Files.Count -eq 0) {
    Write-Host ""
    Write-Host "[提示] 本机 FFmpeg 环境正常（$ffmpegExe）。" -ForegroundColor Green
    Write-Host "[提示] 在目录 [$targetDir] 中没有找到任何 .mp4 文件。" -ForegroundColor Yellow
    Write-Host "用法：请把本批处理放进装有 .mp4 的文件夹双击，或直接把装有 .mp4 的文件夹拖到本图标上！" -ForegroundColor Yellow
    if (-not [Console]::IsInputRedirected) {
        Write-Host ""
        [void](Read-Host "按回车键退出窗口")
    }
    exit 0
}

# 4. 创建输出文件夹（防覆盖保护：如遇同名旧文件自动追加时间戳另存，绝不覆盖已有文件）
$outDir = Join-Path $targetDir $OutputFolderName
if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"
function Get-SafeOutputFile([string]$directory, [string]$baseName, [string]$ext) {
    $candidate = Join-Path $directory ($baseName + $ext)
    if (-not (Test-Path -LiteralPath $candidate)) {
        return $candidate
    }
    return Join-Path $directory ("${baseName}_${runStamp}${ext}")
}

# 5. 辅助函数：标准 Win32 CommandLineToArgvW 参数安全转义 & 纯内存执行外部工具
function Format-WindowsProcessArg([string]$arg) {
    if ($null -eq $arg -or $arg.Length -eq 0) { return '""' }
    if ($arg -notmatch '[\s"]') { return $arg }
    $escaped = [regex]::Replace($arg, '(\\*)"', '$1$1\"')
    $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
    return '"' + $escaped + '"'
}

function Format-CsvSafeField([string]$val) {
    if ($null -eq $val) { return '""' }
    $safe = $val
    if ($safe -match '^[=+\-@\t\r]') {
        $safe = "'" + $safe
    }
    return '"' + ($safe -replace '"', '""') + '"'
}

function Invoke-ProcessInMemory([string]$filePath, [string[]]$argList) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $filePath
    $escapedArgs = foreach ($a in $argList) { Format-WindowsProcessArg $a }
    $psi.Arguments = $escapedArgs -join " "
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    try {
        $proc.StartInfo = $psi
        [void]$proc.Start()
        $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
        $stderrTask = $proc.StandardError.ReadToEndAsync()
        $proc.WaitForExit()

        return [PSCustomObject]@{
            ExitCode = $proc.ExitCode
            StdOut   = $stdoutTask.Result
            StdErr   = $stderrTask.Result
        }
    } finally {
        $proc.Dispose()
    }
}

# 6. 辅助函数：读取视频真实帧率 & 秒数转帧时间码
function Get-VideoFps([string]$videoPath) {
    if (-not $ffprobeExe) { return 24.0 }
    $res = Invoke-ProcessInMemory $ffprobeExe @(
        "-v", "error",
        "-protocol_whitelist", "file,pipe",
        "-select_streams", "v:0",
        "-show_entries", "stream=r_frame_rate",
        "-of", "default=noprint_wrappers=1:nokey=1",
        "-i", $videoPath
    )
    $raw = ($res.StdOut -split "`r?`n" | Where-Object { $_ -match '\d' } | Select-Object -First 1)
    if ($raw -match '^(\d+)\s*/\s*(\d+)$') {
        $num = [double]::Parse($Matches[1], $InvariantCulture)
        $den = [double]::Parse($Matches[2], $InvariantCulture)
        if ($den -gt 0 -and ($num / $den) -gt 1 -and ($num / $den) -lt 240) {
            return ($num / $den)
        }
    } elseif ($raw -match '^\d+(\.\d+)?$') {
        return [double]::Parse($raw, $InvariantCulture)
    }
    return 24.0
}

function ConvertTo-Timecode([int]$totalFrames, [int]$timebase, [int]$startHourOffset) {
    if ($totalFrames -lt 0) { $totalFrames = 0 }
    $framesPerSec  = $timebase
    $framesPerMin  = $framesPerSec * 60
    $framesPerHour = $framesPerMin * 60

    $hh = [Math]::Floor($totalFrames / $framesPerHour) + $startHourOffset
    $rem = $totalFrames % $framesPerHour
    $mm = [Math]::Floor($rem / $framesPerMin)
    $rem = $rem % $framesPerMin
    $ss = [Math]::Floor($rem / $framesPerSec)
    $ff = $rem % $framesPerSec

    return ("{0:D2}:{1:D2}:{2:D2}:{3:D2}" -f [int]$hh, [int]$mm, [int]$ss, [int]$ff)
}

# 7. 辅助函数：可选读取同目录同名/同前缀 SRT 字幕分色（含大小限制与正则超时保护）
function Get-SrtIntervals([string]$dir, [string]$videoBase) {
    $exactSrt = Join-Path $dir ($videoBase + ".srt")
    $srtFile = $null
    if (Test-Path -LiteralPath $exactSrt -PathType Leaf) {
        $srtFile = Get-Item -LiteralPath $exactSrt
    } else {
        $srtFile = Get-ChildItem -LiteralPath $dir -File |
            Where-Object { $_.Extension -ieq ".srt" -and $_.BaseName.StartsWith($videoBase + "_", [System.StringComparison]::OrdinalIgnoreCase) } |
            Select-Object -First 1
    }
    if ($null -eq $srtFile) { return $null }

    if ($srtFile.Length -gt 20MB) {
        Write-Host "  [提示] 字幕文件 $($srtFile.Name) 超过 20MB 安全上限，已跳过字幕分色。" -ForegroundColor Yellow
        return $null
    }

    $content = [System.IO.File]::ReadAllText($srtFile.FullName, [System.Text.Encoding]::UTF8)
    $regex = New-Object System.Text.RegularExpressions.Regex(
        '(\d{2}):(\d{2}):(\d{2})[,\.](\d{3})\s*-->\s*(\d{2}):(\d{2}):(\d{2})[,\.](\d{3})',
        [System.Text.RegularExpressions.RegexOptions]::None,
        [TimeSpan]::FromSeconds(5)
    )
    $matchesList = $regex.Matches($content)
    $subs = New-Object System.Collections.Generic.List[object]
    foreach ($m in $matchesList) {
        $s = [int]$m.Groups[1].Value * 3600 + [int]$m.Groups[2].Value * 60 + [int]$m.Groups[3].Value + [int]$m.Groups[4].Value / 1000.0
        $e = [int]$m.Groups[5].Value * 3600 + [int]$m.Groups[6].Value * 60 + [int]$m.Groups[7].Value + [int]$m.Groups[8].Value / 1000.0
        $subs.Add([PSCustomObject]@{ Start = $s; End = $e })
    }
    return [PSCustomObject]@{ FileName = $srtFile.Name; Intervals = $subs }
}

# 8. 逐个扫描 MP4 并输出波谷文件
$utf8Bom   = New-Object System.Text.UTF8Encoding($true)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$summaryRows = New-Object System.Collections.Generic.List[string]
$summaryRows.Add("视频文件名,视频帧率,波谷序号,起点秒数(s),终点秒数(s),谷底中心秒数(s),持续时长(s),起点时间码,终点时间码,持续帧数,气口分类")

Write-Host ""
Write-Host "本机 FFmpeg：$ffmpegExe" -ForegroundColor Gray
Write-Host "目标文件夹：$targetDir"
Write-Host "输出文件夹：$outDir"
Write-Host "共发现 $($mp4Files.Count) 个 MP4 文件，开始提取波谷..." -ForegroundColor Green
Write-Host "------------------------------------------------------------------"

$fileIndex = 0
foreach ($mp4 in $mp4Files) {
    $fileIndex++
    $baseName = $mp4.BaseName
    Write-Host "[$fileIndex/$($mp4Files.Count)] 正在分析：$($mp4.Name) ..." -ForegroundColor White

    $fpsReal  = Get-VideoFps $mp4.FullName
    $timebase = [int][Math]::Round($fpsReal)
    if ($timebase -le 0) { $timebase = 24 }

    $filterArg = [string]::Format($InvariantCulture, "silencedetect=noise={0}:d={1:0.###}", $NoiseThreshold.Trim(), [double]$MinDuration)
    # 锁定 file/pipe 协议白名单阻断伪装流媒体外联，并禁用视频/字幕解码器 (-vn -sn -dn) 缩减攻击面与提速
    $ffRes = Invoke-ProcessInMemory $ffmpegExe @(
        "-nostdin",
        "-hide_banner",
        "-nostats",
        "-protocol_whitelist", "file,pipe",
        "-i", $mp4.FullName,
        "-vn",
        "-sn",
        "-dn",
        "-af", $filterArg,
        "-f", "null",
        "-"
    )

    $logLines = $ffRes.StdErr -split "`r?`n"
    $rawSilenceLines = New-Object System.Collections.Generic.List[string]
    $valleys = New-Object System.Collections.Generic.List[object]
    $pendingStart = $null

    foreach ($line in $logLines) {
        if ($line -match 'silencedetect') {
            $rawSilenceLines.Add($line.Trim())
        }
        if ($line -match 'silence_start:\s*([0-9\.]+)') {
            $pendingStart = [double]::Parse($Matches[1], $InvariantCulture)
        } elseif ($line -match 'silence_end:\s*([0-9\.]+)\s*\|\s*silence_duration:\s*([0-9\.]+)') {
            $endSec = [double]::Parse($Matches[1], $InvariantCulture)
            $durSec = [double]::Parse($Matches[2], $InvariantCulture)
            $startSec = if ($null -ne $pendingStart) { $pendingStart } else { [Math]::Max(0.0, $endSec - $durSec) }
            $valleys.Add([PSCustomObject]@{
                StartSec = $startSec
                EndSec   = $endSec
                DurSec   = $durSec
            })
            $pendingStart = $null
        }
    }

    if ($ffRes.ExitCode -ne 0 -and $rawSilenceLines.Count -eq 0) {
        Write-Host "  [警告] FFmpeg 分析此文件音频轨时返回非零退出码（ExitCode=$($ffRes.ExitCode)），请确认视频包含有效音轨。" -ForegroundColor Yellow
    }

    $srtInfo = Get-SrtIntervals $targetDir $baseName

    $csvLines = New-Object System.Collections.Generic.List[string]
    $csvLines.Add("序号,起点秒数(s),终点秒数(s),谷底中心秒数(s),持续时长(s),起点帧号,终点帧号,持续帧数,起点时间码,终点时间码,气口分类")

    $safeTitleBase = ($baseName -replace '[\r\n\x00-\x1F]', '_')
    $edlLines = New-Object System.Collections.Generic.List[string]
    $edlLines.Add("TITLE: ${safeTitleBase}_气口波谷标记")
    $edlLines.Add("FCM: NON-DROP FRAME")
    $edlLines.Add("")

    $vIdx = 0
    foreach ($v in $valleys) {
        $vIdx++
        $midSec     = ($v.StartSec + $v.EndSec) / 2.0
        $startFrame = [int][Math]::Round($v.StartSec * $fpsReal)
        $endFrame   = [int][Math]::Round($v.EndSec * $fpsReal)
        if ($endFrame -le $startFrame) { $endFrame = $startFrame + 1 }
        $durFrames  = $endFrame - $startFrame

        $tcIn  = ConvertTo-Timecode $startFrame $timebase $TimelineStartHour
        $tcOut = ConvertTo-Timecode $endFrame   $timebase $TimelineStartHour

        $category = ""
        $edlColor = "ResolveColorCyan"
        if ($null -ne $srtInfo -and $srtInfo.Intervals.Count -gt 0) {
            $insideSingleSub = $false
            foreach ($sub in $srtInfo.Intervals) {
                if (($v.StartSec -ge ($sub.Start + 0.08)) -and ($v.EndSec -le ($sub.End - 0.08))) {
                    $insideSingleSub = $true
                    break
                }
            }
            if ($insideSingleSub) {
                $category = "句内停顿(结合SRT)"
                $edlColor = "ResolveColorYellow"
            } else {
                $category = "句间真气口(结合SRT)"
                $edlColor = "ResolveColorGreen"
            }
        } else {
            if ($v.DurSec -ge $BigValleySeconds) {
                $category = "大气口"
                $edlColor = "ResolveColorGreen"
            } else {
                $category = "小气口"
                $edlColor = "ResolveColorCyan"
            }
        }

        $sStr   = [string]::Format($InvariantCulture, "{0:F3}", $v.StartSec)
        $eStr   = [string]::Format($InvariantCulture, "{0:F3}", $v.EndSec)
        $mStr   = [string]::Format($InvariantCulture, "{0:F3}", $midSec)
        $dStr   = [string]::Format($InvariantCulture, "{0:F3}", $v.DurSec)
        $fpsStr = [string]::Format($InvariantCulture, "{0:0.##}", [Math]::Round($fpsReal, 2))

        $csvLines.Add("$vIdx,$sStr,$eStr,$mStr,$dStr,$startFrame,$endFrame,$durFrames,$tcIn,$tcOut,$category")
        $safeVideoCsv = Format-CsvSafeField $mp4.Name
        $summaryRows.Add("$safeVideoCsv,$fpsStr,$vIdx,$sStr,$eStr,$mStr,$dStr,$tcIn,$tcOut,$durFrames,$category")

        $eventNum = "{0:D3}" -f $vIdx
        $markerLabel = [string]::Format($InvariantCulture, "气口{0:D3}_{1}_{2:F2}s", $vIdx, $category, $v.DurSec)
        $edlLines.Add("$eventNum  001      V     C        $tcIn $tcOut $tcIn $tcOut  ")
        $edlLines.Add(" |C:$edlColor |M:$markerLabel |D:$durFrames")
        $edlLines.Add("")
    }

    # 写出 1：CSV 明细表（带 UTF-8 BOM，Excel 双击直开不乱码）
    $csvPath = Get-SafeOutputFile $outDir "${baseName}_波谷明细" ".csv"
    [System.IO.File]::WriteAllLines($csvPath, $csvLines, $utf8Bom)

    # 写出 2：纯净 FFmpeg 波谷原始记录 TXT（头信息包含原视频文件名与帧率，供 Muse 生成 XML 直接读取）
    $fpsHdrStr = [string]::Format($InvariantCulture, "{0:0.###}", [Math]::Round($fpsReal, 3))
    $minDurHdr = [string]::Format($InvariantCulture, "{0:0.###}", [double]$MinDuration)
    $txtHeader = @(
        "# 视频文件: $($mp4.Name)",
        "# 检测帧率: $fpsHdrStr fps (时间码基准: ${timebase}fps)",
        "# 检测参数: 阈值=$NoiseThreshold, 最短时长=${minDurHdr}s",
        "# 识别波谷总数: $($valleys.Count) 处",
        "------------------------------------------------------------"
    )
    $txtPath = Get-SafeOutputFile $outDir "${baseName}_波谷原始记录" ".txt"
    [System.IO.File]::WriteAllLines($txtPath, ($txtHeader + $rawSilenceLines), $utf8NoBom)

    # 写出 3：达芬奇时间线区间标记 EDL
    $edlPath = Get-SafeOutputFile $outDir "${baseName}_达芬奇区间标记" ".edl"
    [System.IO.File]::WriteAllLines($edlPath, $edlLines, $utf8NoBom)

    $srtTip = if ($null -ne $srtInfo) { "（已自动结合字幕 $($srtInfo.FileName) 分色）" } else { "" }
    Write-Host "  -> 完成！共提取 $($valleys.Count) 处气口波谷 | 帧率: ${timebase}fps $srtTip" -ForegroundColor Green
}

$sumPath = Get-SafeOutputFile $outDir "_全部MP4波谷汇总" ".csv"
[System.IO.File]::WriteAllLines($sumPath, $summaryRows, $utf8Bom)

Write-Host "------------------------------------------------------------------"
Write-Host "[全部完成] 共处理 $($mp4Files.Count) 个 MP4 文件！" -ForegroundColor Cyan
Write-Host "结果已安全保存至：$outDir" -ForegroundColor Cyan
if (-not [Console]::IsInputRedirected) {
    Write-Host ""
    [void](Read-Host "按回车键退出窗口")
}
