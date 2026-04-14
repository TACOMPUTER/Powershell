param(
    [int]$PSWidth = 80,
    [int]$PSHeight = 50,
    [int]$PosX = 0,
    [int]$PosY = 0,
    [bool]$SkipAdminCheck = $false
)



# ===== INIT WINAPI =====
if (-not ("WinAPI" -as [type])) {

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WinAPI {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

}

# lấy handle console
$consoleHandle = [WinAPI]::GetConsoleWindow()



# 🚩 <<<--- XÁC ĐỊNH SCRIPT PATH --->>>

$MainScript = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }

# 🏁 <<<--- END --->>>


# 🚩 <<<--- CHECK ADMIN --->>>

$IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $SkipAdminCheck -and -not $IsAdmin) {

    Write-Host "⚠️ Đang nâng quyền Administrator..." -ForegroundColor Yellow

    Start-Process powershell `
        -ArgumentList "-ExecutionPolicy Bypass -File `"$MainScript`"" `
        -Verb RunAs

    exit
}

# 🏁 <<<--- END --->>>


# 🚩 <<<--- CHỈ CHO 1 SCRIPT CHẠY --->>>

$currentPID = $PID
$scriptName = [System.IO.Path]::GetFileName($MainScript)

Get-CimInstance Win32_Process | Where-Object {
    $_.ProcessId -ne $currentPID -and
    $_.CommandLine -match [regex]::Escape($scriptName)
} | ForEach-Object {
    try {
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    } catch {}
}

# 🏁 <<<--- END --->>>


# 🚩 <<<--- SET WINDOW TITLE --->>>

$fileName = Split-Path $MainScript -Leaf
$folderPath = Split-Path $MainScript -Parent
$adminText = if ($IsAdmin) { "as Admin" } else { "as User" }

$host.UI.RawUI.WindowTitle = "Running '$fileName' $adminText <<< $folderPath"

# 🏁 <<<--- END --->>>


# 🚩 <<<--- RESIZE CONSOLE --->>>

$maxWidth  = $host.UI.RawUI.MaxWindowSize.Width
$maxHeight = $host.UI.RawUI.MaxWindowSize.Height

$PSWidth  = [Math]::Min($PSWidth,  $maxWidth)
$PSHeight = [Math]::Min($PSHeight, $maxHeight)

[Console]::BufferWidth  = [Math]::Max($PSWidth,  [Console]::BufferWidth)
[Console]::BufferHeight = [Math]::Max($PSHeight, [Console]::BufferHeight)

[Console]::WindowWidth  = $PSWidth
[Console]::WindowHeight = $PSHeight

# 🏁 <<<--- END --->>>


# 🚩 <<<--- MOVE WINDOW --->>>

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WinMove {
    [DllImport("user32.dll")]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int W, int H, bool repaint);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
"@

$handle = (Get-Process -Id $PID).MainWindowHandle

$rect = New-Object WinMove+RECT
[WinMove]::GetWindowRect($handle, [ref]$rect)

$widthPx  = $rect.Right - $rect.Left
$heightPx = $rect.Bottom - $rect.Top

[WinMove]::MoveWindow($handle, $PosX, $PosY, $widthPx, $heightPx, $true) | Out-Null

# 🏁 <<<--- END --->>>



# Tài khoản mới lần đầu chạy (ko cần quyền Admin) sẽ thông báo 'Execution Policy Change'. Tắt thông báo chạy lại là được.

# Lấy script gốc (fix callstack)
$callStack = Get-PSCallStack

if ($callStack.Count -gt 1 -and $callStack[1].ScriptName) {
    $MainScript = $callStack[-1].ScriptName
} else {
    $MainScript = $PSCommandPath
}

# Kiểm tra quyền Admin
$IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $SkipAdminCheck -and -not $IsAdmin) {

    Write-Host "⚠️ Đang nâng quyền Administrator..." -ForegroundColor Yellow

    Start-Process powershell `
        -ArgumentList "-ExecutionPolicy Bypass -File `"$MainScript`"" `
        -Verb RunAs

    exit
}



# 🚩 <<<--- CHỈ CHO 1 SCRIPT CHẠY --->>>

$currentPID = $PID
$scriptName = [System.IO.Path]::GetFileName($PSCommandPath)

Get-Process powershell | Where-Object {
    $_.Id -ne $currentPID -and
    $_.Path -ne $null
} | ForEach-Object {

    try {
        # Lấy command line an toàn hơn
        $cmd = (Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" -ErrorAction SilentlyContinue).CommandLine

        if ($cmd -and $cmd -match [regex]::Escape($scriptName)) {
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        # bỏ qua lỗi SID mapping
    }
}

# 🏁 <<<--- END --->>>

# Set font Consolas
$FontInitFlag = "HKCU:\Software\MyITTool"

if (-not (Test-Path $FontInitFlag)) {
    New-Item -Path $FontInitFlag -Force | Out-Null
}

$FontSet = (Get-ItemProperty -Path $FontInitFlag -Name "FontSet" -ErrorAction SilentlyContinue).FontSet

if (-not $FontSet) {
    try {
        # Whitelist font (Admin)
        New-ItemProperty `
          -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Console\TrueTypeFont" `
          -Name "000" -Value "Consolas" -PropertyType String -Force `
          -ErrorAction SilentlyContinue | Out-Null

        # Set đúng key PowerShell
        $psKey = "HKCU:\Console\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe"

        if (-not (Test-Path $psKey)) {
            New-Item -Path $psKey -Force | Out-Null
        }

        Set-ItemProperty -Path $psKey -Name FaceName -Value "Consolas"
        Set-ItemProperty -Path $psKey -Name FontSize -Value 0x00100000

        # Đánh dấu đã set
        Set-ItemProperty -Path $FontInitFlag -Name "FontSet" -Value 1 -Force

        # Mở lại 1 lần duy nhất
        Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
        exit
    }
    catch {}
}

# 🚩 <<<--- KHU VỰC KHAI BÁO BIẾN SỬ DỤNG TOÀN BỘ SCRIPT ---->>>

$ITScriptRoot = $PSScriptRoot	# ...\Software\OS Tools\cmd-Powershell

$LibScript = Join-Path $ITScriptRoot "IT\Library"

$IT113Script = Join-Path $ITScriptRoot "IT\IT-113"

$IT115Script = Join-Path $ITScriptRoot "IT\IT-115"

	# thư mục software: thư mục gốc chứa file IT.ps1 (có thể là 'local' hoặc 'unc: \\server\share')
$SourceSW = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent	#...\Software

	# thư mục software2
		# 1. nếu nằm trong onedrive\tacomputer\software
if ($SourceSW -match "OneDrive\\TACOMPUTER\\Software$") {
    $drive = ([System.IO.Path]::GetPathRoot($SourceSW))
    $SourceSW2 = Join-Path $drive "Software2"
}
		# 2. nếu là driveletter:\software
elseif ($SourceSW -match "^[A-Z]:\\Software$") {
    $SourceSW2 = $SourceSW + "2"
}
		# 3. nếu là \\network\software
elseif ($SourceSW -match "^\\\\.*\\Software$") {
    $SourceSW2 = $SourceSW + "2"
}

	# thư mục dùng để lưu shortcut và file cấu hình
$SystemDriveSW = "C:\SW"

	# file thực thi chính dùng để chạy tool
$ExePath = Join-Path $ITScriptRoot "IT.exe"

	# shortcut đặt tại c:\sw
$SystemDriveSWlnk = "$SystemDriveSW\IT.exe.lnk"

	# lấy computername\username của user đang đăng nhập
$currentUser = "$env:COMPUTERNAME\$env:USERNAME"

	# đường dẫn start menu (dùng biến môi trường để tránh lỗi profile)
$StartMenuProgramsPath = [Environment]::GetFolderPath("Programs")
$StartMenuShortPath = $StartMenuProgramsPath.Substring(
    $StartMenuProgramsPath.IndexOf("\Start Menu")
)

	# shortcut đặt trong start menu
$StartMenuProgramslnk = "$StartMenuProgramsPath\IT.exe.lnk"

	# expand đường dẫn thật từ %appdata%
$ExpandedStartMenuPath = [Environment]::ExpandEnvironmentVariables($StartMenuProgramsPath)

	# expand đường dẫn shortcut start menu
$ExpandedStartMenuLnk = [Environment]::ExpandEnvironmentVariables($StartMenuProgramslnk)

	# file export chứa các biến path dùng cho script khác
$exportvariablePath = "$SystemDriveSW\variable_IT.ps1"

# 🏁 <<<--- KẾT THÚC KHU VỰC KHAI BÁO BIẾN ---->>>



# 🚩 <<<--- XUẤT CÁC BIẾN RA FILE 'VARIABLE_IT.PS1' --->>>

	# xóa file cũ nếu có
if (Test-Path $exportvariablePath) {
    Remove-Item $exportvariablePath -Force
    Write-Host "Đã xóa file cũ: $exportvariablePath`n" -ForegroundColor Yellow
}
	# hàm kiểm tra chuỗi giống đường dẫn
function Is-PathLike($str) {
    return ($str -is [string]) -and (
        $str -match '^[a-zA-Z]:\\' -or
        $str -match '^\\\\' -or
        $str -match '\\.+\\' -or
        $str -match '\\$'
    )
}
	# danh sách biến hệ thống cần loại trừ
$excludedNames = @(
    'HOME', 'PSHOME', 'PROFILE', 'PID', 'ExecutionContext', 'Host', 'ShellId',
    'env', 'args', 'Error', 'MyInvocation', 'PSBoundParameters', 'PSCommandPath',
    'PSCulture', 'PSEdition', 'PSScriptRoot', 'PSUICulture', 'PSVersionTable',
    'input', 'output', 'null'
)
	# lấy các biến hợp lệ
$vars = Get-Variable | Where-Object {
    ($_.Value -is [string]) -and
    (Is-PathLike $_.Value) -and
    (-not ($excludedNames -contains $_.Name)) -and
    ($_.Options -notmatch 'ReadOnly|Constant|AllScope')
}
$lines = @()
foreach ($var in $vars) {
    $name = $var.Name
    $value = '"' + $var.Value.Replace('"', '`"') + '"'
    $lines += "`$$name = $value"
}
	# ghi ra file mới
# Tạo thư mục nếu chưa có
if (-not (Test-Path $SystemDriveSW)) {
    New-Item -Path $SystemDriveSW -ItemType Directory -Force | Out-Null
}

# Sau đó ghi file
$lines | Set-Content $exportvariablePath

# 🏁 <<<--- XUẤT CÁC BIẾN RA FILE 'VARIABLE_IT.PS1' --->>>



# 🚩 <<<--- TẠO SHORTCUT IT.exe --->>>

$WshShell = New-Object -ComObject WScript.Shell

	# đảm bảo C:\SW tồn tại
if (-not (Test-Path $SystemDriveSW)) {
    New-Item -ItemType Directory -Path $SystemDriveSW -Force | Out-Null
}

	# tạo shortcut C:\SW
$ShortcutSystemDrive = $WshShell.CreateShortcut($SystemDriveSWlnk)
$ShortcutSystemDrive.TargetPath = $ExePath
$ShortcutSystemDrive.WorkingDirectory = $ITScriptRoot
$ShortcutSystemDrive.Description = "Shortcut to IT.exe"
$ShortcutSystemDrive.IconLocation = $ExePath
$ShortcutSystemDrive.Save()

	# đảm bảo Start Menu tồn tại
if (-not (Test-Path $ExpandedStartMenuPath)) {
    New-Item -ItemType Directory -Path $ExpandedStartMenuPath -Force | Out-Null
}

	# tạo shortcut Start Menu
if (Test-Path $ExpandedStartMenuLnk) {
    Remove-Item $ExpandedStartMenuLnk -Force
}

$ShortcutStartMenu = $WshShell.CreateShortcut($ExpandedStartMenuLnk)
$ShortcutStartMenu.TargetPath = $ExePath
$ShortcutStartMenu.WorkingDirectory = $ITScriptRoot
$ShortcutStartMenu.Description = "Shortcut to IT.exe"
$ShortcutStartMenu.IconLocation = $ExePath
$ShortcutStartMenu.Save()

	# hiển thị đường dẫn rút gọn
$StartMenuIndex = $ExpandedStartMenuLnk.IndexOf("Start Menu")
$DesiredOutput = "...\" + $ExpandedStartMenuLnk.Substring($StartMenuIndex)

# 🏁 <<<--- TẠO SHORTCUT IT.exe --->>>



# 🚩 <<<--- LẤY DANH SÁCH "WINDOWS SECURITY\EXCLUSION" ĐANG CÓ --->>>

# =====================================================
# LOG
# =====================================================

$Logs = "C:\Temp\DefenderExclusion.txt"
$logDir = Split-Path $Logs

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

Start-Transcript $Logs -Append -Force | Out-Null


# =====================================================
# HELPER
# =====================================================

function Normalize {
    param($p)
    return ($p.TrimEnd('\')).ToLower()
}

function Get-RealPath {
    param($p)

    try {
        return (Get-Item -LiteralPath $p -ErrorAction Stop).FullName
    }
    catch {
        return $p
    }
}


# =====================================================
# WAIT DEFENDER
# =====================================================

function Wait-Defender {

    $svc = Get-Service WinDefend -ErrorAction SilentlyContinue
    if (-not $svc) { return $false }

    for ($i = 0; $i -lt 20; $i++) {

        if ($svc.Status -eq 'Running') {
            try {
                Get-MpPreference -ErrorAction Stop | Out-Null
                return $true
            }
            catch {}
        }

        Start-Sleep 1
        try { $svc.Refresh() } catch {}
    }

    return $false
}

if (-not (Wait-Defender)) {
    Write-Host "❌ Defender chưa sẵn sàng" -ForegroundColor Red
    exit
}

Write-Host "✅ Defender OK" -ForegroundColor Cyan


# =====================================================
# DESIRED CONFIG
# =====================================================

$desiredPaths = @(
    "\\IT\Software",
    "\\IT\Software2",
    "\\IT-E580\Software",
    "\\IT-E580\Software2",
    "C:\SW"
)

$desiredProcess = @(
    "SppExtComObjHook.dll"
)

$desiredExt = @()


# =====================================================
# SCAN DRIVE
# =====================================================

Get-CimInstance Win32_LogicalDisk |
Where-Object { $_.DriveType -in 2,3 } |
ForEach-Object {

    $root = $_.DeviceID

    $scanList = @(
        "$root\Software",
        "$root\Software2",
        "$root\OneDrive\TACOMPUTER\Software"
    )

    foreach ($path in $scanList) {
        if (Test-Path -LiteralPath $path) {
            $desiredPaths += $path
        }
    }
}


# =====================================================
# CLEAN DATA
# =====================================================

$desiredPathsRaw = $desiredPaths |
Where-Object { $_ -and (Test-Path -LiteralPath $_) } |
Sort-Object -Unique

# normalize chỉ để so sánh
$desiredPathsNorm = $desiredPathsRaw | ForEach-Object { Normalize $_ }

$currentRaw = @((Get-MpPreference).ExclusionPath) | Where-Object { $_ }
$currentNorm = $currentRaw | ForEach-Object { Normalize $_ }


# =====================================================
# ENSURE REQUIRED PATH (KHÔNG REMOVE)
# =====================================================

$requiredPaths = @(
    "\\IT\Software",
    "\\IT\Software2",
    "\\IT-E580\Software",
    "\\IT-E580\Software2",
    "C:\SW"
)

$currentRaw = @((Get-MpPreference).ExclusionPath) | Where-Object { $_ }
$currentNorm = $currentRaw | ForEach-Object { Normalize $_ }

$toAdd = $requiredPaths | Where-Object {
    (Normalize $_) -notin $currentNorm
}

function Add-DefenderPathSafe {
    param($path)

    $maxRetry = 3

    for ($i = 1; $i -le $maxRetry; $i++) {

        try {
            Add-MpPreference -ExclusionPath $path -ErrorAction Stop
            Write-Host "ADD PATH : $path" -ForegroundColor Yellow
            return
        }
        catch {
            if ($i -lt $maxRetry) {
                Start-Sleep -Milliseconds 500
            } else {
                Write-Host "FAIL ADD PATH : $path" -ForegroundColor Red
                Write-Host $_.Exception.Message -ForegroundColor DarkRed
            }
        }
    }
}

foreach ($path in $toAdd) {
    try {
        Add-DefenderPathSafe $path
        Write-Host "ADD PATH : $path" -ForegroundColor Yellow
    } catch {
        Write-Host "FAIL ADD PATH : $path" -ForegroundColor Red
    }
}

if ($toAdd.Count -eq 0) {
    Write-Host "✔ PATH đủ, không cần thêm" -ForegroundColor Green
}

# =====================================================
# SCAN DRIVE
# =====================================================

$validDrives = Get-CimInstance Win32_LogicalDisk |
    Where-Object { $_.DriveType -in 2,3 }

$existingPaths = @()

foreach ($drive in $validDrives) {

    $root = $drive.DeviceID

    $softwareRoot = Join-Path $root "Software"
    if (Test-Path $softwareRoot) {
        $existingPaths += $softwareRoot
    }

    $software2Root = Join-Path $root "Software2"
    if (Test-Path $software2Root) {
        $existingPaths += $software2Root
    }

    $softwareOneDrive = Join-Path $root "OneDrive\TACOMPUTER\Software"
    if (Test-Path $softwareOneDrive) {
        $existingPaths += $softwareOneDrive
    }
}

# =====================================================
# ADD DYNAMIC PATH (KHÔNG TRÙNG)
# =====================================================

$currentPaths = @((Get-MpPreference).ExclusionPath) | Where-Object { $_ }
$currentNorm  = $currentPaths | ForEach-Object { Normalize $_ }

$toAdd = $existingPaths | Where-Object {
    (Normalize $_) -notin $currentNorm
}

foreach ($p in $toAdd) {
    try {
        Add-MpPreference -ExclusionPath $p
        Write-Host "ADD DYNAMIC PATH : $p" -ForegroundColor Yellow
    } catch {
        Write-Host "FAIL ADD PATH : $p" -ForegroundColor Red
    }
}

if ($toAdd.Count -eq 0) {
    Write-Host "✔ Không có path mới cần thêm" -ForegroundColor Green
}

# =====================================================
# CLEAN INVALID PATH (KHÔNG ĐỤNG REQUIRED PATH)
# =====================================================

$tempJson = "C:\Temp\DefenderExclusions.json"

$data = @{
    Path = @(
        "\\IT\Software",
        "\\IT\Software2",
        "\\IT-E580\Software",
        "\\IT-E580\Software2",
        "C:\SW"
    )
    Process   = @()
    Extension = @()
}

# đảm bảo folder tồn tại
$dir = Split-Path $tempJson
if (-not (Test-Path $dir)) {
    New-Item -ItemType Directory -Path $dir | Out-Null
}

$data | ConvertTo-Json -Depth 3 | Set-Content $tempJson -Encoding UTF8



# =====================================================
# SYNC PROCESS - SppExtComObjHook.dll (FIXED PATH)
# =====================================================

	# Tạo JSON tạm
$proc = "SppExtComObjHook.dll"
$fullPath = "$env:WINDIR\System32\$proc"

$currentProcess = @((Get-MpPreference).ExclusionProcess) | Where-Object { $_ }

if (Test-Path -LiteralPath $fullPath) {

    if ($proc -notin $currentProcess) {
        try {
            Add-MpPreference -ExclusionProcess $proc
            Write-Host "ADD PROCESS : $proc" -ForegroundColor Yellow
        } catch {}
    } else {
        Write-Host "OK PROCESS  : $proc" -ForegroundColor Green
    }

}
else {

    if ($proc -in $currentProcess) {
        try {
            Remove-MpPreference -ExclusionProcess $proc
            Write-Host "REMOVE PROCESS : $proc (not found)" -ForegroundColor Red
        } catch {}
    } else {
        Write-Host "SKIP PROCESS : $proc (not found)" -ForegroundColor DarkGray
    }

}

	# RESTORE từ JSON
$data = Get-Content $tempJson | ConvertFrom-Json

foreach ($p in $data.Path) {

    if ($p -like "\\*") {
        # UNC chạy nền để không khựng
        Start-Job {
            param($x)
            Add-MpPreference -ExclusionPath $x -ErrorAction SilentlyContinue
        } -ArgumentList $p | Out-Null

        Write-Host "ADD UNC (bg): $p" -ForegroundColor DarkYellow
    }
    else {
        Add-MpPreference -ExclusionPath $p -ErrorAction SilentlyContinue
        Write-Host "ADD PATH : $p" -ForegroundColor Yellow
    }
}

# =====================================================
# SYNC EXTENSION (ONLY MANAGE desiredExt)
# =====================================================

function Test-ExtensionExists {
    param($ext)

    foreach ($dir in @(
        "C:\SW",
        "C:\Software"
    )) {
        try {
            if (Get-ChildItem -Path $dir -Filter "*.$ext" -File -Recurse -Depth 2 -ErrorAction SilentlyContinue) {
                return $true
            }
        } catch {}
    }

    return $false
}

$currentExt = @((Get-MpPreference).ExclusionExtension) | Where-Object { $_ }

foreach ($ext in $desiredExt) {

    $exists = Test-ExtensionExists $ext

    if ($exists) {
        if ($ext -notin $currentExt) {
            Add-MpPreference -ExclusionExtension $ext
            Write-Host "ADD EXT : $ext" -ForegroundColor Yellow
        } else {
            Write-Host "OK EXT  : $ext" -ForegroundColor Green
        }
    }
    else {
        if ($ext -in $currentExt) {
            Remove-MpPreference -ExclusionExtension $ext
            Write-Host "REMOVE EXT : $ext (not found)" -ForegroundColor Red
        } else {
            Write-Host "SKIP EXT : $ext (not found)" -ForegroundColor DarkGray
        }
    }
}

# =====================================================
# SHOW RESULT (FIX CASE CHUẨN)
# =====================================================

$pref = Get-MpPreference

Write-Host ""
Write-Host "===== PATH =====" -ForegroundColor Cyan
foreach ($path in $pref.ExclusionPath) {
    Write-Host (Get-RealPath $path) -ForegroundColor Yellow
}

Write-Host ""
Write-Host "===== PROCESS =====" -ForegroundColor Cyan
foreach ($proc in $pref.ExclusionProcess) {
    Write-Host $proc -ForegroundColor Yellow
}

Write-Host ""
Write-Host "===== EXTENSION =====" -ForegroundColor Cyan
foreach ($ext in $pref.ExclusionExtension) {
    Write-Host $ext -ForegroundColor Yellow
}

# Khởi chạy ứng dụng Windows Security Health
try {
    Start-Process -FilePath "$env:WINDIR\system32\SecurityHealthSystray.exe"
    Write-Host "Đã khởi chạy SecurityHealthSystray.exe." -ForegroundColor Yellow
    Write-Host ""
} catch {
    Write-Warning "Không thể khởi chạy SecurityHealthSystray.exe. Lỗi: $($_.Exception.Message)"
}

# =====================================================
# DONE
# =====================================================

Write-Host ""
Write-Host "DONE" -ForegroundColor Green

Stop-Transcript | Out-Null
Start-Sleep 2

# 🏁 <<<--- LẤY DANH SÁCH "WINDOWS SECURITY\EXCLUSION" ĐANG CÓ --->>>



function Run-IT-xxx {
    param([string]$ScriptPath)

    # minimize
    [WinAPI]::ShowWindow($consoleHandle, 6)

    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`"" -Wait

    # restore
    [WinAPI]::ShowWindow($consoleHandle, 9)
}



# 🚩 <<<--- CHỌN HỖ TRỢ 113 114 115 --->>>

function Show-Menu-IT {
	Clear-Host

    # 🚩 <<<--- LOGO IT --->>>

	Write-Host ("+" * $Host.UI.RawUI.WindowSize.Width) -ForegroundColor DarkGray
	$text = " IT support, Scripted by TACOMPUTER & GPT, 0933.848.990 "
	$width = $Host.UI.RawUI.WindowSize.Width
	$pad = [Math]::Max(0, $width - $text.Length)
	$left  = [Math]::Floor($pad / 2)
	$right = $pad - $left
	Write-Host ("+" * $left) -ForegroundColor DarkGray -NoNewline
	Write-Host $text -NoNewline
	Write-Host ("+" * $right) -ForegroundColor DarkGray
	Write-Host ("+" * $Host.UI.RawUI.WindowSize.Width) -ForegroundColor DarkGray

	# 🏁 <<<--- LOGO IT --->>>



	# 🚩 <<<--- THÔNG TIN CƠ BẢN MÁY TÍNH --->>>

	$IsLaptop = (Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue) -ne $null

	if ($IsLaptop) {
		Write-Host "<<< Laptop Information >>>" -ForegroundColor Cyan
	}
	else {
		Write-Host "<<< PC Information >>>" -ForegroundColor Cyan
	}

	# ===== FORMAT CONFIG =====
	$global:LabelWidth = 18
	$global:SubLabelWidth = 16

	# ===== UI FUNCTIONS =====
	function Show-Line {
		param($label, $value)

		Write-Host ("{0,-$global:LabelWidth}   >>> " -f $label) -NoNewline -ForegroundColor Green
		Write-Host $value -ForegroundColor Yellow
	}

	function Show-SubLine {
		param($label, $value)

		Write-Host ("  {0,-$global:SubLabelWidth} → " -f $label) -NoNewline -ForegroundColor Blue
		Write-Host $value -ForegroundColor Blue
	}

	# ===== GET INFO =====
	$CS      = Get-CimInstance Win32_ComputerSystem
	$BB      = Get-CimInstance Win32_BaseBoard
	$CPU     = Get-CimInstance Win32_Processor
	$RAM     = @(Get-CimInstance Win32_PhysicalMemory)
	$arrays  = Get-CimInstance Win32_PhysicalMemoryArray
	$BIOS    = Get-CimInstance Win32_BIOS
	$GPU     = Get-CimInstance Win32_VideoController
	$OS      = Get-CimInstance Win32_OperatingSystem
	$RegOS   = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"

	# ===== SYSTEM =====
	Show-Line "Brand (OEM)" $CS.Manufacturer
	Show-Line "Mainboard" $BB.Manufacturer
	Show-Line "Product" $BB.Product
	Show-Line "Model" $CS.Model
	Show-Line "Serial" $BIOS.SerialNumber
	Show-Line "BIOS ver" $BIOS.SMBIOSBIOSVersion

	Write-Host ""

	# ===== CPU =====
	if ($CPU.Count -gt 1) {
		Show-Line "CPU" ""
		$i = 1
		foreach ($c in $CPU) {
			Show-SubLine ("CPU $i") $c.Name
			$i++
		}
	}
	else {
		Show-Line "CPU" $CPU.Name
	}

	# ===== RAM =====
	$totalRAM = "{0:N0}" -f (($RAM | Measure-Object Capacity -Sum).Sum / 1GB)
	Show-Line "RAM (Total)" "$totalRAM GB"

	$i = 1
	foreach ($ram in $RAM) {
		$size = "{0:N0}" -f ($ram.Capacity / 1GB)
		$slot = if ($ram.DeviceLocator) { $ram.DeviceLocator } else { "Slot $i" }

		Show-SubLine $slot "$size GB  $($ram.Speed) MHz"
		$i++
	}

	# SLOT INFO
	$TotalSlots = ($arrays | Measure-Object MemoryDevices -Sum).Sum
	$usedSlots = $RAM.Count
	$availableSlots = $TotalSlots - $usedSlots
	if ($availableSlots -lt 0) { $availableSlots = 0 }

	Show-SubLine "Available Slots" "$availableSlots/$TotalSlots"

	# ===== STORAGE =====
	$disks = Get-CimInstance Win32_DiskDrive

	$totalDisk = "{0:N0}" -f (($disks | Measure-Object Size -Sum).Sum / 1GB)
	Show-Line "Storage (Total)" "$totalDisk GB"

	$i = 1
	foreach ($disk in $disks) {
		$size = "{0:N0}" -f ($disk.Size / 1GB)
		Show-SubLine ("Disk $i") "$($disk.Model)  $size GB"
		$i++
	}

	# ===== GPU =====
	$vgaCount = 0
	$gpuCount = 0

	foreach ($g in $GPU) {
		if ($g.Name -match "NVIDIA|AMD|Radeon|GeForce|RTX|GTX") {
			$vgaCount++
		}
		else {
			$gpuCount++
		}
	}

	Show-Line "Graphics" ("{0} VGA / {1} GPU" -f $vgaCount, $gpuCount)

	foreach ($g in $GPU) {
		if ($g.Name -match "NVIDIA|AMD|Radeon|GeForce|RTX|GTX") {
			$label = "VGA"
		}
		else {
			$label = "GPU"
		}

		Show-SubLine $label $g.Name
	}

	Write-Host ""

	# ===== WINDOWS =====
	$version = $RegOS.DisplayVersion
	if (!$version) { $version = $RegOS.ReleaseId }

	$build = "$($RegOS.CurrentBuild).$($RegOS.UBR)"

	Show-Line "Windows" "$($RegOS.ProductName) | $version | $build"

	# ===== USER =====
	Write-Host ""
	Write-Host "Current User         >>> " -NoNewline
	Write-Host $currentUser -ForegroundColor Yellow

	Write-Host ("+" * $Host.UI.RawUI.WindowSize.Width) -ForegroundColor DarkGray

	# 🏁 <<<--- THÔNG TIN CƠ BẢN MÁY TÍNH --->>>

	# 🚩 <<<--- KIỂM TRA TỒN TẠI "C:\SW\IT.EXE.LNK" --->>>

	Write-Host "Current 'SOFTWARE' path  >>> " -NoNewline -ForegroundColor Cyan
	Write-Host $SourceSW -ForegroundColor Yellow
	Write-Host "Current 'SOFTWARE2' path >>> " -NoNewline -ForegroundColor Cyan
	if (Test-Path $SourceSW2) {
		Write-Host $SourceSW2 -ForegroundColor Yellow
	}
	else {
		Write-Host "Đường dẫn SOFTWARE2 không tồn tại" -ForegroundColor Red
	}
	Write-Host "Shortcut 'IT.exe.lnk' exists in >>> " -NoNewline -ForegroundColor Cyan

		# kiểm tra shortcut C:\SW
	if (Test-Path $SystemDriveSWlnk) {
		Write-Host $SystemDriveSW -NoNewline -ForegroundColor Yellow
	}
	else {
		Write-Host $SystemDriveSW " (không tồn tại)" -NoNewline -ForegroundColor Red
	}

	Write-Host " & " -NoNewline -ForegroundColor Cyan

		# kiểm tra shortcut Start Menu
	if (Test-Path $ExpandedStartMenuLnk) {
		Write-Host $StartMenuShortPath -ForegroundColor Yellow
	}
	else {
		Write-Host $StartMenuShortPath " (không tồn tại)" -ForegroundColor Red
	}
	Write-Host ("+" * $Host.UI.RawUI.WindowSize.Width) -ForegroundColor DarkGray

	# 🏁 <<<--- KIỂM TRA TỒN TẠI "C:\SW\IT.EXE.LNK" --->>>



	# 🚩 <<<--- HIỂN THỊ DANH SÁCH WINDOWS DEFENDER EXCLUSION --->>>

	Write-Host "<<< Current 'Windows Security\Exclusions' list >>>" -ForegroundColor Cyan

	$preferences = Get-MpPreference

	$paths = $preferences.ExclusionPath
	$proc  = $preferences.ExclusionProcess
	$ext   = $preferences.ExclusionExtension

	if (!$paths) { $paths = @("Không có") }
	if (!$proc)  { $proc  = @("Không có") }
	if (!$ext)   { $ext   = @("Không có") }

		# Tìm độ dài chuỗi lớn nhất mỗi cột
	$col1 = (($paths + "ExclusionPath") | Measure-Object Length -Maximum).Maximum + 3
	$col2 = (($proc  + "ExclusionProcess") | Measure-Object Length -Maximum).Maximum + 3
	$col3 = (($ext   + "ExclusionExtension") | Measure-Object Length -Maximum).Maximum + 3

	$max = ($paths.Count,$proc.Count,$ext.Count | Measure-Object -Maximum).Maximum

	Write-Host ("{0,-$col1}{1,-$col2}{2}" -f "ExclusionPath","ExclusionProcess","ExclusionExtension")
	Write-Host ("{0,-$col1}{1,-$col2}{2}" -f ("-"*13),("-"*16),("-"*18))

	for ($i=0; $i -lt $max; $i++) {

		$p1 = if ($i -lt $paths.Count) { $paths[$i] } else { "" }
		$p2 = if ($i -lt $proc.Count)  { $proc[$i] } else { "" }
		$p3 = if ($i -lt $ext.Count)   { $ext[$i] } else { "" }

		Write-Host ("{0,-$col1}{1,-$col2}{2}" -f $p1,$p2,$p3) -ForegroundColor Yellow
	}

	Write-Host ("+" * $Host.UI.RawUI.WindowSize.Width) -ForegroundColor DarkGray

	# 🏁 <<<--- DANH SÁCH WINDOWS DEFENDER EXCLUSION --->>>

	Write-Host "User vui lòng nhập số " -NoNewline
    Write-Host "115" -ForegroundColor Yellow -NoNewline
    Write-Host " để được hỗ trợ" -NoNewline
    Write-Host ": " -NoNewline

    $topMenu = Read-Host

    switch ($topMenu) {

		"111" { GoTo-IT-111 }
		"115" { GoTo-IT-115 }
		"113" { GoTo-IT-113 }
		default { return }

	}
}

function GoTo-IT-111 {
    Write-Host "`n→ Đang khởi động lại script..." -ForegroundColor Cyan
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""

    exit
}

function GoTo-IT-115 {
    Write-Host "`n→ Đang chuyển đến 'KHU VỰC NGƯỜI DÙNG'..." -ForegroundColor Cyan
	Run-IT-xxx "$IT115Script\IT-115.ps1" | Out-Null

	return
}

function GoTo-IT-113 {
    Write-Host "`n→ Đang chuyển đến 'KHU VỰC IT'..." -ForegroundColor Cyan
	Run-IT-xxx "$IT113Script\IT-113.ps1" | Out-Null

	return
}

	# bắt đầu chương trình
while ($true) {
    Show-Menu-IT
}

# 🏁 <<<--- CHỌN HỖ TRỢ 113 114 115 --->>>
