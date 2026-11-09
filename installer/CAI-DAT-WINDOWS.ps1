# Cài add-in "Chuẩn hoá văn bản hành chính" cho Microsoft Word trên Windows (thư mục dùng chung đáng tin cậy).
# Script chép manifest.xml vào một thư mục cục bộ, chia sẻ thư mục đó và đăng ký làm "Trusted Add-in Catalog".
# Không tải gì từ mạng. Chỉ bước chia sẻ thư mục cần quyền quản trị (Windows sẽ hỏi UAC).
$ErrorActionPreference = 'Stop'

$here     = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifest = Join-Path $here 'manifest.xml'
if (-not (Test-Path $manifest)) { throw 'Không thấy manifest.xml cạnh script này. Hãy giải nén đủ file zip rồi chạy lại.' }

$dir       = Join-Path $env:USERPROFILE 'VanBanFormatterAddin'
$shareName = 'VanBanFormatterAddin'
$catalogId = '{3b7e2a64-9f1c-4d58-a0b2-5c6d7e8f9a10}'

New-Item -ItemType Directory -Force -Path $dir | Out-Null
Copy-Item $manifest (Join-Path $dir 'manifest.xml') -Force

# 1. Chia sẻ thư mục (cần quyền quản trị)
if (-not (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue)) {
  $user = "$env:USERDOMAIN\$env:USERNAME"
  $cmd  = "New-SmbShare -Name '$shareName' -Path '$dir' -ReadAccess '$user' | Out-Null"
  $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if ($isAdmin) { Invoke-Expression $cmd }
  else { Start-Process powershell -Verb RunAs -Wait -ArgumentList '-NoProfile', '-Command', $cmd }
}
if (-not (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue)) {
  Write-Host 'Không tạo được thư mục chia sẻ. Dùng cách thủ công trong HUONG-DAN.md (Upload My Add-in).' -ForegroundColor Yellow
  exit 1
}

# 2. Đăng ký thư mục chia sẻ làm catalog đáng tin cậy của Office (tài khoản hiện tại)
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"
New-Item -Path $key -Force | Out-Null
New-ItemProperty -Path $key -Name 'Id'    -Value $catalogId                       -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name 'Url'   -Value "\\$env:COMPUTERNAME\$shareName" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name 'Flags' -Value 1                                -PropertyType DWord  -Force | Out-Null

Write-Host ''
Write-Host 'Đã cài xong. Bước tiếp theo:' -ForegroundColor Green
Write-Host '  1. Thoát hoàn toàn Word rồi mở lại.'
Write-Host '  2. Chèn (Insert) > Phần bổ trợ (Add-ins) > Phần bổ trợ của tôi (My Add-ins) > thẻ SHARED FOLDER > chọn "Chuẩn hoá văn bản hành chính" > Add.'
Write-Host '  3. Nút "Chuẩn hoá NĐ 30" sẽ xuất hiện ở thẻ Trang đầu (Home).'
Write-Host 'Cần Internet lần đầu để Word tải giao diện add-in. Nội dung văn bản của bạn không được gửi đi.'
