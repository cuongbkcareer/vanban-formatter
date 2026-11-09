#!/bin/bash
# Cài add-in "Chuẩn hoá văn bản hành chính" cho Microsoft Word trên macOS.
# Script chỉ chép file manifest.xml vào thư mục add-in của Word, không tải gì từ mạng và không cần quyền quản trị.
set -euo pipefail
cd "$(dirname "$0")"

WEF="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"

if [ ! -f manifest.xml ]; then
  echo "Không thấy manifest.xml cạnh script này. Hãy giải nén đủ file zip rồi chạy lại."
  exit 1
fi

if [ ! -d "/Applications/Microsoft Word.app" ]; then
  echo "Cảnh báo: không thấy /Applications/Microsoft Word.app. Cần Word cho Mac (Microsoft 365 hoặc 2021 trở lên)."
fi

if ! mkdir -p "$WEF" 2>/dev/null; then
  echo
  echo "Không tạo được thư mục: $WEF"
  echo "macOS đang chặn Terminal truy cập dữ liệu của Word. Cách xử lý:"
  echo "  1. Mở Word một lần rồi đóng lại (để Word tạo thư mục của nó)."
  echo "  2. Mở Cài đặt hệ thống > Quyền riêng tư & Bảo mật > Toàn quyền truy cập ổ đĩa,"
  echo "     bật cho Terminal, rồi chạy lại script này."
  echo "Hoặc làm thủ công: chép manifest.xml vào thư mục trên (Finder > Đi > Đi tới thư mục)."
  exit 1
fi

cp -f manifest.xml "$WEF/vanban-formatter.manifest.xml"

echo
echo "✔ Đã cài add-in vào: $WEF"
echo
echo "Bước tiếp theo:"
echo "  1. Thoát hoàn toàn Word (Cmd+Q) rồi mở lại."
echo "  2. Mở một tài liệu bất kỳ > thẻ Trang đầu (Home) > nút \"Chuẩn hoá NĐ 30\" ở nhóm \"Văn bản hành chính\"."
echo "  3. Nếu chưa thấy nút: Chèn (Insert) > Phần bổ trợ (Add-ins) > Phần bổ trợ của tôi (My Add-ins) > mục Developer Add-ins."
echo
echo "Cần Internet lần đầu để Word tải giao diện add-in. Nội dung văn bản của bạn không được gửi đi."
