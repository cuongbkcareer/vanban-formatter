#!/bin/bash
# Cài add-in "Chuẩn hoá văn bản hành chính" cho Microsoft Word trên macOS.
#
#   Cài:  bash cai-dat-mac.sh        (hoặc dán lệnh một dòng trong HUONG-DAN-CAI-DAT.md)
#   Gỡ:   bash cai-dat-mac.sh --uninstall
#
# Script chỉ chép MỘT file (manifest.xml) vào thư mục add-in của Word. Không cần quyền quản trị,
# không cài phần mềm nào khác. Nội dung script này ngắn, bạn có thể đọc trước khi chạy.
#
# Vì sao có bước "Finder": macOS không cho Terminal ghi vào dữ liệu của Word, nhưng Finder thì được.
# Script thử chép trực tiếp trước; nếu bị chặn thì nhờ Finder chép (macOS sẽ hỏi
# "Terminal muốn điều khiển Finder", bạn bấm Cho phép một lần).
set -uo pipefail

BASE_URL="${VANBAN_BASE_URL:-__BASE_URL__}"
ADDIN_ID="81a47bdf-3c91-4d65-9322-112a4078e6c5"
NAME="vanban-formatter.manifest.xml"
DOCS="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
# VANBAN_WEF_DIR chỉ dùng để kiểm thử: chép thẳng vào thư mục chỉ định, bỏ qua Finder.
WEF="${VANBAN_WEF_DIR:-$DOCS/wef}"

say() { printf '%s\n' "$*"; }
die() { printf '\n✘ %s\n' "$*" >&2; exit 1; }

TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

# --- Gỡ cài đặt -------------------------------------------------------------
if [ "${1:-}" = "--uninstall" ]; then
  if [ -n "${VANBAN_WEF_DIR:-}" ]; then
    rm -f "$WEF/$NAME" && say "✔ Đã gỡ add-in." && exit 0
  fi
  if rm -f "$WEF/$NAME" 2>/dev/null && [ ! -e "$WEF/$NAME" ]; then
    say "✔ Đã gỡ add-in. Thoát hẳn Word (Cmd+Q) rồi mở lại."; exit 0
  fi
  osascript - "$DOCS" "$NAME" >/dev/null 2>&1 <<'APPLESCRIPT'
on run argv
  set docsF to (POSIX file (item 1 of argv)) as alias
  tell application "Finder"
    if exists file (item 2 of argv) of folder "wef" of docsF then delete file (item 2 of argv) of folder "wef" of docsF
  end tell
end run
APPLESCRIPT
  say "✔ Đã gỡ add-in (nếu có). Thoát hẳn Word (Cmd+Q) rồi mở lại."; exit 0
fi

# --- Kiểm tra điều kiện -----------------------------------------------------
[ "$(uname)" = "Darwin" ] || die "Script này chỉ dành cho macOS."
[ -d "/Applications/Microsoft Word.app" ] || say "Cảnh báo: không thấy /Applications/Microsoft Word.app (cần Word cho Mac, Microsoft 365 hoặc 2021 trở lên)."
if pgrep -x "Microsoft Word" >/dev/null 2>&1; then
  say "Lưu ý: Word đang mở. Cài xong hãy thoát hẳn Word (Cmd+Q) rồi mở lại."
fi

# --- Lấy manifest.xml: cạnh script nếu có, nếu không thì tải từ trang của dự án --------
SRC="$TMP/$NAME"
HERE="${BASH_SOURCE[0]:-}"
if [ -n "$HERE" ] && [ -f "$HERE" ] && [ -f "$(dirname "$HERE")/manifest.xml" ]; then
  cp "$(dirname "$HERE")/manifest.xml" "$SRC"
else
  say "Đang tải manifest.xml từ $BASE_URL ..."
  curl -fsSL --proto '=https' --tlsv1.2 --max-time 30 "$BASE_URL/manifest.xml" -o "$SRC" \
    || die "Không tải được manifest.xml (kiểm tra Internet) hoặc dùng bản zip có sẵn manifest.xml."
fi
grep -q "$ADDIN_ID" "$SRC" 2>/dev/null || die "File manifest.xml không đúng (thiếu mã add-in)."

# --- Chép vào thư mục add-in của Word ---------------------------------------
install_direct() {
  mkdir -p "$WEF" 2>/dev/null && cp -f "$SRC" "$WEF/$NAME" 2>/dev/null && [ -f "$WEF/$NAME" ]
}

install_via_finder() {
  local out
  out="$(osascript - "$SRC" "$DOCS" 2>&1 <<'APPLESCRIPT'
on run argv
  set srcFile to (POSIX file (item 1 of argv)) as alias
  try
    set docsF to (POSIX file (item 2 of argv)) as alias
  on error
    return "NO_DOCS"
  end try
  tell application "Finder"
    if not (exists folder "wef" of docsF) then make new folder at docsF with properties {name:"wef"}
    duplicate srcFile to folder "wef" of docsF with replacing
    if exists file "vanban-formatter.manifest.xml" of folder "wef" of docsF then return "OK"
    return "FAILED"
  end tell
end run
APPLESCRIPT
)"
  case "$out" in
    OK) return 0 ;;
    NO_DOCS) die "Chưa có thư mục dữ liệu của Word. Hãy mở Word một lần, thoát (Cmd+Q) rồi chạy lại." ;;
    *) say "Finder báo: $out"; return 1 ;;
  esac
}

if install_direct; then
  :
elif [ -n "${VANBAN_WEF_DIR:-}" ]; then
  die "Không chép được vào $WEF."
else
  say "Terminal bị macOS chặn ghi vào dữ liệu của Word, chuyển sang nhờ Finder chép..."
  say "(Nếu macOS hỏi \"Terminal muốn điều khiển Finder\", hãy bấm Cho phép.)"
  install_via_finder || die "Không cài được. Làm thủ công: chép manifest.xml vào $WEF (xem HUONG-DAN-CAI-DAT.md)."
fi

say
say "✔ Đã cài add-in."
say
say "Bước tiếp theo:"
say "  1. Thoát hẳn Word (Cmd+Q) rồi mở lại."
say "  2. Mở một tài liệu bất kỳ > thẻ Trang đầu (Home) > nút \"Chuẩn hoá NĐ 30\" (nhóm \"Văn bản hành chính\")."
say "  3. Nếu chưa thấy nút: Chèn (Insert) > Phần bổ trợ (Add-ins) > Phần bổ trợ của tôi (My Add-ins) > Developer Add-ins."
say
say "Cần Internet khi mở khung add-in lần đầu. Nội dung văn bản của bạn không được gửi đi."
