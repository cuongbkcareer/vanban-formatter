# Cài add-in "Chuẩn hoá văn bản hành chính" cho Word

Add-in thêm nút **Chuẩn hoá NĐ 30** vào thẻ Trang đầu (Home) của Word. Nó đọc văn bản thô trong tài liệu, nhận diện từng thành phần, cho bạn kiểm tra rồi tạo văn bản đúng thể thức theo NĐ 30/2020/NĐ-CP.

**Yêu cầu:** Word trong Microsoft 365, hoặc Office 2021/2024, trên Windows hoặc Mac. Office 2016/2019 bản cũ dùng trình duyệt Internet Explorer nên **không** chạy được.

## Quyền riêng tư

- Toàn bộ xử lý nằm trong Word trên máy bạn. Mã của dự án không gọi mạng (có test tự động kiểm tra), không dùng AI, không gửi nội dung văn bản đi đâu.
- Giao diện add-in được tải từ GitHub Pages (`https://cuongbkcareer.github.io/vanban-formatter/`), và Word tải thư viện `office.js` từ máy chủ của Microsoft. Hai nguồn này nhận được yêu cầu tải file, **không** nhận được nội dung văn bản của bạn.
- Cơ quan muốn kiểm soát chặt hơn có thể tự host: xem mục cuối.

## macOS (cách dễ nhất: một dòng lệnh)

1. Mở **Terminal** (Cmd+Space, gõ `Terminal`).
2. Dán dòng sau rồi nhấn Enter:

```bash
curl -fsSL https://cuongbkcareer.github.io/vanban-formatter/cai-dat-mac.sh | bash
```

3. Nếu macOS hỏi **"Terminal muốn điều khiển Finder"**, bấm **Cho phép** (chỉ một lần). Lý do: macOS không cho Terminal ghi vào dữ liệu của Word, nhưng Finder thì được, nên script nhờ Finder chép giúp.
4. Thoát hẳn Word (Cmd+Q), mở lại, rồi vào thẻ Trang đầu tìm nút **Chuẩn hoá NĐ 30**.

Nếu báo "Chưa có thư mục dữ liệu của Word": mở Word một lần, thoát (Cmd+Q), chạy lại lệnh.

Script chỉ chép một file (`manifest.xml`) vào `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`. Muốn đọc trước khi chạy: mở https://cuongbkcareer.github.io/vanban-formatter/cai-dat-mac.sh trong trình duyệt. Dùng `curl` nên file không bị gắn cờ tải từ Internet, và bạn không gặp cảnh báo "Apple không thể xác minh…".

**Gỡ cài đặt:**

```bash
curl -fsSL https://cuongbkcareer.github.io/vanban-formatter/cai-dat-mac.sh | bash -s -- --uninstall
```

### Nếu đã tải file zip

Giải nén zip, mở Terminal, gõ `bash ` (có dấu cách), kéo file `CAI-DAT-MAC.command` vào cửa sổ Terminal rồi Enter. Làm theo cách này không bị cảnh báo của macOS. (Bấm đúp vào file sẽ bị Gatekeeper chặn vì file chưa có chữ ký của Apple. Muốn bấm đúp thì vào Cài đặt hệ thống > Quyền riêng tư & Bảo mật > Vẫn mở.)

### Cài thủ công bằng Finder (không dùng Terminal)

1. Mở Word một lần rồi thoát (Cmd+Q).
2. Finder > Đi > Đi tới thư mục, dán `~/Library/Containers/com.microsoft.Word/Data/Documents`.
3. Tạo thư mục mới tên `wef` (chữ thường) nếu chưa có, kéo `manifest.xml` vào đó.
4. Mở lại Word.

Nếu không thấy nút trên thẻ Trang đầu: Chèn > Phần bổ trợ > Phần bổ trợ của tôi > Developer Add-ins.

## Nút không hiện trên thẻ Trang đầu (Home)?

Theo tài liệu của Microsoft, add-in cài thủ công trên Mac luôn tìm được ở **Home > Add-ins** (mục Developer Add-ins). Nút riêng trên thanh ribbon là tính năng thêm, không phải lúc nào cũng hiện với add-in cài thủ công.

Nếu bạn vừa cài bản mới (v0.1.2 trở lên) mà vẫn thấy bản cũ hoặc chưa thấy nút, Word đang giữ manifest cũ trong bộ nhớ đệm:

1. Thoát hẳn Word (Cmd+Q).
2. Chạy lại lệnh cài ở trên (để chép manifest mới).
3. Xoá **toàn bộ nội dung** (không xoá từng file riêng lẻ) của các thư mục sau bằng Finder (Đi > Đi tới thư mục): `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` và `~/Library/Containers/com.microsoft.Word/Data/Library/Caches/`. Đừng xoá thư mục `wef` ở bước này, vì đó là nơi chứa manifest vừa cài.
4. Mở lại Word, mở một tài liệu, vào Home > Add-ins.

## Windows

> Phần Windows được viết theo tài liệu của Microsoft nhưng **chưa được kiểm thử trên máy Windows thật**. Nếu gặp lỗi, hãy dùng cách 2 và báo lại.

**Cách 1: dùng script.** Giải nén zip, bấm đúp `CAI-DAT-WINDOWS.bat`, đồng ý hộp thoại UAC khi Windows hỏi (cần để chia sẻ một thư mục cục bộ). Mở lại Word, vào Chèn > Phần bổ trợ > Phần bổ trợ của tôi > thẻ SHARED FOLDER > chọn add-in > Add.

**Cách 2: tải lên thủ công.** Trong Word: Chèn > Phần bổ trợ > Phần bổ trợ của tôi > mũi tên bên cạnh "Quản lý phần bổ trợ của tôi" > **Tải lên phần bổ trợ** (Upload My Add-in) > chọn `manifest.xml`. Cách này mất khi bạn đóng Word ở một số phiên bản; khi đó dùng cách 1.

## Triển khai cho cả cơ quan

Nếu cơ quan có tenant Microsoft 365: Trung tâm quản trị Microsoft 365 > Cài đặt > Ứng dụng tích hợp > Ứng dụng đã triển khai > Tải lên ứng dụng tùy chỉnh > chọn `manifest.xml`, rồi gán cho người dùng. Người dùng không cần cài gì.

## Tự host (cơ quan cần kiểm soát mã)

1. Chép thư mục `tu-host/` của file zip lên máy chủ nội bộ phục vụ bằng **HTTPS**.
2. Mở `manifest.xml`, thay mọi chuỗi `https://cuongbkcareer.github.io/vanban-formatter` bằng địa chỉ máy chủ của bạn (kể cả dòng `AppDomain`).
3. Cài lại manifest như các bước trên.

Word vẫn cần tải `office.js` từ `appsforoffice.microsoft.com`; Microsoft không cho phép đóng gói thư viện này riêng.

## Gỡ cài đặt trên Windows

Xoá khoá registry `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{3b7e2a64-9f1c-4d58-a0b2-5c6d7e8f9a10}`, chạy `Remove-SmbShare -Name VanBanFormatterAddin` (PowerShell quản trị) và xoá thư mục `%USERPROFILE%\VanBanFormatterAddin`.

## Cập nhật

Giao diện cập nhật tự động khi bản mới được đưa lên GitHub Pages (có thể cần đóng và mở lại khung add-in). Chỉ khi manifest đổi mới cần cài lại.
