# Cài add-in "Chuẩn hoá văn bản hành chính" cho Word

Add-in thêm nút **Chuẩn hoá NĐ 30** vào thẻ Trang đầu (Home) của Word. Nó đọc văn bản thô trong tài liệu, nhận diện từng thành phần, cho bạn kiểm tra rồi tạo văn bản đúng thể thức theo NĐ 30/2020/NĐ-CP.

**Yêu cầu:** Word trong Microsoft 365, hoặc Office 2021/2024, trên Windows hoặc Mac. Office 2016/2019 bản cũ dùng trình duyệt Internet Explorer nên **không** chạy được.

## Quyền riêng tư

- Toàn bộ xử lý nằm trong Word trên máy bạn. Mã của dự án không gọi mạng (có test tự động kiểm tra), không dùng AI, không gửi nội dung văn bản đi đâu.
- Giao diện add-in được tải từ GitHub Pages (`https://cuongbkcareer.github.io/vanban-formatter/`), và Word tải thư viện `office.js` từ máy chủ của Microsoft. Hai nguồn này nhận được yêu cầu tải file, **không** nhận được nội dung văn bản của bạn.
- Cơ quan muốn kiểm soát chặt hơn có thể tự host: xem mục cuối.

## macOS

**Cách 1: dùng script.** Giải nén zip, chuột phải `CAI-DAT-MAC.command` > Mở (Open) > Mở. (macOS hỏi vì file tải từ Internet.) Làm theo hướng dẫn hiện ra.

**Cách 2: thủ công.** Chép `manifest.xml` vào thư mục sau (Finder > Đi > Đi tới thư mục, dán đường dẫn). Mở Word một lần trước nếu thư mục chưa có.

```
~/Library/Containers/com.microsoft.Word/Data/Documents/wef
```

Thoát hẳn Word (Cmd+Q), mở lại. Nút nằm ở thẻ Trang đầu. Nếu không thấy: Chèn > Phần bổ trợ > Phần bổ trợ của tôi > Developer Add-ins.

Nếu macOS chặn Terminal truy cập thư mục của Word: Cài đặt hệ thống > Quyền riêng tư & Bảo mật > Toàn quyền truy cập ổ đĩa > bật Terminal.

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

## Gỡ cài đặt

- **Mac:** xoá file `vanban-formatter.manifest.xml` trong thư mục `wef` ở trên.
- **Windows:** xoá khoá registry `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{3b7e2a64-9f1c-4d58-a0b2-5c6d7e8f9a10}`, chạy `Remove-SmbShare -Name VanBanFormatterAddin` (PowerShell quản trị) và xoá thư mục `%USERPROFILE%\VanBanFormatterAddin`.

## Cập nhật

Giao diện cập nhật tự động khi bản mới được đưa lên GitHub Pages (có thể cần đóng và mở lại khung add-in). Chỉ khi manifest đổi mới cần cài lại.
