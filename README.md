# vanban-formatter

Biến văn bản thô thành văn bản hành chính đúng thể thức theo **NĐ 30/2020/NĐ-CP (Phụ lục I)**, xuất file `.docx`.
Chạy hoàn toàn offline: không gọi mạng, không dùng LLM, không gửi dữ liệu đi đâu.

## Cách dùng

```bash
npm install
npm run cli -- fixtures/cong-van.txt --preview          # xem từng dòng được nhận diện thế nào
npm run cli -- fixtures/cong-van.txt -o out/cong-van.docx
npm run cli -- raw.txt --issuer "SỞ TÀI CHÍNH" --place "Hà Nội" --date 2025-03-05 -o out.docx
npm run build                                           # tạo docs/vanban.html (một file, mở bằng trình duyệt)
npm test
```

Giao diện web (`docs/vanban.html`) cho phép dán văn bản, **sửa nhãn từng dòng** rồi mới tải `.docx`. Trang có CSP chặn mọi kết nối ra ngoài.
Nếu trình duyệt không mở được file trực tiếp, chạy `python3 -m http.server --directory docs` rồi vào `http://127.0.0.1:8000/vanban.html`.

## Add-in cho Word (Windows và Mac)

Khung bên cạnh trong Word: lấy văn bản từ tài liệu đang mở, nhận diện, cho sửa nhãn, rồi **mở thành tài liệu mới** đúng thể thức (hoặc thay nội dung tài liệu hiện tại, chế độ thử nghiệm).
Cài trên Mac bằng một dòng lệnh trong Terminal:

```bash
curl -fsSL https://cuongbkcareer.github.io/vanban-formatter/cai-dat-mac.sh | bash
```

Chi tiết (Mac, Windows, cả cơ quan): [HUONG-DAN-CAI-DAT.md](HUONG-DAN-CAI-DAT.md). File zip nằm ở mục Releases; giao diện được host tại GitHub Pages (thư mục `docs/`).

```bash
npm run build      # tạo docs/ (taskpane, manifest, bản web độc lập)
npm run release    # build + tạo dist/vanban-formatter-addin-vX.Y.Z.zip
ADDIN_BASE_URL=https://may-chu-noi-bo/vanban npm run build   # manifest trỏ tới máy chủ tự host
```

## Kiến trúc

```
văn bản thô → classify (nhãn từng dòng + độ tin cậy) → assemble (Model + tuỳ chọn) → buildIR → renderDocx
```

| Lớp | File | Ghi chú |
|---|---|---|
| IR dùng chung | `src/ir.ts`, `src/render-docx.ts` | Mọi profile đều sinh IR, chỉ có một renderer `.docx` |
| Profile NĐ 30 | `src/profiles/vn-nd30/` | `classify.ts` (parser), `assemble.ts`, `build-ir.ts`, `layout.json` (quy cách trình bày) |
| Giao diện | `src/cli.ts`, `web/` | |

**Thêm một định dạng mới** (ví dụ thư tiếng Anh): tạo `src/profiles/<tên>/` gồm parser (văn bản → các khối), `layout.json` và hàm dựng IR; dùng lại `renderDocx`.
Layout là dữ liệu (JSON), còn parser là code riêng cho mỗi ngôn ngữ và thể loại.

## Phạm vi hiện tại

- Nhận diện: cơ quan chủ quản/ban hành, số ký hiệu, địa danh ngày tháng, tên loại (28 loại + công văn), trích yếu, kính gửi, thẩm quyền ban hành, căn cứ, chương/điều, tiêu đề I./1./a), nơi nhận, chức vụ và họ tên người ký.
- Tự sinh: quốc hiệu, tiêu ngữ, đường kẻ, bảng phần đầu, bảng nơi nhận/chữ ký, số trang (ẩn trang 1), ngày tháng theo quy tắc thêm số 0.
- Chưa có: phụ lục, dấu "Mục", bảng biểu trong nội dung, nhiều cấp cơ quan chủ quản phức tạp, văn bản mật/khẩn (dấu độ khẩn, độ mật), chữ ký số.

## Lưu ý quan trọng

- **Số liệu trong `layout.json` được nhập từ trí nhớ, CHƯA đối chiếu với văn bản gốc** (`"verified": false`). Hãy đối chiếu từng mục với Phụ lục I và các mẫu kèm theo, và xác nhận NĐ 30/2020 còn hiệu lực, trước khi dùng chính thức.
- Đường kẻ dưới cơ quan ban hành, tiêu ngữ và trích yếu được ước lượng độ dài theo công thức, không đo chính xác.
- Văn bản thô luôn có ca nhận diện sai: hãy xem bước kiểm tra nhãn (dòng nền vàng là độ tin cậy thấp).
