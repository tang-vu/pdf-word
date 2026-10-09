# Ứng Dụng Chuyển Đổi PDF Ảnh Sang Word

Ứng dụng web giúp chuyển đổi tệp PDF có chứa hình ảnh quét sang tài liệu Word có cấu trúc, với mỗi trang PDF được chuyển thành một hình ảnh trong tệp Word.

## Tính Năng Chính

- Tải lên và xử lý tệp PDF
- Chuyển đổi từng trang PDF thành hình ảnh có chất lượng cao
- Tạo tệp Word (.docx) chứa tất cả các hình ảnh với định dạng đẹp
- Theo dõi tiến trình chuyển đổi với thanh tiến độ
- Hiển thị nhật ký chi tiết trong quá trình chuyển đổi
- Xử lý PDF hoàn toàn trong trình duyệt; worker PDF.js được tải từ cùng máy chủ với ứng dụng, không qua CDN

## Cài Đặt

1. Clone repository:
```
git clone https://github.com/tang-vu/pdf-word.git
cd pdf-word
```

2. Cài đặt các thư viện phụ thuộc:
```
npm install
```

3. Chạy ứng dụng ở chế độ phát triển:
```
npm run dev
```

4. Để xây dựng sản phẩm:
```
npm run build
```

## Kiểm Tra

```
npm test
npm run lint
npm run build
```

Các kiểm thử dùng trình chạy tích hợp của Node.js, TypeScript và các thư viện hiện có, không cần cài thêm thư viện kiểm thử. Kiểm thử thực thi hàm chuyển đổi của component với bộ mô phỏng React state, PDF.js, canvas và thao tác tải xuống; tệp DOCX được đóng gói thật và kiểm tra dữ liệu ảnh, thứ tự trang, bố cục, các lỗi giữa chừng và thao tác thử lại. Việc hiển thị PDF và tải worker trong trình duyệt cần được kiểm tra riêng.

Nếu bất kỳ trang nào không thể chuyển đổi, ứng dụng hiển thị số trang gặp lỗi và không tạo hoặc tải xuống tệp Word thiếu trang. Chỉ báo 100% khi tệp DOCX hoàn chỉnh đã được tạo và bắt đầu tải xuống.

## Bảo Mật

Ứng dụng đặt `isEvalSupported: false` theo biện pháp giảm thiểu được Mozilla công bố trong [GHSA-wgrm-67xf-hhpq](https://github.com/mozilla/pdf.js/security/advisories/GHSA-wgrm-67xf-hhpq). Đây là thay đổi cấu hình; dự án vẫn dùng PDF.js 3.11 và các cảnh báo kiểm toán dependency vẫn còn. Việc nâng cấp PDF.js được dành cho một thay đổi riêng. Worker classic được Vite đóng gói từ đúng package PDF.js đã cài, cùng origin với ứng dụng. Giấy phép upstream được giữ nguyên tại `public/pdfjs-LICENSE.txt` và được Vite sao chép vào `dist/pdfjs-LICENSE.txt`; header giấy phép trong worker cũng được giữ nguyên.

Với đường dẫn triển khai lồng nhau, dùng ví dụ `npm run build -- --base=/pdf-word/` và phục vụ toàn bộ thư mục `dist`, bao gồm worker. Vite tạo URL worker theo `base`. Trình duyệt vẫn cần kết nối tới máy chủ để tải ứng dụng và worker; dự án không cung cấp service worker hay cam kết hoạt động offline.

## Công Nghệ Sử Dụng

- React
- TypeScript
- Vite
- PDF.js
- docx.js
- TailwindCSS

## Lưu Ý

- Ứng dụng hoạt động tốt nhất với các tệp PDF có chứa hình ảnh quét
- Với các tệp PDF lớn, quá trình chuyển đổi có thể mất nhiều thời gian
- Tệp Word đầu ra sẽ chứa các hình ảnh được trích xuất từ PDF với độ phân giải tương tự bản gốc

## Giấy Phép

MIT 
