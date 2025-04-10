# Ứng Dụng Chuyển Đổi PDF Ảnh Sang Word

Ứng dụng web giúp chuyển đổi tệp PDF có chứa hình ảnh quét sang tài liệu Word có cấu trúc, với mỗi trang PDF được chuyển thành một hình ảnh trong tệp Word.

## Tính Năng Chính

- Tải lên và xử lý tệp PDF
- Chuyển đổi từng trang PDF thành hình ảnh có chất lượng cao
- Tạo tệp Word (.docx) chứa tất cả các hình ảnh với định dạng đẹp
- Theo dõi tiến trình chuyển đổi với thanh tiến độ
- Hiển thị nhật ký chi tiết trong quá trình chuyển đổi
- Xử lý hoàn toàn trong trình duyệt, không cần kết nối internet (ngoại trừ lúc tải worker PDF.js)

## Cài Đặt

1. Clone repository:
```
git clone https://github.com/yourusername/pdf-word.git
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