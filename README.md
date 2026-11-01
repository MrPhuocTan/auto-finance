# Công Cụ Đối Soát & Gom Khoản Tiền Excel (Auto Finance)

Ứng dụng web tĩnh hỗ trợ kế toán tìm nhanh các tổ hợp hóa đơn / khoản thu chi trong file Excel khớp với một số tiền thanh toán cho trước (chính xác `0đ` hoặc trong biên độ sai số `± 500đ`).

- **Địa chỉ sử dụng trực tuyến**: `https://mrphuoctan.github.io/auto-finance/`
- **Bảo mật dữ liệu**: Toàn bộ quá trình đọc file Excel và tính toán chạy trực tiếp trên trình duyệt của máy người dùng, không gửi dữ liệu kế toán lên bất kỳ máy chủ bên ngoài nào.

## Tính năng chính

1. **Đọc trực tiếp sổ kế toán Excel (`.xlsx`, `.xls`, `.csv`)**:
   - Tự động dò dòng tiêu đề bảng (bỏ qua phần tên công ty, địa chỉ ở đầu trang).
   - Tự động nhận diện cột Ngày (`DD/MM/YYYY`), Số tiền (hỗ trợ cả số âm dạng `-500.000` hoặc ngoặc đơn `(500.000)` cho khoản chiết khấu / trả hàng / bù trừ công nợ), Diễn giải, Số chứng từ.
   - Tự động phát hiện và bỏ qua dòng `Tổng cộng / Cộng phát sinh` cuối bảng để tránh sai lệch kết quả.
2. **Lọc khoảng ngày `Từ ngày - Đến ngày` & Số tiền mục tiêu**:
   - Chọn nhanh theo từng tháng có trong file hoặc nhập khoảng ngày bất kỳ.
   - Tự động phân cách hàng nghìn khi nhập/dán số tiền và hiển thị dòng **đọc số tiền bằng chữ tiếng Việt** để tránh nhầm số `0`.
   - Mặc định tìm cả phương án **khớp chính xác (`0đ`)** và **gần khớp (`± 500đ`)**, có thể tùy chỉnh biên độ sai số.
3. **Tiện ích nghiệp vụ kế toán**:
   - Ghim dòng bắt buộc phải có hoặc loại trừ dòng bất kỳ trực tiếp trên bảng sổ chi tiết.
   - Đánh dấu **Đã gom theo từng đợt (`Đợt 1`, `Đợt 2`...)** để đối soát liên tiếp nhiều khoản chuyển khoản trên cùng một file mà không bị trùng lặp hóa đơn.
   - Sao chép nhanh tổ hợp sang bộ nhớ tạm để dán thẳng vào Excel (`Ctrl + V`) hoặc xuất toàn bộ kết quả ra file `.xlsx` mới.
   - Chế độ **Cỡ chữ Lớn** và giao diện bảng biểu truyền thống rõ nét dành cho kế toán.

## Chạy thử & Kiểm thử tại máy cục bộ

```bash
npm test
npx serve .
```
