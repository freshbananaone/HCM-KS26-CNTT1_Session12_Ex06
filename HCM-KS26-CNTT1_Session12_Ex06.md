# Bài 6: Hoàn thiện bộ hồ sơ triển khai nền tảng Smart-Home

**Smart-Home** là nền tảng đấu giá nhà đất trực tuyến, có thêm chức năng quản lý bất động sản cho thuê. Báo cáo gồm nội dung SRS, mẫu tính toán Excel và kịch bản Pitch Deck. Các mức giá, chiết khấu, hoa hồng và số liệu gọi vốn dưới đây là giả định phục vụ bài tập.

## Phần 1. Phân tích

Một bảng Excel tốt cần phục vụ được việc tính toán, lọc và tra cứu. Bảng trông đẹp nhưng gộp ô hoặc nhập số dưới dạng chữ vẫn dễ gây sai khi xử lý dữ liệu.

| Tiêu chí | Bảng “xấu” | Bảng “đúng” |
|---|---|---|
| Tiêu đề | Nhiều tầng tiêu đề, tên cột không rõ nghĩa. | Một hàng tiêu đề, mỗi cột có tên riêng. |
| Field — trường dữ liệu | Một cột chứa cả địa chỉ, giá và tên môi giới. | Mỗi cột mô tả một thuộc tính, như mã nhà đất hoặc giá thuê. |
| Record — bản ghi | Một bất động sản trải trên nhiều hàng, khó xác định dữ liệu thuộc dòng nào. | Trong bảng danh mục, mỗi hàng là một bất động sản có mã duy nhất. |
| Gộp ô | Merge Cell để nhóm các nhà cùng loại hoặc cùng môi giới. | Không Merge Cell; thông tin chung được điền đầy đủ trên từng hàng. |
| Kiểu dữ liệu | Nhập `10 triệu/tháng` trực tiếp vào ô giá. | Nhập số `10000000`, dùng định dạng để hiển thị tiền tệ. |
| Tính liên tục | Có hàng trống, tiêu đề lặp hoặc tổng cộng chen giữa. | Dữ liệu liền mạch; tổng hợp đặt ở sheet riêng. |
| Tra cứu | Nhập tên môi giới tùy ý, mã trùng hoặc vùng tra cứu thay đổi khi kéo. | Dùng mã thống nhất, bảng tra có khóa duy nhất và vùng tham chiếu cố định. |

Với dự án này, em tách danh mục nhà đất và danh sách môi giới thành hai sheet để dễ cập nhật tỷ lệ hoa hồng mà không phải sửa từng dòng.

## Phần 2. Triển khai

### 2.1. Nội dung SRS dùng cho file Word

**Cấu trúc Heading**

```text
Heading 1: SRS — Nền tảng Smart-Home
├── Heading 2: 1. Mục đích và phạm vi
├── Heading 2: 2. Người dùng và phân quyền
├── Heading 2: 3. Quy trình nghiệp vụ
├── Heading 2: 4. Yêu cầu chức năng
│   ├── Heading 3: 4.1. Tài khoản và hồ sơ nhà đất
│   ├── Heading 3: 4.2. Phiên đấu giá và trả giá
│   └── Heading 3: 4.3. Cho thuê và hoa hồng
├── Heading 2: 5. Dữ liệu và quy tắc nghiệp vụ
├── Heading 2: 6. Yêu cầu phi chức năng
└── Heading 2: 7. Tiêu chí nghiệm thu
```

**1. Mục đích và phạm vi:** giúp chủ nhà, người mua và môi giới theo dõi giao dịch trên cùng một nền tảng. Phiên bản đầu hỗ trợ đăng tin, duyệt hồ sơ, đấu giá tăng dần, quản lý giá thuê và tính hoa hồng. Kết quả đấu giá được chuyển sang bước hoàn tất giao dịch; chưa triển khai tự động thanh toán và sang tên.

**2. Người dùng và phân quyền:** chủ nhà quản lý tin của mình; người mua đăng ký phiên và trả giá; người thuê gửi yêu cầu thuê; môi giới theo dõi bất động sản được giao; quản trị viên duyệt tin, quản lý phiên và xử lý báo cáo.

**3. Quy trình nghiệp vụ:** chủ nhà gửi hồ sơ → quản trị viên duyệt → công bố phiên và điều kiện tham gia → người mua đăng ký → trả giá → đóng phiên → thông báo kết quả. Luồng cho thuê gồm xem tin → gửi yêu cầu → xác nhận thời hạn, mức giảm → tính tiền thuê và hoa hồng.

**4. Yêu cầu chức năng**

| Nhóm | Mã | Yêu cầu |
|---|---|---|
| 4.1. Tài khoản và hồ sơ | FR-01 | Người dùng đăng nhập và thao tác theo vai trò; không được sửa tin thuộc tài khoản khác. |
| 4.1. Tài khoản và hồ sơ | FR-02 | Chủ nhà nhập thông tin, hình ảnh và hồ sơ; chỉ tin được quản trị viên duyệt mới hiển thị công khai. |
| 4.2. Phiên đấu giá | FR-03 | Quản trị viên thiết lập giá khởi điểm, bước giá, giờ mở/đóng và điều kiện tham gia trước khi công bố phiên. |
| 4.2. Phiên đấu giá | FR-04 | Hệ thống chỉ nhận giá của người đủ điều kiện trong thời gian mở phiên; từ chối giá không hợp lệ và lưu lịch sử giá được chấp nhận. |
| 4.2. Phiên đấu giá | FR-05 | Khi đóng phiên, ghi nhận người có giá hợp lệ cao nhất và gửi thông báo; nếu chưa có giá hợp lệ, ghi nhận phiên không có người thắng. |
| 4.3. Cho thuê và hoa hồng | FR-06 | Tính giá thuê sau chiết khấu, tổng tiền thuê theo số tháng và hoa hồng theo mã môi giới; hiển thị rõ căn cứ tính. |

**5. Dữ liệu và quy tắc nghiệp vụ:** dữ liệu chính gồm người dùng, bất động sản, môi giới, phiên đấu giá, lượt trả giá và yêu cầu thuê. Giá đầu tiên phải từ giá khởi điểm trở lên; các giá tiếp theo tối thiểu bằng giá hiện tại cộng bước giá. Dùng thời gian máy chủ, chỉ nhận giá khi `giờ mở ≤ thời điểm nhận < giờ đóng`. Các lượt gửi đồng thời phải được xử lý tuần tự để tránh xác nhận hai mức giá dựa trên cùng một giá cũ. Chiết khấu nằm trong khoảng 0–100%, số tháng thuê là số nguyên dương. Chiết khấu thuê không áp dụng cho giá trúng đấu giá.

**6. Yêu cầu phi chức năng:** mục tiêu thử nghiệm là 200 người truy cập đồng thời, 95% yêu cầu xem tin phản hồi trong 3 giây. Hệ thống dùng HTTPS, kiểm tra quyền ở máy chủ, sao lưu hằng ngày và lưu nhật ký thay đổi giá, hồ sơ, trạng thái phiên. Đây là chỉ tiêu cần kiểm thử, chưa phải số liệu vận hành.

**7. Tiêu chí nghiệm thu:** kiểm tra đủ FR-01 đến FR-06; chặn người không có quyền, giá dưới bước giá và giá đến sau giờ đóng. Khi có hai lượt trả giá đồng thời, kết quả phải nhất quán với thứ tự xử lý. Phần tiền thuê và hoa hồng phải khớp các ví dụ tính toán bên dưới.

**Trình bày Word:** dùng đúng Heading 1/2/3 theo sơ đồ, kiểm tra bằng Navigation Pane và tạo mục lục tự động. Trang đầu ghi tên dự án, phiên bản, ngày cập nhật; thân bài dùng một phông chữ thống nhất, có số trang. Trước khi gửi đối tác, cập nhật toàn bộ mục lục và rà lại tên mục, số trang.

### 2.2. Mẫu Excel quản lý nhà đất và tính toán

**Sheet `NhaDat`:** hàng 1 là header; mỗi dòng dữ liệu là một bất động sản cho thuê. Các cột được bố trí như sau:

```text
A: Mã BĐS                 E: Tỷ lệ chiết khấu      I: Mã môi giới
B: Loại BĐS               F: Tiền giảm/tháng       J: Tỷ lệ hoa hồng
C: Giá thuê gốc/tháng     G: Giá thuê sau giảm     K: Tiền hoa hồng
D: Số tháng thuê          H: Tổng tiền thuê
```

Các trường tiền dùng đơn vị **đồng**, riêng cột C, F, G tính theo **tháng**. Cột E và J định dạng phần trăm. Dữ liệu đầu vào tại hàng 2–5:

| A — Mã BĐS | B — Loại BĐS | C — Giá thuê gốc/tháng | D — Số tháng | E — Chiết khấu | I — Mã môi giới |
|---|---|---:|---:|---:|---|
| BDS001 | Căn hộ | 10.000.000 | 12 | 5% | MG01 |
| BDS002 | Nhà riêng | 8.000.000 | 6 | 0% | MG02 |
| BDS003 | Nhà phố | 15.000.000 | 12 | 10% | MG03 |
| BDS004 | Căn hộ | 12.000.000 | 3 | 5% | MG01 |

**Sheet `HoaHong`:** bảng tra nằm tại A1:C4, với dữ liệu tại A2:C4.

| A — Mã môi giới | B — Tên môi giới | C — Tỷ lệ hoa hồng |
|---|---|---:|
| MG01 | Nguyễn Minh An | 50% |
| MG02 | Trần Thu Hà | 40% |
| MG03 | Lê Hoàng Nam | 60% |

Trong bài tập, hoa hồng trả **một lần**, bằng tỷ lệ quy định nhân với **một tháng tiền thuê sau giảm**, do chủ nhà chi trả. Tổng tiền thuê của khách không cộng khoản này; ví dụ chưa xét thuế, phí và tiền đặt cọc.

Nhập công thức tại hàng 2 của `NhaDat`, sau đó kéo xuống:

| Ô | Công thức | Ý nghĩa |
|---|---|---|
| F2 | `=ROUND(C2*E2,0)` | Tiền giảm mỗi tháng. |
| G2 | `=C2-F2` | Giá thuê mỗi tháng sau giảm. |
| H2 | `=G2*D2` | Tổng tiền thuê trong thời hạn đã nhập. |
| J2 | `=VLOOKUP(I2,HoaHong!$A$2:$C$4,3,FALSE)` | Tra đúng mã môi giới để lấy tỷ lệ hoa hồng. |
| K2 | `=ROUND(G2*J2,0)` | Hoa hồng môi giới trả một lần. |

`VLOOKUP` tìm mã trong cột đầu tiên của vùng tra, lấy kết quả ở cột thứ 3; `FALSE` yêu cầu khớp chính xác. Nếu mã không tồn tại, cần kiểm tra lỗi `#N/A`, không tự chuyển thành hoa hồng bằng 0. [Hướng dẫn VLOOKUP của Microsoft](https://support.microsoft.com/en-us/excel/functions/vlookup-function).

**Kết quả đối chiếu, đơn vị đồng:**

| Mã BĐS | Giảm/tháng | Thuê sau giảm/tháng | Tổng tiền thuê | Hoa hồng |
|---|---:|---:|---:|---:|
| BDS001 | 500.000 | 9.500.000 | 114.000.000 | 4.750.000 |
| BDS002 | 0 | 8.000.000 | 48.000.000 | 3.200.000 |
| BDS003 | 1.500.000 | 13.500.000 | 162.000.000 | 8.100.000 |
| BDS004 | 600.000 | 11.400.000 | 34.200.000 | 5.700.000 |

Không Merge Cell trong hai bảng; bật Filter để tra cứu. Dùng Data Validation giới hạn tỷ lệ chiết khấu từ 0 đến 1, số tháng là số nguyên dương và mã môi giới lấy từ danh sách có sẵn. Nếu Excel dùng dấu `;` để phân tách đối số, thay các dấu phẩy trong công thức bằng `;`.

### 2.3. Kịch bản Pitch Deck gọi vốn

Áp dụng **10/20/30**: tối đa **10 slide**, trình bày trong **20 phút**, chữ trên slide tối thiểu **30 pt**. Mỗi slide chỉ mang một thông điệp; phần giải thích để trong lời nói. [Quy tắc 10/20/30 của Guy Kawasaki](https://guykawasaki.com/the_102030_rule/).

| Slide | Thông điệp chính | Nội dung đưa lên slide | Thời gian |
|---|---|---|---:|
| 1 | Smart-Home giúp giao dịch nhà đất dễ theo dõi hơn. | Tên dự án và một câu giới thiệu giá trị. | 1 phút |
| 2 | Thông tin phân tán khiến các bên khó theo sát giao dịch. | Ba vấn đề cần kiểm chứng: tin thiếu thông tin, lịch sử giá khó theo dõi, phí môi giới chưa rõ. | 2 phút |
| 3 | Một nền tảng chung kết nối toàn bộ quy trình. | Sơ đồ: hồ sơ được duyệt → đấu giá/thuê → ghi nhận kết quả. | 2 phút |
| 4 | Người mua có thể theo dõi phiên đấu giá rõ ràng. | Một hình mô phỏng màn hình: giá hiện tại, giờ đóng, nút trả giá. | 2 phút |
| 5 | Thử nghiệm ở một khu vực giúp tập trung nguồn lực. | Nhóm khách hàng đầu tiên: chủ nhà và người mua/thuê tại một quận được chọn khảo sát. | 2 phút |
| 6 | Doanh thu dự kiến đến từ giao dịch hoàn tất. | Mô hình phí dịch vụ; cần thử nghiệm mức phí và khả năng chi trả trước khi chốt. | 2 phút |
| 7 | Giá trị khác biệt nằm ở thông tin có thể kiểm tra. | Ba điểm thiết kế: hồ sơ được duyệt, lịch sử trả giá, cách tính phí công khai. | 2 phút |
| 8 | Giai đoạn thử nghiệm phải chứng minh có nhu cầu. | Mục tiêu 3 tháng: 50 chủ nhà, 100 tin được duyệt, 20 giao dịch hoàn tất. | 2 phút |
| 9 | Nhóm triển khai theo từng mốc để kiểm soát tiến độ. | Vai trò cần có: sản phẩm, kỹ thuật, vận hành; lộ trình 12 tháng: MVP → thử nghiệm → mở rộng. | 2 phút |
| 10 | Khoản vốn đề xuất giúp hoàn thành kế hoạch 12 tháng. | Đề xuất 1 tỷ đồng: 50% sản phẩm, 30% vận hành, 20% tiếp cận khách hàng. | 3 phút |

**Tổng thời gian: 20 phút.** Các số ở slide 8 là mục tiêu, không phải thành tích đã đạt; mức vốn ở slide 10 là đề xuất của bài tập.

Mỗi slide dành khoảng **30% diện tích trống** làm nguyên tắc thiết kế, có lề thoáng và khoảng cách giữa các khối. White Space là vùng không chứa nội dung, không bắt buộc có màu trắng. Chỉ dùng một hình hoặc một cụm số liệu chính; tránh dán nguyên bảng Excel hay đoạn SRS lên slide.

## Phần 3. Kiểm thử

### 3.1. Sửa lỗi trôi vùng tra cứu bằng dấu `$`

Giả sử J2 đang dùng công thức chưa khóa vùng:

```excel
=VLOOKUP(I2,HoaHong!A2:C4,3,FALSE)
```

Khi kéo đến J5, Excel đổi công thức thành:

```excel
=VLOOKUP(I5,HoaHong!A5:C7,3,FALSE)
```

Vùng tra đã trượt khỏi danh sách môi giới. Vì vậy, mã `MG01` ở I5 không được tìm thấy dù vẫn có trong bảng gốc.

Sửa J2 thành công thức sau rồi kéo lại:

```excel
=VLOOKUP(I2,HoaHong!$A$2:$C$4,3,FALSE)
```

- `$A$2:$C$4` khóa cả cột và hàng, giữ nguyên vùng tra khi sao chép công thức.
- `I2` để tương đối, nên lần lượt đổi thành `I3`, `I4`, `I5` theo từng bất động sản.
- `$A2` chỉ khóa cột; `A$2` chỉ khóa hàng; `$A$2` khóa cả hai. Trong bài này dùng tham chiếu tuyệt đối cho toàn bộ vùng tra.

Dấu `$` giữ tham chiếu khi sao chép công thức, không phải chức năng bảo vệ dữ liệu khỏi chỉnh sửa. [Giải thích tham chiếu tương đối và tuyệt đối của Microsoft](https://support.microsoft.com/en-us/excel/get-started/overview-of-formulas-in-excel).

### 3.2. Các trường hợp kiểm tra

| Tình huống | Kết quả cần đạt |
|---|---|
| Kéo công thức từ J2 đến J5 | Vùng tra luôn là `HoaHong!$A$2:$C$4`; mã cần tìm đổi theo từng hàng. |
| Tra MG01 tại hàng 2 và hàng 5 | Cả hai trả về 50%; K2 là 4.750.000, K5 là 5.700.000 đồng. |
| Nhập mã không có trong bảng | Báo `#N/A` để sửa mã hoặc bổ sung bảng tra. |
| Thêm môi giới mới ở hàng 5 của `HoaHong` | Mở rộng vùng đã khóa đến hàng 5; dấu `$` không tự thêm bản ghi vào vùng tra. |
| Kiểm tra chiết khấu 0% và 100% | Giá sau giảm lần lượt bằng giá gốc và 0; không phát sinh giá âm. |
| Rà soát hồ sơ trước khi nộp | SRS đúng Heading 1/2/3; Excel đúng Field/Record, không Merge Cell; Pitch Deck đủ 10 slide, tổng 20 phút, chữ từ 30 pt và có khoảng trống. |

Các kết quả trên là số liệu đối chiếu và tiêu chí kiểm thử cho bộ hồ sơ, chưa phải biên bản chạy thử hệ thống thực tế.
