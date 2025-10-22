# xlsx-trimmer

Công cụ dòng lệnh giúp giảm kích thước file Excel (`.xlsx`) bằng cách loại bỏ các hàng, cột và dữ liệu không cần thiết.

## Vấn đề

Các file Excel, đặc biệt là những file được tạo hoặc chỉnh sửa bởi các hệ thống tự động, thường có kích thước phình to một cách bất thường. Điều này xảy ra do file chứa hàng triệu hàng hoặc hàng nghìn cột trống không được sử dụng, nhưng Excel vẫn lưu thông tin về chúng. Những file này có thể trở nên rất chậm khi mở và chiếm dụng nhiều dung lượng lưu trữ.

`xlsx-trimmer` được tạo ra để giải quyết vấn đề này bằng cách "cắt tỉa" file `.xlsx`, chỉ giữ lại vùng dữ liệu thực sự được sử dụng.

## Cách hoạt động

File `.xlsx` thực chất là một file ZIP chứa các file XML. Công cụ này hoạt động theo các bước sau:
1.  Giải nén file `.xlsx` vào một thư mục tạm.
2.  Đọc và phân tích các file XML của từng worksheet để xác định vùng dữ liệu đã sử dụng (dựa trên ô cuối cùng có chứa giá trị).
3.  Ghi lại các file XML của worksheet, loại bỏ tất cả các hàng và cột nằm ngoài vùng dữ liệu đã sử dụng.
4.  Xóa bỏ một số thành phần có thể gây phình to file như `conditionalFormatting`, `dataValidations`, `calcChain.xml`, v.v.
5.  Nén lại các file đã được xử lý thành một file `.xlsx` mới với kích thước nhỏ hơn đáng kể.

## Cách sử dụng

### Cú pháp

```shell
xlsx-trimmer <đường-dẫn-tới-file-hoặc-thư-mục> [tùy-chọn]
```

### Tham số

-   `<đường-dẫn-tới-file-hoặc-thư-mục>`: (Bắt buộc) Đường dẫn đến một file `.xlsx` duy nhất hoặc một thư mục chứa các file `.xlsx` cần xử lý.

### Tùy chọn

-   `-o, --output-dir <thư-mục-đầu-ra>`: Chỉ định thư mục để lưu các file đã được xử lý. Nếu không cung cấp, file mới sẽ được lưu cùng thư mục với file gốc.
-   `--threshold-mb <số-MB>`: Chỉ xử lý các file có kích thước lớn hơn ngưỡng megabyte được chỉ định. Mặc định là `10`.
-   `--suffix <hậu-tố>`: Hậu tố được thêm vào tên file đầu ra. Mặc định là `_trimmed`. Ví dụ: `BaoCao.xlsx` sẽ trở thành `BaoCao_trimmed.xlsx`.

### Ví dụ

1.  **Xử lý một file duy nhất:**
    ```shell
    xlsx-trimmer "D:\BaoCaoThang\report.xlsx"
    ```
    Thao tác này sẽ tạo ra file `report_trimmed.xlsx` trong cùng thư mục nếu kích thước của nó lớn hơn 10MB.

2.  **Xử lý tất cả file trong một thư mục và lưu vào nơi khác:**
    ```shell
    xlsx-trimmer "D:\BaoCaoThang" -o "D:\BaoCaoDaXuLy"
    ```

3.  **Xử lý file lớn hơn 50MB với hậu tố khác:**
    ```shell
    xlsx-trimmer "D:\Data" --threshold-mb 50 --suffix _fixed
    ```

## Biên dịch từ mã nguồn

Để biên dịch chương trình, bạn cần cài đặt [Rust](https://www.rust-lang.org/tools/install).

Sau đó, chạy lệnh sau trong thư mục gốc của dự án:
```shell
cargo build --release
```
File thực thi sẽ được tạo tại `target/release/xlsx-trimmer.exe`.
