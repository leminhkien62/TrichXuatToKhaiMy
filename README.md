# Trích Xuất Tờ Khai Hải Quan Mỹ

Tool dùng Python để trích xuất dữ liệu từ **CBP Form 7501 - Entry Summary** và xuất kết quả ra Excel.

## 1. Yêu cầu môi trường

* Windows 10/11
* Python 3.12+
* Git

Kiểm tra Python:

```powershell
python --version
```

Ví dụ:

```text
Python 3.12.x
```

---

## 2. Clone project

Clone source code từ GitHub:

```powershell
git clone https://github.com/leminhkien62/TrichXuatToKhaiMy.git
```

Đi vào thư mục project:

```powershell
cd TrichXuatToKhaiMy
```

---

## 3. Tạo môi trường `.venv`

Tạo virtual environment:

```powershell
python -m venv .venv
```

Sau khi tạo thành công, thư mục sẽ có:

```text
TrichXuatToKhaiMy/
├── .venv/
├── parse_7501.py
├── requirements.txt
└── README.md
```

---

## 4. Kích hoạt `.venv`

Trên PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Nếu thành công, đầu dòng terminal sẽ xuất hiện:

```text
(.venv) PS C:\...\TrichXuatToKhaiMy>
```

### Nếu PowerShell chặn chạy script

Chạy:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Sau đó:

```powershell
.\.venv\Scripts\Activate.ps1
```

---

## 5. Cài thư viện

Sau khi đã activate `.venv`:

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Các thư viện chính:

```text
PyMuPDF
pandas
openpyxl
```

`tkinter` là thư viện có sẵn trong Python Windows nên không cần cài bằng `pip`.

---

## 6. Kiểm tra thư viện

Chạy:

```powershell
python -c "import pymupdf, pandas, openpyxl, tkinter; print('Environment OK')"
```

Nếu xuất hiện:

```text
Environment OK
```

thì môi trường đã sẵn sàng.

---

## 7. Chạy chương trình

Chạy:

```powershell
python parse_7501.py
```

Chương trình sẽ mở giao diện để:

1. Chọn thư mục chứa file PDF.
2. Chọn thư mục lưu kết quả.
3. Thực hiện trích xuất dữ liệu.
4. Xuất file Excel kết quả.

---

## 8. Đóng gói thành `.exe`

Nếu cần tạo file `.exe`, trước tiên cài PyInstaller:

```powershell
pip install pyinstaller
```

Kiểm tra:

```powershell
pyinstaller --version
```

Sau đó build:

```powershell
pyinstaller --onefile --windowed --name "TrichXuatToKhaiMy" parse_7501.py
```

File `.exe` sẽ nằm tại:

```text
dist/
└── TrichXuatToKhaiMy.exe
```

Có thể chạy trực tiếp:

```powershell
.\dist\TrichXuatToKhaiMy.exe
```

---

## 9. Khi clone project lần sau

Không cần cài lại Python.

Chỉ cần:

```powershell
git clone https://github.com/leminhkien62/TrichXuatToKhaiMy.git
cd TrichXuatToKhaiMy
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python parse_7501.py
```

---

## 10. Cập nhật code mới từ GitHub

Nếu project đã được clone trước đó:

```powershell
cd TrichXuatToKhaiMy
.\.venv\Scripts\Activate.ps1
git pull
```

Sau khi cập nhật thư viện:

```powershell
pip install -r requirements.txt
```

Sau đó chạy:

```powershell
python parse_7501.py
```

---

## 11. Cập nhật code lên GitHub

Sau khi sửa code:

```powershell
git add .
git commit -m "Update parse 7501"
git push
```

---

## 12. Lưu ý

Không đưa `.venv` lên GitHub.

Trong `.gitignore` nên có:

```gitignore
.venv/
__pycache__/
build/
dist/
*.pyc
*.spec
```

Không đưa các file dữ liệu thực tế như PDF tờ khai hoặc file Excel chứa dữ liệu nội bộ lên repository nếu không cần thiết.
