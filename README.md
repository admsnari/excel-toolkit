# Excel GUI App

A lightweight and user-friendly Python desktop application for processing Excel files. This GUI app allows users to select and process Excel files with specific brand filters using a modern themed interface built with `ttkbootstrap`.

## 🚀 Features

- Open and process `.xlsx` files
- Filters Excel data by brand patterns like: `SLBS`, `2L1T`, `3L1T`, `4L1T`, `2L2T`, `3L`, `4L`
- Modern GUI using [`ttkbootstrap`](https://ttkbootstrap.readthedocs.io)
- One-click Excel file selection
- Suitable for non-technical users
- Packaged as a `.exe` for Windows 11 (with `pyinstaller`)

## 🖼 GUI Preview

*(Add screenshot here if available)*

---

## 🧰 Requirements

- Python 3.10+
- [ttkbootstrap](https://pypi.org/project/ttkbootstrap/)
- [openpyxl](https://pypi.org/project/openpyxl/) or [pandas](https://pypi.org/project/pandas/) (whichever you used for Excel handling)

To install dependencies:

```bash
pip install ttkbootstrap pandas openpyxl
python -m PyInstaller --noconfirm --onefile --windowed excel_gui_app.py  # build command, run on windows
```

