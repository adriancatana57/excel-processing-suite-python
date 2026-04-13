# 🔍 Advanced VLookup — Python Desktop App

> A powerful, streaming-capable VLOOKUP desktop application built with Python and CustomTkinter. Supports `.xlsx`, `.csv`, and `.txt` files, with multi-column import, repeat handling, and split-file output — all wrapped in a modern dark-mode GUI.

---

## ✨ Features

- **Multi-format support** — Works with `.xlsx`, `.csv`, and `.txt` as both source and reference files
- **Advanced key matching** — Join two files on any chosen key column, just like VLOOKUP in Excel — but without Excel's limitations
- **Multi-column import** — Bring in multiple columns from the reference file in a single operation
- **Repeat/duplicate handling** — Optionally collect all repeated values for a key into separate columns (e.g., `Col`, `Col_2`, `Col_3`)
- **Filter to matches only** — Output only rows that have a match in the reference file
- **Column selection** — Choose which columns from the main file to keep in the output
- **Split output** — Automatically split large results into multiple files (configurable line limit, multiples of 100,000)
- **Output formats** — Save results as `.csv`, `.txt`, or `.xlsx`
- **Live preview** — Preview the first 5 joined rows before committing to a full export
- **Clean export** — Export any loaded file as a clean, uncorrupted `.xlsx` (forces all values as text — no accidental number/date conversion)
- **Sheet selector** — For `.xlsx` files with multiple sheets, choose which sheet to use
- **File history** — Tracks today's generated files, with one-click folder navigation
- **Auto encoding & separator detection** — Detects file encoding and delimiter automatically
- **Dark mode UI** — Built with [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) for a modern look

---

## 🖥️ Screenshots

> *(Add screenshots here after first run)*
<img width="958" height="500" alt="Processing" src="https://github.com/user-attachments/assets/f76e199b-0bf3-4a1b-bb25-6ac7174e90b4" />
<img width="959" height="502" alt="Complete" src="https://github.com/user-attachments/assets/d9321e9b-4349-4a0e-af44-fe599cc9db4a" />

---

## 🚀 Getting Started

### Requirements

- Python 3.9+
- The following packages:

```bash
pip install customtkinter openpyxl xlsxwriter chardet
```

> `pandas` is optional — the app uses streaming I/O and does not require it for core functionality.

### Run

```bash
python vlookup_app.py
```

---

## 📖 How to Use

1. **Load the Main File** — the file you want to enrich (like the "lookup array" side)
2. **Load the Reference File** — the file that contains the data you want to bring in
3. **Choose key columns** — one from each file (the columns to match on)
4. **Select Columns to Import** — which columns from the reference file to bring into the output
5. *(Optional)* Configure repeats, match-only filter, column selection, output format, and split settings
6. Click **Preview** to preview the first 5 rows
7. Click **Process** to run the full export

---

## 📁 Project Structure

```
excel_suite.py       # Main application (single file)
assets/
  icon.ico           # App icon (optional, for Windows)
```

---

## 🛠️ Technical Notes

- **Streaming architecture** — Files are read row-by-row; the reference map is built in memory once. This means large files (millions of rows) can be processed without loading everything into RAM simultaneously.
- **All values written as text** — XlsxWriter is configured with `strings_to_numbers: False` and `strings_to_formulas: False` to prevent silent data corruption on codes, IDs, and leading-zero values.
- **Thread-safe UI** — Long operations run in a background thread with a busy dialog; the UI stays responsive.
- **DPI-aware on Windows** — Uses `SetProcessDpiAwareness(2)` for crisp rendering on high-DPI displays.

---

## 📄 License

This project is licensed under the **Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)** license.

You are free to:
- Share and adapt the code for **non-commercial purposes**, with attribution

You may **not**:
- Use this project or derivatives for commercial purposes without explicit written permission from the author

See [LICENSE](LICENSE) for full terms, or visit [creativecommons.org/licenses/by-nc/4.0](https://creativecommons.org/licenses/by-nc/4.0/).

---

## 👤 Author

Built by **Adrian CATANA** — feel free to connect on [LinkedIn](https://linkedin.com/in/adrian-c-catana).

---

## 🤝 Contributing

This is a personal project shared for portfolio purposes. Issues and suggestions are welcome via GitHub Issues.
