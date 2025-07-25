# IEC 61850 Signal Extractor

This tool extracts signal names from IEC 61850-compliant XML (SCL) files and exports them to a professional-looking Excel file.

## ✅ Features

- Clean, styled Excel output grouped by dataset (`DS1`, `DS2`, etc.)
- GUI-based file selector — no command line required
- Fully portable `.exe` with no Python installation needed
- Handles complex SCL structures (supports >100 signals)

## 🚀 How to Use

### 📦 Recommended: Download Executable

1. Go to the [Releases](https://github.com/eledebar/IEC61850-SignalExtractor/releases) section
2. Download the latest `.exe` file
3. Double-click it
4. Select your `.xml` SCL file when prompted
5. The `.xlsx` output will be generated in the same folder

### 🧑‍💻 Alternative: Run from Python (for developers)

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the script:
   ```bash
   python generate_signals_gui.py
   ```

## 📝 Output Format

The resulting Excel file will:

- Create one column per dataset (`DS1`, `DS2`, etc.)
- List all fully constructed signal names based on:
  - `lnClass`
  - `lnInst`
  - `doName`
  - `daName`
  - `fc` → translated to `.mag.f`, `.stVal`, `.ctlVal`, etc.

## 📄 Example XML

A fully **fictitious** and complex example SCL file is provided in the `example/` folder as:

```
example_demo.xml
```

This file is not based on any real project and is safe to use for testing.

## 📜 License

MIT License — see [LICENSE](LICENSE)
