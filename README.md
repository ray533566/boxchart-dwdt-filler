# 📊 BoxChart dW/dT Filler

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://YOUR-APP-NAME.streamlit.app)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A Streamlit web app for optical transceiver QE engineers.  
Upload your **RawData** + **Format template** → get a filled Excel with **dW/dT avg / max / min** per channel and SN, with all **box-plot charts intact**.

---

## ✨ Features

- **Drag-and-drop upload** — no local Python install needed
- **Two sheets filled automatically**

  | Sheet | Channels used | Column A | Columns B–M | Columns P–AA |
  |---|---|---|---|---|
  | `Normal_Operational Current` | `1_Operational` → `4_Operational` | TESTSN | avg / max / min × CH1–4 | max / avg / min × CH1–4 |
  | `Bias400_Maximum Current` | `1_Maximum` → `4_Maximum` | TESTSN | avg / max / min × CH1–4 | max / avg / min × CH1–4 |

- **Box-plot charts preserved** — sheet XML is patched directly inside the ZIP; openpyxl is never used for saving (it strips `chartEx` objects)
- **Live preview** — stat cards + channel distribution table before you run
- **CLI mode** — run the same logic from the command line without Streamlit

---

## 🚀 Quick Start (local)

```bash
# 1. Clone
git clone https://github.com/ray533566/boxchart-dwdt-filler.git
cd boxchart-dwdt-filler

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run Streamlit app
streamlit run app.py

# 4. Or run CLI script directly
python fill_boxchart.py \
  --raw  RawData_for_BoxPlot.xlsx \
  --fmt  Format_Mode_hopping_BoxChart620.xlsx \
  --out  Output_filled.xlsx
```

---

## 📁 Repository Structure

```
boxchart-dwdt-filler/
│
├── app.py                  # Streamlit web UI
├── fill_boxchart.py        # Core logic (also CLI entry point)
├── requirements.txt        # Python dependencies
│
├── .streamlit/
│   └── config.toml         # Dark theme configuration
│
├── sample_data/            # (optional) small anonymized samples
│   ├── RawData_sample.xlsx
│   └── Format_sample.xlsx
│
└── README.md
```

---

## 🔧 How It Works

### Input format — RawData

| Column | Description |
|--------|-------------|
| `TESTSN` | Serial number (e.g. `P204261000100`) |
| `CHNumber` | `1_Operational`, `2_Operational`, `3_Operational`, `4_Operational`, `1_Maximum` … `4_Maximum` |
| `dW/dT` | Wavelength-vs-temperature slope value (nm/°C) |

### Aggregation logic

For each unique `TESTSN` × `CHNumber` combination:

```
avg  =  AVERAGE(dW/dT)   →  rounded to 4 decimal places
max  =  MAX(dW/dT)
min  =  MIN(dW/dT)
```

### Column mapping

```
B  = avg CH1    C  = max CH1    D  = min CH1
E  = avg CH2    F  = max CH2    G  = min CH2
H  = avg CH3    I  = max CH3    J  = min CH3
K  = avg CH4    L  = max CH4    M  = min CH4

P  = max CH1    Q  = avg CH1    R  = min CH1   (reordered for chart source)
S  = max CH2    T  = avg CH2    U  = min CH2
V  = max CH3    W  = avg CH3    X  = min CH3
Y  = max CH4    Z  = avg CH4    AA = min CH4
```

### Why direct ZIP patching?

Excel box plots use the `chartEx` format, which openpyxl silently drops when it saves.  
This tool reads the original `.xlsx` as a ZIP archive, replaces only three files
(`sheet2.xml`, `sheet3.xml`, `sharedStrings.xml`), and writes a new ZIP — keeping all
chart relationships untouched.

---

## ☁️ Deploy to Streamlit Community Cloud

1. **Fork** this repository to your GitHub account  
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub  
3. Click **New app**  
4. Select your fork → branch `main` → main file `app.py`  
5. Click **Deploy** — your app will be live at `https://your-app-name.streamlit.app`  
6. Update the badge URL at the top of this README with your real app URL

---

## 🖥 CLI Reference

```
python fill_boxchart.py [OPTIONS]

Options:
  --raw   PATH    Path to RawData_for_BoxPlot.xlsx       (default: RawData_for_BoxPlot.xlsx)
  --fmt   PATH    Path to Format_Mode_hopping_BoxChart620.xlsx  (default: Format_Mode_hopping_BoxChart620.xlsx)
  --out   PATH    Output file path                        (default: Format_Mode_hopping_BoxChart620_updated.xlsx)
```

The CLI prints a verification summary and chart-file count after saving.

---

## 📦 Dependencies

| Package | Purpose |
|---------|---------|
| `pandas` | Read Excel, aggregate dW/dT |
| `lxml` | Parse/build sharedStrings XML |
| `openpyxl` | Verification read-back only |
| `streamlit` | Web UI |

---

## 🙋 Author

**Ray** — Quality Engineer, Optical Transceiver Manufacturing  
GitHub: [@ray533566](https://github.com/ray533566)

---

## 📄 License

MIT — free to use and modify.
