# 📡 RadioTT Analyzer

**ComparateurRadioTT** is a modern desktop application built with **Python** and **PyQt5** that allows Tunisie Telecom engineers to analyze and compare azimuth and geographic coordinate data of radio network sites from Excel files.  
It supports both **azimuth alignment verification** and **geographical coordinate matching**, with a clean and user-friendly interface respecting UI/UX design principles.

---

## 🚀 Features

- ✅ Compare **Azimuths (2G, 3G, 4G)** for consistency.
- ✅ Compare **Coordinates (longitude, latitude)** across entries.
- ✅ Highlights identical or differing values with **color-coded cells** in Excel.
- ✅ Intuitive **PyQt5 interface** with radio buttons and progress bar.

---


## 🛠️ Installation

### Requirements

  - Python 3.8+
  - Required libraries:
  ```bash
  pip install openpyxl PyQt5
  ```
### Executable file
There is a file **comarateurRadioTT.exe** that allows you to run the app directly without installing the required librairies (PyQt5 and openpyxl)
All you need is to clone this reposotory
```
git clone https://github.com/ayaarbi/ComparateurRadioTT.git

```
and run **comarateurRadioTT.exe**

---

## 🧠 How It Works
1. Launch the app.
2. Choose between:
  - "Azimuth Comparison"
  - "Coordinates Comparison"
3. Select an Excel file (.xlsx) with radio site data.
4. The app processes the data and highlights:
  - Identical values (green)
  - Differences or missing data (red)
5. A result file is saved automatically.

---
## 📄 Excel File Format
The app expects an Excel file with two header rows respecting all specifities in the code, The excel file that we wrok with could not be shared here because it a confidential file.
A PDF repport is added to this project to explain more the context although it was written in French.



  
