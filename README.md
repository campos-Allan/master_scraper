# 📦 Master Scraper

Automate the daily consolidation of logistics data from multiple PDF and Excel files into a unified, ready-to-share spreadsheet. Designed to handle inconsistent file formats and minimize manual work, `master_scraper` streamlines the process of tracking product storage and transportation across several locations.

---

## 🚀 Objective

Eliminate manual copy-paste and error-prone data entry by automating the extraction, transformation, and consolidation of daily logistics data from diverse sources.

---

## ⚙️ Features

- 📥 Reads and processes 5 PDF and 5 Excel files per day
- 🏷️ Handles multiple products (X1, Y1, Y2) and transportation types (Modal A & B)
- 🗂️ Organizes data by storage location and trip type
- 📝 Outputs a model Excel file (`molde.xlsx`) ready for direct copy into your master spreadsheet
- 🗑️ Moves processed files to a trash folder to prevent duplication
- 🔄 Easily adaptable to changes in file formats

---

## 📁 Project Structure

```
master_scraper-main/
├── files/                  # Example input files (PDFs, Excels)
├── files.py                # File path and variable definitions
├── main.py                 # Main script to run the workflow
├── molde.xlsx              # Model output Excel file
├── R1_S1_insert_reader.py  # Reader for R1 and S1 insert files
├── S1_insert_reader.py     # Reader for S1 insert files
├── T1_insert_reader.py     # Reader for T1 insert files
├── T1_reader.py            # Reader for T1 files
├── trash/                  # Folder for processed files
└── var.py                  # Variable definitions
```

---

## 🧰 Requirements

- Python 3.8+
- Windows OS
- Excel (for output viewing)
- Java

**Python packages:**
- pandas
- openpyxl

Install with:

```sh
pip install pandas openpyxl
```

---

## 🧠 How It Works

1. Place your daily PDF and Excel files in the `files/` directory.
2. Run the tool via the GUI (`GUI.py`) or directly (`main.py`).
3. The script:
    - Loads file paths and variables from `files.py` and `var.py`
    - Reads and processes each file using specialized readers
    - Consolidates and formats the data into `molde.xlsx`
    - Moves processed files to `trash/` to avoid reprocessing
4. Open `molde.xlsx` and copy the organized data into your shared master spreadsheet.

---

## 🗺️ Workflow & Scheme

Below is a schematic of the product transportation and storage process automated by this tool:

![Scheme](https://i.imgur.com/iiXjSai.png)

---

## ✅ Usage

1. **Prepare your input files:**  
   Place all new PDF and Excel files in the `files/` folder.

2. **Run the GUI (recommended):**
   ```sh
   python GUI.py
   ```
   - Click the button to process files.

   **Or run directly:**
   ```sh
   python main.py
   ```

3. **Copy results:**  
   Open `molde.xlsx` and transfer the data to your master spreadsheet.

---

## ⚠️ Notes

- Example files are provided in `files/` for demonstration. Some code sections use mock data due to the inability to share original files.
- Processed files are automatically moved to `trash/` to prevent duplicate entries.

---

## 📊 Results

- The output Excel file highlights new and updated data for each day and trip type, using color coding for easy identification.

![Results](https://i.imgur.com/9tgvQ5d.png)

Each storage spot has a sheet with all the storage info extracted from the files and divided per day, now it's only a matter of copying the info into the shared master spreadsheet.

![Results2](https://i.imgur.com/JSGzAlR.png)

---

## 📈 Benefits

- Saves hours of manual work and reduces errors in daily logistics reporting
- Ensures consistent, organized, and up-to-date data
- Empowers non-technical users with a simple interface

---

## 👨‍💻 Author

Created for internal automation of logistics data consolidation and reporting.  
Adapt and use responsibly!
