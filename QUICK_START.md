# Quick Start Guide - ICP-MS Batch Processor

## Installation (One-Time Setup)

### Step 1: Install Python
If you don't have Python installed:
- Download from: https://www.python.org/downloads/
- Install Python 3.7 or higher
- **Important**: Check "Add Python to PATH" during installation (Windows)

### Step 2: Download the Software
1. Go to: https://github.com/twinmum1277/ICP-MS_Pipeline_GUI
2. Click green **"Code"** button → **"Download ZIP"**
3. Unzip to a convenient location (e.g., Desktop or Documents)

### Step 3: Install Dependencies
Open Terminal (Mac) or Command Prompt (Windows):

**Mac:**
```bash
cd ~/Downloads/ICP-MS_Pipeline_GUI-main
pip3 install -r requirements.txt
```

**Windows:**
```bash
cd C:\Users\YourName\Downloads\ICP-MS_Pipeline_GUI-main
pip install -r requirements.txt
```

Wait for packages to install (~1-2 minutes).

---

## Running the Application

### Every Time You Use It:

1. **Open Terminal/Command Prompt**

2. **Navigate to the folder**:
   ```bash
   cd /path/to/ICP-MS_Pipeline_GUI
   ```

3. **Launch the GUI**:
   ```bash
   python3 icpms_gui.py    # Mac/Linux
   python icpms_gui.py     # Windows
   ```

4. **The GUI opens** - you're ready to process data!

---

## First-Time Usage

### Prepare Your Files:

**Template files are included in the download!** Look in the main folder for:
- `digest_template.csv`
- `icv_template.csv`  
- `sort_template.csv`
- `crm_values_template.csv`

Copy these templates to your working folder and replace the example data with your actual data.

---

You need 3-4 files for each analysis:

1. **SORT file** (required)
   - Export from MassHunter
   - Save as CSV
   - See `sort_template.csv` for example

2. **DIGEST file** (required)
   - Your dilution factors
   - Can be Excel or CSV
   - See `digest_template.csv` for format
   - **Key**: sample names must match SORT file

3. **ICV file** (required)
   - Your ICV target values
   - CSV format
   - See `icv_template.csv` for format

4. **REF file** (optional)
   - Reference material certified values
   - Your existing SRMs.csv works perfectly!

### Sample Naming (Important!):

In your SORT file, name samples:
- **Blanks**: `BLANK_1`, `BLANK_2`
- **ICV**: `ICV_1`, `ICV_2` (number them!)
- **Reference Materials**: `SRM_DOLT-5_1`, `SRM_DORM-4_1`
  - Format: `SRM_[REF_NAME]_[Number]`
  - REF_NAME must match your REF file
- **Regular Samples**: `Sample_01`, `Core_A_5cm`, etc. (anything else)

---

## Processing Workflow

### In the GUI:

1. **Set Working Folder**
   - Click "Set Folder…"
   - Navigate to your data folder
   - Click "Select"

2. **Select Files**
   - Click "Browse…" for SORT → select your MassHunter export
   - Click "Browse…" for DIGEST → select your dilution factors
   - Click "Browse…" for ICV → select your targets file
   - Click "Browse…" for REF → select your reference materials (if you have them)

3. **Choose Units**
   - Most users: Select **ppb (no conversion)**
   - Only select ppm if you specifically want ppm output

4. **Click "Run Batch"**
   - Wait for processing (usually 5-30 seconds)
   - Watch the status bar

5. **Review Summary Panel**
   - ✓ Green checkmark = Success!
   - 🔴 Red warnings = Check for missing samples
   - Review QC pass rates:
     - ICV should be >80% (ideally >90%)
     - REF should be >80%

6. **Open Output Excel**
   - Saved as `Batch_Results.xlsx` in SORT file's folder
   - Review "Corrected ppm (Wide)" sheet
   - Check QC Summary for any red cells (failed QC)

---

## Interpreting Results

### Processing Summary Panel:

**Batch Summary Section:**
- Shows counts of samples, ICV, REF, blanks, elements

**QC Pass Rates:**
- **Green numbers (≥80%)**: Good! Most elements passing QC
- **Orange numbers (60-80%)**: Marginal - check specific elements
- **Red numbers (<60%)**: Problem - investigate calibration or instrument

**Warnings:**
- **Yellow highlighted rows in Excel**: Missing DIGEST data
  - These samples were NOT corrected (used df=1.0)
  - Add them to DIGEST file and re-run

### QC Summary Sheet:

- **icv_pass = TRUE**: Element passed ICV check (90-110%)
- **ref_pass = TRUE**: Element passed REF check (80-120%)
- **Red cells**: Failed QC - consider excluding this element or re-analyzing

### Corrected ppm (Wide):

- **Red cells**: Below detection limit (< 3× blank SD)
- **Yellow rows**: Missing DIGEST correction
- **Normal cells**: Good data!

---

## Common Issues & Solutions

| Issue | Cause | Solution |
|-------|-------|----------|
| File picker doesn't open | GUI window too small | Resize window to see buttons |
| "Sample not found in DIGEST" | Name mismatch | Ensure names match (case-insensitive, spaces→underscores) |
| All df values = 1 | Sample names don't match | Check DIGEST sample names match SORT |
| REF recovery = NaN | REF name not found | Check sample name format: `SRM_DOLT-5_1` |
| ICV recovery = 0.1% | Wrong units | Uncheck "÷1000" option (use ppb) |
| Elements missing | Wrong column names | Check SORT headers have element format |

---

## Getting Help

1. **Check the Processing Summary** - it shows specific warnings
2. **Review TEMPLATE_GUIDE.md** - detailed format specifications
3. **Look at template files** - examples of correct formats
4. **Check GitHub Issues** - search for similar problems
5. **Email the developer** - include:
   - Screenshot of Processing Summary
   - First few rows of each input file
   - Error message if any

---

## Tips for Success

✓ **Use template files** as starting points  
✓ **Keep REF file updated** with all reference materials  
✓ **Number your QC samples** (ICV_1, ICV_2, not just "ICV")  
✓ **Check the summary panel** for warnings before using data  
✓ **Verify QC pass rates** are acceptable (>80%)  
✓ **Review yellow/red highlighting** in Excel output  

---

**That's it! You're ready to process ICP-MS data efficiently!** 🚀

For detailed information, see:
- `README.md` - Full project documentation
- `TEMPLATE_GUIDE.md` - Input file format details
- `INSTALLATION_GUIDE.md` - Complete installation instructions

