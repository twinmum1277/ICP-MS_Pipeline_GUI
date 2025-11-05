# ICP-MS Batch Processor - Installation & Usage Guide

## Overview

The ICP-MS Batch Processor consists of **two Python files** that work together:

1. **`icpms_gui.py`** - The graphical user interface (GUI)
   - Provides the visual interface with buttons and file selection
   - Captures user inputs and settings
   - Displays processing summary and warnings
   - When you click "Run Batch", it calls functions from the second file

2. **`process_icpms_batch.py`** - The processing engine
   - Contains all the data processing functions
   - Loads and parses files
   - Performs calculations (blank correction, dilution factors, QC)
   - Generates Excel output
   - Imported by `icpms_gui.py` (line 18: `from process_icpms_batch import process_batch`)

**How they work together:**
- The GUI file imports the processing engine
- When you click "Run Batch", the GUI calls `process_batch()` from the engine
- The engine does all the heavy lifting and returns results
- The GUI displays the results in the summary panel

---

## Installation Instructions

### For Users (Simple Installation)

#### Prerequisites
- **Python 3.7 or higher** installed on your computer
  - Check: Open Terminal (Mac) or Command Prompt (Windows) and type: `python3 --version`
  - If not installed, download from: https://www.python.org/downloads/

#### Step 1: Get the Files

**Option A - Download from GitHub:**
1. Go to: https://github.com/twinmum1277/ICP-MS_Pipeline_GUI
2. Click the green **"Code"** button
3. Click **"Download ZIP"**
4. Unzip the downloaded file
5. You'll get a folder called `ICP-MS_Pipeline_GUI-main`

**Option B - Clone with Git (if you have Git installed):**
```bash
git clone https://github.com/twinmum1277/ICP-MS_Pipeline_GUI.git
cd ICP-MS_Pipeline_GUI
```

#### Step 2: Install Required Packages

Open Terminal (Mac) or Command Prompt (Windows), navigate to the folder, and run:

**Mac/Linux:**
```bash
cd /path/to/ICP-MS_Pipeline_GUI
pip3 install -r requirements.txt
```

**Windows:**
```bash
cd C:\path\to\ICP-MS_Pipeline_GUI
pip install -r requirements.txt
```

This installs the required packages:
- `pandas` - Data manipulation
- `numpy` - Numerical operations
- `openpyxl` - Excel file reading
- `xlsxwriter` - Excel file writing with formatting

#### Step 3: Verify Installation

Test that everything is installed correctly:

```bash
python3 -c "import pandas, numpy, openpyxl, xlsxwriter; print('✓ All packages installed successfully!')"
```

If you see the success message, you're ready to go!

---

## Running the Application

### Quick Start

1. **Navigate to the folder** containing the Python files:
   ```bash
   cd /path/to/ICP-MS_Pipeline_GUI
   ```

2. **Launch the GUI**:
   ```bash
   python3 icpms_gui.py
   ```
   
   On Windows:
   ```bash
   python icpms_gui.py
   ```

3. **The GUI window will open** - follow the on-screen instructions!

### Using the GUI

#### Step-by-Step Workflow:

1. **Set Working Folder (Optional but Recommended)**
   - Click "Set Folder…" at the top
   - Select the folder containing your data files
   - This makes file selection faster

2. **Select Input Files**
   - Click "Browse…" for each required file:
     - **SORT file**: MassHunter export (CSV)
     - **DIGEST file**: Dilution factors (CSV or Excel)
     - **ICV file**: Calibration targets (CSV)
     - **REF file** (optional): Reference material values (CSV)

3. **Choose Output Units**
   - Select **ppb** (no conversion) - default
   - Or **ppm** (÷1000) if you want ppm output

4. **Click "Run Batch"**
   - Processing happens automatically
   - Progress shown in status bar
   - Results appear in Processing Summary panel

5. **Review Results**
   - Check the Processing Summary panel for:
     - Sample counts
     - QC pass rates
     - Any warnings about missing data
   - Output Excel file saved in same folder as SORT file

---

## Organizing Your Workflow with Templates

### Recommended Folder Structure:

We recommend organizing your files like this:

```
Your_Analysis_Folder/
├── SORT/
│   ├── 2025-11-05_batch1.csv       (MassHunter export)
│   ├── 2025-11-12_batch2.csv
│   └── ...
├── DIGEST/
│   ├── 2025-11-05_dilutions.xlsx   (Your dilution factors)
│   ├── 2025-11-12_dilutions.xlsx
│   └── ...
├── ICV/
│   └── icv_targets.csv             (Reuse same file usually)
├── REF/
│   └── SRMs_master.csv             (Master lookup - grows over time)
└── OUTPUT/
    ├── Batch_Results_2025-11-05.xlsx
    ├── Batch_Results_2025-11-12.xlsx
    └── ...
```

**Or simpler - all in one folder per batch:**

```
Batch_2025-11-05/
├── sort.csv              (MassHunter export)
├── digest.xlsx           (Dilution factors)
├── icv.csv               (ICV targets - copy from master)
├── SRMs.csv              (REF values - copy from master)
└── Batch_Results.xlsx    (Output - auto-generated)
```

### Using the Template Files:

The software includes template files to help you get started:

#### 1. **For Your First Run:**
1. Copy the template files to your working folder:
   - `digest_template.csv` → rename to `your_digest.csv`
   - `icv_template.csv` → rename to `your_icv.csv`
   - `sort_template.csv` → just for reference (MassHunter creates your actual SORT file)

2. Open each template in Excel/Numbers/Google Sheets

3. **Replace template data with your actual data**:
   - Keep the column headers exactly as shown
   - Replace the example values with your real values
   - Save

#### 2. **Building Your Master Files:**

**ICV File** (icv.csv or icv_targets.csv):
- Create once, reuse for all analyses (unless targets change)
- Include all elements you typically analyze
- Update only when you change your ICV standard

**REF File** (SRMs.csv or ref_master.csv):
- Start with one reference material
- **Add new rows** each time you use a new reference material
- Grows over time into a master lookup table
- Keep one master copy and use for all analyses

**DIGEST File**:
- Create NEW for each batch (dilution factors change)
- Can start from previous batch and modify
- Or use template and fill in your sample names

#### 3. **Template Workflow Example:**

**First time:**
```bash
# Copy templates to your folder
cp digest_template.csv my_first_analysis/digest.csv
cp icv_template.csv my_first_analysis/icv.csv
# Edit these files with your data
# Export SORT from MassHunter to my_first_analysis/sort.csv
# Run GUI, select files, process!
```

**Subsequent analyses:**
```bash
# Reuse ICV and REF files
cp my_first_analysis/icv.csv my_second_analysis/icv.csv
cp my_first_analysis/SRMs.csv my_second_analysis/SRMs.csv
# Create new DIGEST for new samples
# Export new SORT from MassHunter
# Run GUI and process!
```

### What to Keep vs. Recreate:

| File | Strategy | Notes |
|------|----------|-------|
| **SORT** | New each batch | Created by MassHunter |
| **DIGEST** | New each batch | Sample-specific dilution factors |
| **ICV** | Reuse | Only update if ICV standard changes |
| **REF** | Reuse & grow | Add new reference materials as needed |

---

## File Requirements

### Required Files:

#### 1. SORT File (MassHunter Export)
- **Format**: CSV
- **Content**: Raw instrument data
- **Headers**: 
  - `Acq. Date-Time` or similar (date/time column)
  - `Sample Name` or similar (sample identifier)
  - Element columns like: `63  Cu  [ He ]`, `75 -> 91  As  [ O2 ]`
- **Sample naming**:
  - Blanks: Include "BLANK" (e.g., `BLANK_1`, `BLANK_2`)
  - ICV: Include "ICV" (e.g., `ICV_1`, `ICV_2`)
  - Reference materials: `SRM_[NAME]_[#]` (e.g., `SRM_DOLT-5_1`)
  - Regular samples: Any other name

#### 2. DIGEST File (Dilution Factors)
- **Format**: CSV or Excel (.xlsx, .xls)
- **Required columns**:
  - `sample_id`: Sample name (must match SORT file)
  - `df`: Dilution factor (number)
- **Notes**: 
  - Sample names are auto-normalized (uppercase, spaces→underscores)
  - QC samples (ICV, BLANK, SRM_*) don't need DIGEST entries (default df=1)

#### 3. ICV File (Calibration Targets)
- **Format**: CSV
- **Required columns**:
  - `element`: Element symbol (Ag, Al, As, etc.)
  - `icv_target`: Target concentration for ICV samples (in same units as raw data)
- **Optional column**:
  - `ref_target`: Generic reference material targets (usually omit if using REF file)

#### 4. REF File (Reference Material Values) - Optional
- **Format**: CSV (wide or long format)
- **Wide format** (as shown in your SRMs.csv):
  - Row 1: Element symbols (Ag, Al, As...)
  - Row 2: Element full names (Silver, Aluminum...)
  - Row 3: Units (mg/kg)
  - Row 4+: Reference material data (column 2 = REF name, rest = values)
- **Long format**:
  - Columns: `ref_name`, `element`, `target_value`
- **Notes**:
  - Values assumed to be in mg/kg (ppm) - auto-converted to ppb
  - REF name must match what's extracted from sample names

---

## Sample Naming Conventions

**CRITICAL**: Sample names must follow these patterns for proper QC processing:

### QC Samples:
- **Blanks**: `BLANK_1`, `BLANK_2`, `BLANK_3`
- **ICV**: `ICV_1`, `ICV_2` (run before/after batch)
- **Reference Materials**: `SRM_[REF_NAME]_[Replicate#]`
  - Examples: `SRM_DOLT-5_1`, `SRM_DOLT-5_2`, `SRM_DORM-4_1`
  - The REF name (e.g., "DOLT-5") must match entries in your REF file

### Regular Samples:
- Any naming that doesn't include BLANK, ICV, SRM_, or DUP
- Examples: `LIAN_32`, `Sample_001`, `Core_A_5cm`

---

## Understanding the Output

The application generates an Excel file (`Batch_Results.xlsx`) with 4 sheets:

### 1. Corrected ppm (Wide)
- **Format**: Rows = samples, Columns = elements
- **Color coding**:
  - 🔴 **Red cells**: Below detection limit
  - 🟡 **Yellow rows**: Missing DIGEST data (df=1.0, not corrected!)
- **Values**: Corrected concentrations with 3 decimal places

### 2. Corrected ppm (Long)
- **Format**: Long table with all data points
- **Columns**: sample_id, element, channel_id, corrected, etc.
- **Use**: For detailed analysis or further processing

### 3. QC Summary
- **Format**: One row per element
- **Shows**:
  - Selected channel for each element
  - ICV recovery % (target: 90-110%)
  - REF recovery % (target: 80-120%)
  - Pass/fail flags
- **Color coding**: 🔴 **Red**: Failed QC (outside target range)

### 4. Below Detection Limit
- **Lists**: All measurements below MDL (3× blank SD)
- **Columns**: sample_id, element, channel_id, raw_conc, avg_blank, mdl

---

## Troubleshooting

### "File picker not opening"
- **Mac**: Resize the GUI window to ensure Browse buttons are visible
- **Solution**: Clear Python cache: `rm -rf __pycache__` then relaunch

### "Processing Summary shows warnings"
- **Yellow highlighted samples**: Missing from DIGEST file
  - These used df=1.0 (no dilution correction)
  - Add them to your DIGEST file with correct dilution factors
  - Re-run the analysis

### "REF recovery shows NaN"
- **Cause**: REF samples not found or not matching REF file
- **Check**:
  - Sample names start with `SRM_` (e.g., `SRM_DOLT-5_1`)
  - REF name matches exactly in REF file (e.g., "DOLT-5")
  - REF file was uploaded

### "ICV/REF recovery percentages are wrong"
- **Check units**:
  - If raw data is in ppb and ICV targets are in ppb → Select **ppb** output
  - If raw data is in ppb and you want ppm output → Select **ppm** output
  - REF file values are automatically converted from mg/kg to ppb

### "Column headers not recognized"
- **SORT file**: Must have `Sample Name` or similar in column header
- **DIGEST file**: Must have `sample_id` and `df` columns
- **ICV file**: Must have `element` and `icv_target` columns

### "Elements missing from output"
- Verify element symbols in ICV/REF files match SORT file
- Check that element columns in SORT have format: `63  Cu  [ He ]`

---

## File Structure

When sharing with others, include these files:

### Essential Files (Required):
- `icpms_gui.py` - GUI application
- `process_icpms_batch.py` - Processing engine
- `requirements.txt` - Package dependencies
- `README.md` - Project documentation

### Template Files (Helpful):
- `digest_template.csv` - Example DIGEST format
- `icv_template.csv` - Example ICV format
- `sort_template.csv` - Example SORT format
- `crm_values_template.csv` - Example REF format (old name, shows long format)

### Documentation (Helpful):
- `TEMPLATE_GUIDE.md` - Detailed format specifications
- `INSTALLATION_GUIDE.md` - This file
- `LICENSE` - MIT License

---

## Sharing with Colleagues

### Simple Method:
1. **Zip the entire folder** including all Python files and templates
2. Share the ZIP file
3. Recipient follows Installation Instructions above

### Using Git/GitHub:
1. Share the repository URL: https://github.com/twinmum1277/ICP-MS_Pipeline_GUI
2. Recipient clones it: `git clone https://github.com/twinmum1277/ICP-MS_Pipeline_GUI.git`
3. They follow Installation Instructions

---

## Advanced: Running from Command Line

The processor can also run without the GUI:

```bash
python3 process_icpms_batch.py \
  --sort path/to/sort.csv \
  --digest path/to/digest.csv \
  --icv path/to/icv.csv \
  --ref path/to/ref.csv \
  --out path/to/output.xlsx
```

Options:
- `--no-div1000`: Don't divide by 1000 (keep in ppb)

---

## Tips for Best Results

### Data Preparation:
1. **Clean your SORT file**:
   - Remove extra header rows if needed
   - Ensure proper column names
   - Number your ICV samples: `ICV_1`, `ICV_2`, etc.

2. **Build a master REF file**:
   - Include all reference materials you use
   - Wide format (like your SRMs.csv) is fine
   - The code auto-converts it

3. **Keep DIGEST file updated**:
   - Include all sample names from each run
   - ICV, BLANK, SRM samples can be omitted (default df=1)

### Workflow:
1. Export data from MassHunter → SORT file
2. Prepare DIGEST file with dilution factors
3. Use existing ICV and REF files (or update as needed)
4. Run the GUI
5. Check Processing Summary for warnings
6. Review QC pass rates
7. Open Excel output and verify results

---

## Support

For issues or questions:
- Check the Processing Summary panel for specific error messages
- Review this guide and TEMPLATE_GUIDE.md
- Check GitHub Issues: https://github.com/twinmum1277/ICP-MS_Pipeline_GUI/issues

---

## System Requirements

- **Operating System**: Mac, Windows, or Linux
- **Python**: 3.7 or higher
- **Disk Space**: ~50 MB for software + space for data files
- **RAM**: 2 GB minimum (4 GB recommended for large batches)
- **Display**: Minimum 800×700 resolution

---

## Version Information

Current version includes:
- ✓ Smart column detection for MassHunter exports
- ✓ Excel/CSV support for DIGEST files
- ✓ Wide and long format REF file support
- ✓ Automatic REF name extraction and matching
- ✓ Multi-mode channel selection
- ✓ Color-coded Excel output
- ✓ Processing summary with QC metrics
- ✓ Missing sample detection and warnings

---

Last updated: November 2025

