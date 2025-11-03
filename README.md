# ICP-MS Batch Processor

A Python-based GUI application for processing Agilent MassHunter ICP-MS data exports. This tool automates data correction, quality control calculations, and report generation for batch ICP-MS analyses.

## Features

- 📊 **Automated Data Processing**: Processes raw MassHunter CSV exports with blank correction and dilution factor application
- 🎯 **Quality Control**: Calculates ICV and SRM recovery percentages with pass/fail criteria
- 🔬 **Multi-Mode Support**: Handles multiple gas modes (He, O2, etc.) and mass-shift modes automatically
- 📈 **Channel Selection**: Intelligently selects the best measurement channel for elements analyzed in multiple modes
- 📑 **Excel Reports**: Generates comprehensive output with corrected concentrations, QC summary, and below-detection-limit data
- 🖥️ **User-Friendly GUI**: Simple Tkinter interface for file selection and batch processing

## Installation

### Prerequisites

- Python 3.7 or higher
- Required packages:
  ```bash
  pip install pandas numpy openpyxl xlsxwriter
  ```

### Setup

1. Clone this repository:
   ```bash
   git clone https://github.com/YOUR_USERNAME/ICP-MS_Pipeline_GUI.git
   cd ICP-MS_Pipeline_GUI
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### GUI Mode (Recommended)

1. Launch the GUI:
   ```bash
   python icpms_gui.py
   ```

2. Select your input files:
   - **SORT file**: MassHunter data export (CSV)
   - **DIGEST file**: Sample dilution factors (CSV or Excel)
   - **ICV file**: ICV and SRM target values (CSV)
   - **CRM file** (optional): Certified reference material values

3. Click **Run Batch** to process

4. Output Excel file will be saved in the same directory as your SORT file

### Command Line Mode

```bash
python process_icpms_batch.py \
  --sort path/to/sort.csv \
  --digest path/to/digest.csv \
  --icv path/to/icv.csv \
  --out path/to/output.xlsx \
  --crm path/to/crm.csv  # optional
```

## Input File Formats

### SORT File (MassHunter Export)
Raw instrument data with sample names and element concentrations. See `sort_template.csv` for format.

### DIGEST File
Dilution/digestion factors for each sample. Supports CSV or Excel.

Required columns:
- `sample_id`: Sample identifier (must match SORT file)
- `df`: Dilution factor (numeric)

### ICV File
Target concentration values for quality control samples.

Required columns:
- `element`: Element symbol (e.g., Cu, Zn, As)
- `icv_target`: ICV target concentration
- `srm_target`: SRM target concentration (optional)

### CRM File (Optional)
Certified reference material values that override ICV file SRM targets.

See `TEMPLATE_GUIDE.md` for detailed format specifications and examples.

## Output

The tool generates an Excel workbook with three sheets:

1. **Corrected ppm**: Final concentration values for all samples (blanks, ICV, SRM excluded)
2. **QC Summary**: ICV and SRM recovery percentages with pass/fail status per element
3. **Below Detection Limit**: Samples with values below the method detection limit (3× blank SD)

## Sample Naming Conventions

For proper quality control processing, name your samples:
- **Blanks**: Include "BLANK" (e.g., `BLANK_1`, `BLANK_2`)
- **ICV samples**: Include "ICV" (e.g., `ICV_1`, `ICV_2`)
- **Reference materials**: Include "SRM", "NIST", "CRM", or "USGS"
- **Duplicates**: Include "DUP" (excluded from final output)

## Processing Workflow

1. Load SORT (raw data), DIGEST (dilution factors), and ICV (QC targets)
2. Parse instrument data and normalize sample names
3. Compute blank statistics (mean, SD, MDL)
4. Apply blank correction: `(raw - avg_blank) × df ÷ 1000`
5. Calculate ICV and SRM recovery percentages
6. Select best measurement channel for multi-mode elements
7. Generate Excel output with results and QC summary

## Quality Control Criteria

- **ICV Acceptance**: 90-110% recovery
- **SRM Acceptance**: 80-120% recovery
- **Method Detection Limit (MDL)**: 3 × blank standard deviation

## Templates

Template files are provided to help you format your input data:
- `sort_template.csv` - Example MassHunter export format
- `digest_template.csv` - Dilution factors template
- `icv_template.csv` - QC targets template
- `crm_values_template.csv` - Optional CRM values template

See `TEMPLATE_GUIDE.md` for comprehensive documentation.

## Troubleshooting

### Common Issues

**File picker not appearing (macOS)**
- Resize the GUI window to ensure Browse buttons are visible
- The app auto-sizes on launch but may need adjustment

**Sample names don't match**
- Sample names are normalized (uppercase, spaces → underscores)
- Ensure DIGEST uses the same naming convention as SORT

**Missing elements in output**
- Verify element symbols in ICV file match those in SORT headers
- Check for typos in element symbols

**High number of below-detection-limit values**
- Check for blank contamination
- Verify instrument stability during blank measurements

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## License

MIT License - See LICENSE file for details

## Author

Developed at Dartmouth College

## Acknowledgments

This tool was developed to streamline ICP-MS data processing workflows for geochemical and environmental analyses.
