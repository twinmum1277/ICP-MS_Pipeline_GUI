# ICP-MS Batch Processor - Input File Templates

This guide explains the expected format for each input file used by the ICP-MS Batch Processor GUI.

---

## 1. SORT File (MassHunter Export)

**Format:** CSV  
**Template:** `sort_template.csv`

This is the raw data export from Agilent MassHunter ICP-MS software.

### Expected Structure:
- **Column 1**: Date/Time (e.g., "Acq. Date-Time")
- **Column 2**: Sample Name (e.g., "Sample Name")
- **Columns 3+**: Element concentration columns with headers like:
  - `63  Cu  [ He ]` - standard mass mode
  - `75 -> 91  As  [ O2 ]` - mass-shift mode
  - Format: `mass  Element  [ GasMode ]`

### Optional Second Row:
- May have "Conc." repeated across all data columns (will be automatically removed)

### Sample Naming Conventions:
- **BLANK** samples: Include "BLANK" in the name (e.g., `BLANK_1`, `BLANK_2`)
- **ICV** samples: Include "ICV" in the name (e.g., `ICV_1`, `ICV_2`)
- **Reference materials**: Include "SRM", "NIST", "CRM", or "USGS" in the name
- **Regular samples**: Any other naming convention

### Notes:
- Sample names will be normalized (converted to uppercase, spaces replaced with underscores)
- The processor handles multiple gas modes (He, O2, etc.) and mass-shift modes automatically

---

## 2. DIGEST File (Dilution Factors)

**Format:** CSV or Excel (.xlsx, .xls)  
**Template:** `digest_template.csv`

Contains dilution/digestion factors for each sample.

### Required Columns:
- **sample_id**: Must match the sample names in the SORT file
- **df**: Dilution factor (numeric)

### Example:
```csv
sample_id,df
BLANK_1,1
BLANK_2,1
ICV_1,1
ICV_2,1
SRM_NIST_2710,1
SAMPLE_001,50
SAMPLE_002,50
SAMPLE_003,100
SAMPLE_004,100
```

### Notes:
- Sample names will be normalized to match SORT file (uppercase, underscores)
- If a sample from SORT is not found in DIGEST, it will use a default df = 1.0
- All blanks, ICVs, and reference materials typically have df = 1

---

## 3. ICV File (Calibration Targets)

**Format:** CSV or Excel (.xlsx, .xls)  
**Template:** `icv_template.csv`

Contains target concentration values for ICV and SRM quality control samples.

### Required Columns:
- **element**: Element symbol (e.g., Cu, Zn, As, Pb)
- **icv_target**: Target concentration for ICV samples (numeric, typically in ppb or ppm)
- **srm_target**: Target concentration for SRM/reference material samples (optional)

### Example:
```csv
element,icv_target,srm_target
Cu,10.0,45.2
Zn,10.0,127.5
As,10.0,11.6
Pb,10.0,20.1
```

### Notes:
- Element symbols will be matched against elements detected in SORT file
- `icv_target` is used to calculate % recovery for ICV samples (acceptable range: 90-110%)
- `srm_target` can be left blank if using a separate CRM values file
- Only include elements that you're analyzing

---

## 4. CRM Values File (Optional)

**Format:** CSV  
**Template:** `crm_values_template.csv`

Optional file with more accurate reference material target values. If provided, these override the `srm_target` values in the CAL file.

### Required Columns:
- **element**: Element symbol
- **target_value**: Certified/accepted concentration value for the reference material

### Example:
```csv
element,target_value
Cu,45.2
Zn,127.5
As,11.6
Pb,20.1
```

### Notes:
- Use this when you have certified reference material values
- Particularly useful for NIST SRMs, USGS standards, etc.
- More accurate than generic srm_target values

---

## Processing Workflow

1. **Load SORT**: Raw instrument data with sample names and element concentrations
2. **Load DIGEST**: Dilution factors matched to sample names
3. **Load ICV**: Target values for quality control calculations
4. **Compute blanks**: Average and standard deviation from BLANK samples
5. **Correct samples**: `(raw - avg_blank) × df ÷ 1000` → corrected ppm
6. **QC calculations**: 
   - ICV recovery % = (corrected / icv_target) × 100
   - SRM recovery % = (corrected / srm_target) × 100
7. **Channel selection**: For elements with multiple channels/modes, select best based on QC
8. **Output Excel** with 3 sheets:
   - "Corrected ppm" - Final concentration values
   - "QC Summary" - Recovery percentages and pass/fail
   - "Below Detection Limit" - Samples below MDL (3× blank SD)

---

## Sample Naming Best Practices

### For proper QC processing, name your samples:
- **Blanks**: `BLANK_1`, `BLANK_2`, `BLANK_3`, etc.
- **Initial Calibration Verification (ICV)**: `ICV_1`, `ICV_2` (typically run before and after sample batch)
- **Reference Materials**: `SRM_2710`, `NIST_1643f`, `CRM_BCR-2`, `USGS_AGV-2`
- **Duplicates** (optional): Include "DUP" in name - these will be excluded from final output
- **Regular Samples**: Any naming convention that doesn't contain keywords above

---

## Quick Start

1. Use the provided templates as starting points
2. Replace the example data with your actual:
   - Sample names and concentrations (SORT)
   - Dilution factors (DIGEST)
   - Target values (ICV)
3. Ensure sample names match across SORT and DIGEST files
4. Launch the GUI and select your files
5. Click "Run Batch" to generate output Excel file

---

## Troubleshooting

### Common Issues:

**Sample names don't match:**
- The processor normalizes names (uppercase, spaces → underscores)
- `Sample 001` in SORT becomes `SAMPLE_001`
- Ensure DIGEST uses the same naming convention

**Missing elements in output:**
- Check that elements in ICV file match those in SORT column headers
- Element symbols are case-sensitive initially (but will be stripped/normalized)

**QC failures:**
- ICV acceptance: 90-110% recovery
- SRM acceptance: 80-120% recovery
- Check blank contamination if recoveries are off

**"Below Detection Limit" showing too many samples:**
- MDL = 3 × blank standard deviation
- High blank variability increases MDL
- Check for blank contamination or unstable instrument

---

For questions or issues, check the main processor docstring in `process_icpms_batch.py`.

