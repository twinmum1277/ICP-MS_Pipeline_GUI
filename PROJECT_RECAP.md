# ICP-MS Pipeline GUI - Project Recap

## Overview

This is a Python-based GUI application for processing Agilent MassHunter ICP-MS data exports. It automates:
- Data correction (blank subtraction, dilution factor application)
- Quality control calculations (ICV, SRM recovery)
- Report generation (Excel output with multiple sheets)

## Current Architecture

### Main Components

1. **`icpms_gui.py`** - Tkinter GUI front-end
   - File selection interface
   - Processing summary panel
   - Settings (output units: ppb/ppm)

2. **`process_icpms_batch.py`** - Core processing engine
   - File loading and parsing
   - Data correction pipeline
   - QC calculations
   - Excel report generation

### Input Files

1. **SORT file** - MassHunter CSV export (raw instrument data)
2. **DIGEST file** - Dilution factors (CSV or Excel)
3. **ICV file** - QC target values (CSV)
4. **REF file** (optional) - Reference material values (CSV)

### Processing Pipeline

```
1. Load SORT → Parse headers → Convert to long format
2. Load DIGEST → Match sample IDs → Apply dilution factors
3. Load ICV → Load REF (optional) → Get QC targets
4. Compute blank statistics (mean, SD, MDL = 3×SD)
5. Correct samples: (raw - avg_blank) × df ÷ 1000
6. Calculate ICV recovery % (target: 90-110%)
7. Calculate REF recovery % (target: 80-120%)
8. Select best channel for multi-mode elements
9. Generate Excel output:
   - "Corrected ppm (Wide)" - Final concentrations (pivot table)
   - "Corrected ppm (Long)" - Long format data
   - "QC Summary" - Recovery percentages and pass/fail
   - "Below Detection Limit" - Samples below MDL
```

## Current Quality Control Checks

### 1. ICV (Initial Calibration Verification)
- **Calculation**: `(corrected_value / icv_target) × 100`
- **Acceptance**: 90-110% recovery
- **Purpose**: Verify calibration accuracy

### 2. SRM/REF (Standard Reference Material)
- **Calculation**: `(corrected_value / ref_target) × 100`
- **Acceptance**: 80-120% recovery
- **Purpose**: Verify analytical accuracy with certified materials

### 3. Below Detection Limit (BDL)
- **Calculation**: `(raw_conc - avg_blank) < MDL` where `MDL = 3 × blank_SD`
- **Purpose**: Flag samples below method detection limit
- **Output**: Separate sheet listing all BDL values

### 4. Channel Selection
- For elements with multiple measurement channels/modes
- Selects channel that passes both ICV and REF criteria
- Falls back to channel with REF recovery closest to 100%

## Current MassHunter Compatibility

### What Works
- ✅ Detects date/time column (searches for 'date', 'time', 'acq' keywords)
- ✅ Detects sample name column (searches for 'sample', 'name' keywords)
- ✅ Handles "Conc." row (automatically removes if present)
- ✅ Parses element headers: `63  Cu  [ He ]` and `75 -> 91  As  [ O2 ]`
- ✅ Handles header in row 2 (if row 1 has many "Unnamed" columns)

### Potential Issues
- ⚠️ May fail if MassHunter export format changes
- ⚠️ Column detection relies on keyword matching (may miss non-standard names)
- ⚠️ No validation that parsed data makes sense
- ⚠️ Debug prints may clutter output (should be optional)

## Areas for Improvement

### 1. MassHunter Compatibility (Priority: HIGH)
**Goal**: Ensure widget works with direct MassHunter exports without manual editing

**Potential Issues to Address**:
- Different header row positions
- Alternative column naming conventions
- Missing or extra metadata rows
- Encoding issues with special characters
- Different date/time formats

**Recommended Approach**:
- Add more robust header detection
- Support multiple MassHunter export formats
- Add validation and error messages
- Create test cases with real MassHunter exports

### 2. Additional Quality Control Checks (Priority: MEDIUM)
**Goal**: Expand QC capabilities for comprehensive data validation

**Potential Additional Checks**:

#### A. Duplicate Sample Precision
- **What**: Calculate RSD (relative standard deviation) for duplicate samples
- **Acceptance**: RSD < 10-20% (element-dependent)
- **Implementation**: Identify samples with "DUP" in name, calculate RSD per element

#### B. ICV Drift Check
- **What**: Compare ICV before batch vs after batch
- **Acceptance**: Drift < 5-10% between first and last ICV
- **Implementation**: Track ICV sequence, calculate drift per element

#### C. Blank Contamination Check
- **What**: Flag blanks with values above threshold
- **Acceptance**: Blank < 3×MDL or < 10% of lowest sample
- **Implementation**: Compare blank values to MDL and sample values

#### D. Carryover Check
- **What**: Check if high-concentration samples affect subsequent samples
- **Acceptance**: No significant carryover pattern
- **Implementation**: Analyze sequence order, flag suspicious patterns

#### E. Internal Standard Recovery (if available)
- **What**: Check internal standard recovery (if MassHunter exports this)
- **Acceptance**: Typically 80-120%
- **Implementation**: Parse IS data from SORT file if present

#### F. Linearity Check
- **What**: Verify calibration curve linearity (if calibration data available)
- **Acceptance**: R² > 0.99
- **Implementation**: Would require calibration data export

#### G. Sample Sequence Validation
- **What**: Ensure proper sequence (blanks, ICV, samples, ICV, blanks)
- **Acceptance**: Follows expected pattern
- **Implementation**: Validate sample order against expected sequence

## Next Steps

1. **Test with Real MassHunter Exports**
   - Collect sample exports from different MassHunter versions
   - Test edge cases (missing columns, different formats)
   - Document any manual editing currently required

2. **Enhance MassHunter Parser**
   - Add more flexible header detection
   - Support alternative export formats
   - Add validation and clear error messages
   - Make debug output optional/configurable

3. **Design Additional QC Checks**
   - Prioritize which QC checks are most valuable
   - Design data structures to support new checks
   - Plan output format (new Excel sheets? columns in existing sheets?)

4. **Implement Selected QC Checks**
   - Start with highest priority checks
   - Add to processing pipeline
   - Update Excel output format
   - Update GUI to display new QC metrics

## Questions for Discussion

1. **MassHunter Compatibility**:
   - What specific issues have you encountered with MassHunter exports?
   - Are there particular export formats or versions that don't work?
   - What manual editing is currently required?

2. **Additional QC Checks**:
   - Which QC checks are most important for your workflow?
   - Do you have duplicate samples in your batches?
   - Do you run ICV before and after batches?
   - Are internal standard values available in your exports?
   - What other QC metrics do you currently track manually?

3. **Output Format**:
   - How should additional QC checks be displayed?
   - New Excel sheets? Additional columns? Summary dashboard?
   - What pass/fail criteria should be used?

