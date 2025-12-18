# ICP-MS Pipeline - Improvement Plan

## Issue 1: MassHunter Export Compatibility

### Current Status
The code has basic MassHunter export handling, but may require manual editing in some cases.

### Current Implementation
- ✅ Detects date/time column (keywords: 'date', 'time', 'acq')
- ✅ Detects sample name column (keywords: 'sample', 'name')
- ✅ Handles "Conc." row removal
- ✅ Handles header in row 2 (if row 1 has many "Unnamed" columns)
- ✅ Parses element headers: `63  Cu  [ He ]` and `75 -> 91  As  [ O2 ]`

### Potential Issues
1. **Column Detection**: Relies on keyword matching - may fail with non-standard column names
2. **Header Position**: Only checks first 2 rows - MassHunter may use different formats
3. **Encoding**: No explicit encoding handling (may fail with special characters)
4. **Validation**: No validation that parsed data is reasonable
5. **Error Messages**: Debug prints may not be user-friendly

### Recommended Improvements

#### A. Enhanced Header Detection
```python
# More robust column detection:
- Check multiple possible header rows (0, 1, 2)
- Look for common MassHunter column patterns
- Validate that detected columns contain expected data types
- Provide clear error messages if detection fails
```

#### B. Multiple Format Support
```python
# Support common MassHunter export variations:
- Different date/time formats
- Alternative column naming (e.g., "Acquisition Time" vs "Acq. Date-Time")
- Files with/without metadata rows
- Different element header formats
```

#### C. Better Error Handling
```python
# User-friendly error messages:
- "Could not find sample name column. Expected column containing 'sample' or 'name'"
- "Found X element columns but expected Y. Check file format."
- "Sample names found: [list]. Does this look correct?"
```

#### D. Validation
```python
# Validate parsed data:
- Check that sample_id column has non-empty values
- Verify element columns contain numeric data
- Warn if many values are missing or zero
- Check for duplicate sample names
```

### Action Items
1. **Test with Real Exports**: Collect sample MassHunter exports to identify edge cases
2. **Enhance Parser**: Add more flexible detection and validation
3. **Add Logging**: Replace debug prints with proper logging (optional verbose mode)
4. **Create Test Suite**: Test with various MassHunter export formats

---

## Issue 2: Additional Quality Control Checks

### Current QC Checks
1. **ICV Recovery**: 90-110% acceptance
2. **SRM/REF Recovery**: 80-120% acceptance  
3. **Below Detection Limit**: MDL = 3× blank SD
4. **Channel Selection**: Best channel for multi-mode elements

### Proposed Additional QC Checks

#### 1. Duplicate Sample Precision (HIGH PRIORITY)
**Purpose**: Verify analytical precision by comparing duplicate samples

**Calculation**:
- Identify duplicate samples (contain "DUP" in name)
- Match duplicates to original samples
- Calculate RSD (Relative Standard Deviation): `RSD = (SD / mean) × 100`

**Acceptance Criteria**:
- RSD < 10% for major elements (>100 ppb)
- RSD < 20% for trace elements (<100 ppb)
- Element-specific thresholds configurable

**Implementation**:
```python
def calculate_duplicate_precision(corrected_df):
    """
    Calculate RSD for duplicate samples.
    Returns DataFrame with sample_id, element, original_value, 
    duplicate_value, mean, sd, rsd_pct, pass/fail
    """
    # Find duplicates (sample names with "DUP")
    # Match to original samples
    # Calculate statistics
    # Return results
```

**Output**: New Excel sheet "Duplicate Precision" or column in QC Summary

---

#### 2. ICV Drift Check (MEDIUM PRIORITY)
**Purpose**: Detect instrument drift during batch analysis

**Calculation**:
- Identify ICV samples and their sequence order
- Compare first ICV vs last ICV
- Calculate drift: `drift_pct = |(last_ICV - first_ICV) / first_ICV| × 100`

**Acceptance Criteria**:
- Drift < 5% for stable elements
- Drift < 10% for volatile elements
- Configurable per element

**Implementation**:
```python
def calculate_icv_drift(icv_data, corrected_df):
    """
    Calculate drift between first and last ICV.
    Returns DataFrame with element, first_ICV_recovery, 
    last_ICV_recovery, drift_pct, pass/fail
    """
    # Sort ICV by acquisition time
    # Get first and last ICV per element
    # Calculate drift
    # Return results
```

**Output**: Column in QC Summary sheet

---

#### 3. Blank Contamination Check (MEDIUM PRIORITY)
**Purpose**: Ensure blanks are not contaminated

**Calculation**:
- Compare blank values to MDL
- Compare blank values to lowest sample value
- Flag if: `blank_value > 3×MDL` OR `blank_value > 10% of lowest_sample`

**Acceptance Criteria**:
- Blank < 3×MDL (already calculated)
- Blank < 10% of lowest sample value
- Configurable thresholds

**Implementation**:
```python
def check_blank_contamination(blank_stats, corrected_df):
    """
    Check for blank contamination.
    Returns DataFrame with element, blank_value, mdl, 
    lowest_sample_value, contamination_flag
    """
    # Get blank values per element
    # Get lowest sample value per element
    # Compare and flag
    # Return results
```

**Output**: Column in QC Summary or separate "Blank QC" sheet

---

#### 4. Carryover Check (LOW PRIORITY)
**Purpose**: Detect sample carryover effects

**Calculation**:
- Analyze sample sequence
- Flag if high-concentration sample followed by unexpectedly high blank/low sample
- Pattern detection: `sample[i] > threshold` followed by `sample[i+1] > expected`

**Acceptance Criteria**:
- No significant carryover pattern
- Configurable thresholds

**Implementation**:
```python
def check_carryover(corrected_df, sequence_col='acq_time'):
    """
    Check for sample carryover.
    Returns DataFrame with flagged samples and suspected source
    """
    # Sort by acquisition time
    # Identify high-concentration samples
    # Check subsequent samples
    # Flag suspicious patterns
    # Return results
```

**Output**: Column in "Below Detection Limit" sheet or separate "Carryover" sheet

---

#### 5. Internal Standard Recovery (IF AVAILABLE)
**Purpose**: Monitor instrument performance via internal standards

**Calculation**:
- Parse internal standard data from MassHunter export (if available)
- Calculate recovery: `(measured / expected) × 100`

**Acceptance Criteria**:
- IS recovery: 80-120% (typical)
- Configurable per IS element

**Implementation**:
- Requires parsing IS data from SORT file
- May need MassHunter export format that includes IS data
- **Status**: Depends on export format availability

**Output**: New "Internal Standard" sheet or column in QC Summary

---

#### 6. Sample Sequence Validation (LOW PRIORITY)
**Purpose**: Ensure proper analytical sequence

**Calculation**:
- Validate sequence follows expected pattern:
  - Blanks at start
  - ICV before samples
  - Samples in middle
  - ICV after samples
  - Blanks at end

**Acceptance Criteria**:
- Sequence follows expected pattern
- Warn if pattern is unusual

**Implementation**:
```python
def validate_sequence(corrected_df):
    """
    Validate sample sequence.
    Returns warnings if sequence is unusual
    """
    # Check sequence pattern
    # Flag deviations
    # Return warnings
```

**Output**: Warning in processing summary or separate "Sequence Validation" sheet

---

### Recommended Implementation Order

1. **Phase 1** (Immediate):
   - Duplicate Sample Precision
   - ICV Drift Check
   - Enhanced Blank Contamination Check

2. **Phase 2** (Next):
   - Carryover Check
   - Sample Sequence Validation

3. **Phase 3** (Future):
   - Internal Standard Recovery (if data available)
   - Additional checks as needed

---

### Output Format Options

#### Option A: Additional Excel Sheets
- "Duplicate Precision" sheet
- "ICV Drift" sheet
- "Blank QC" sheet
- etc.

**Pros**: Clear separation, easy to review
**Cons**: Many sheets, harder to get overview

#### Option B: Enhanced QC Summary Sheet
- Add columns for each new QC check
- One row per element with all QC metrics

**Pros**: Single view of all QC
**Cons**: Sheet may become wide

#### Option C: Hybrid Approach (RECOMMENDED)
- Add key metrics to QC Summary (duplicate RSD, ICV drift)
- Create separate detailed sheets for complex checks (carryover, sequence)
- Add QC summary dashboard with pass/fail overview

**Pros**: Best of both worlds
**Cons**: More complex to implement

---

## Questions for Discussion

### MassHunter Compatibility
1. What specific MassHunter export issues have you encountered?
2. What manual editing is currently required?
3. Do you have sample exports we can test with?
4. Are there specific MassHunter versions/formats to prioritize?

### Additional QC Checks
1. Which QC checks are most important for your workflow?
2. Do you run duplicate samples? How are they named?
3. Do you run ICV before and after batches?
4. Are internal standard values available in your exports?
5. What other QC metrics do you track manually?
6. What pass/fail criteria should we use for each check?

### Output Format
1. How should additional QC checks be displayed?
2. Do you prefer separate sheets or enhanced QC Summary?
3. Should we create a QC dashboard/overview sheet?
4. What format would be most useful for reporting?

---

## Next Steps

1. **Immediate**: Test current MassHunter parser with your exports
2. **Short-term**: Implement Phase 1 QC checks (duplicates, drift, blanks)
3. **Medium-term**: Enhance MassHunter parser based on findings
4. **Long-term**: Implement Phase 2-3 QC checks as needed

