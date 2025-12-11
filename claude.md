# Implementation Plan: RSLogix 500 to Excel Converter

This document provides a step-by-step implementation plan for recreating or adapting this RSLogix 500 parser for future projects.

## Table of Contents
1. [Project Overview](#project-overview)
2. [Input Requirements](#input-requirements)
3. [Implementation Steps](#implementation-steps)
4. [Design Decisions](#design-decisions)
5. [Code Structure](#code-structure)
6. [Customization Guide](#customization-guide)
7. [Common Issues & Solutions](#common-issues--solutions)

---

## Project Overview

### Goal
Convert RSLogix 500 PLC ladder logic diagrams into two structured Excel files:
1. **Alarm Summary** - List of alarm conditions
2. **Cause & Effect Matrix** - Input/output relationships (interlocks)

### Why This Approach?
- RSLogix 500 L5X files lose critical information during conversion from .RSS format
- PDF exports preserve visual layout and descriptions from the original ladder logic
- Manual extraction ensures accuracy for safety-critical fire system documentation

---

## Input Requirements

### Required Files

1. **PDF Ladder Logic** (`test.pdf`)
   - Export from RSLogix 500 with descriptions visible
   - Must show tag addresses and yellow description labels
   - Include all rungs with I/O tags

2. **CSV Tags File** (optional reference)
   - `TRAFIGURA_FIRE_SYSTEM_PLC_1_NEW-Controller-Tags.CSV`
   - Can be used to cross-reference tag descriptions
   - Not required if PDF has complete information

3. **Example Templates**
   - `examples/Alarm_Summary_Example.xlsx` - Format reference
   - `examples/Cause_Effect_Example.xlsx` - Format reference

### Input Data Structure

From the PDF, you need to extract:

```python
# Tag Descriptions
tag_descriptions = {
    'I:0/0': 'Pull Station 1 Zone 2',    # Input address: Description
    'O:0/0': 'Deluge Valve Zone 2 Open', # Output address: Description
    'B3:0/0': 'Fire Alarm Zone 1',       # Internal bit: Description
}

# Ladder Rungs
rungs = [
    {
        'rung': '0001',                           # Rung number
        'inputs': ['I:0/1', 'I:0/3', 'B3:4/6'],  # Input tags (causes)
        'outputs': ['B3:0/0', 'B3:2/0'],         # Output tags (effects)
        'description': 'Fire Alarm Zone 1',      # Rung description
        'timer': 'T4:1',                         # Optional: Timer used
        'counter': 'C5:0',                       # Optional: Counter used
    }
]
```

---

## Implementation Steps

### Step 1: Analyze the PDF

**Goal:** Understand the ladder logic structure

1. Open the PDF and identify:
   - Physical inputs (I: addresses) - These are your CAUSES
   - Physical outputs (O: addresses) - These are your EFFECTS
   - Internal bits (B3:, B11:, etc.) that store alarm states
   - Yellow description labels showing what each tag does

2. Map the I/O addressing:
   ```
   I:0/0 through I:0/19  → First input module (Bul.1766)
   I:1/0 through I:1/15  → Second input module (1762-IQ16)
   O:0/0 through O:0/x   → Output module
   ```

3. Identify the rung patterns:
   - XIC (Examine If Closed) - Normal input check
   - XIO (Examine If Open) - Inverted input check
   - OTE (Output Energize) - Sets output when conditions true
   - TON (Timer On Delay) - Delays before activating
   - CTU (Count Up) - Counts events

### Step 2: Extract Tag Descriptions

**Goal:** Build the `tag_descriptions` dictionary

For each tag in the PDF:

```python
# Find the yellow label boxes in the PDF
# Example from Page 1, Rung 0001:
# Yellow box shows: "Pull Station 2 Zone 1" next to "I:0/1"

tag_descriptions['I:0/1'] = 'Pull Station 2 Zone 1'
```

**Systematic approach:**
1. Go page by page through the PDF
2. For each yellow label, note the tag address and description
3. Pay attention to:
   - Input addresses (I:)
   - Output addresses (O:)
   - Status bits (B3:)
   - Timer/Counter descriptions (T4:, C5:)

### Step 3: Extract Ladder Rungs

**Goal:** Build the `rungs` list

For each rung:

```python
# Example: Rung 0001 from PDF
# Shows multiple inputs (OR'd together) leading to two outputs

{
    'rung': '0001',
    'inputs': [
        'I:0/1',   # Pull Station 2 Zone 1
        'I:0/3',   # Pull Station 4 Zone 1
        'I:0/5',   # Pull Station 6 Zone 1
        'B11:0/1', # Pull Station 8 Zone 1
        'B14:0/0', # Pull Station 10 from Office
        'B3:4/6'   # 2 Detectors In Alarm Zone 1
    ],
    'outputs': [
        'B3:0/0',  # Fire Alarm Zone 1
        'B3:2/0'   # Plant ESD
    ],
    'description': 'Fire Alarm Zone 1'
}
```

**Reading the ladder:**
- Vertical lines between inputs = OR logic
- Horizontal sequence = AND logic
- Multiple outputs on right side = all activate together
- Look for branching logic (parallel paths)

### Step 4: Implement the Parser Script

**File structure:**

```python
# parse_fire_system.py

def extract_data_from_pdf():
    """
    Returns: (rungs, tag_descriptions)
    """
    # Paste your extracted data here
    tag_descriptions = { ... }
    rungs = [ ... ]
    return rungs, tag_descriptions

def build_alarm_summary(tag_descriptions):
    """
    Filter for alarm-only tags
    """
    # Only include tags with 'alarm' in description
    # Return list of alarm dictionaries

def build_cause_effect_matrix(rungs, tag_descriptions):
    """
    Build interlock list from rungs
    """
    # Filter rungs with physical I/O
    # Return list of interlock dictionaries

def generate_alarm_summary_excel(alarms, output_file):
    """
    Create formatted Excel file
    """
    # Use openpyxl to create spreadsheet
    # Apply formatting (colors, borders, column widths)

def generate_cause_effect_excel(interlocks, tag_descriptions, output_file):
    """
    Create formatted C&E matrix
    """
    # Similar to alarm summary
    # Add effect columns dynamically

def main():
    """
    Run the full pipeline
    """
    rungs, tags = extract_data_from_pdf()
    alarms = build_alarm_summary(tags)
    interlocks = build_cause_effect_matrix(rungs, tags)
    generate_alarm_summary_excel(alarms, 'Alarm_Summary_Output.xlsx')
    generate_cause_effect_excel(interlocks, tags, 'Cause_Effect_Output.xlsx')
```

### Step 5: Configure Filtering

**Alarm Summary filtering:**

Current rule: Only include tags with "alarm" in description (case-insensitive)

```python
if 'alarm' in description.lower():
    # Include this tag
```

**To add more conditions:**

```python
# Include alarms AND warnings
if 'alarm' in description.lower() or 'warning' in description.lower():
    # Include this tag

# Exclude certain types
if 'alarm' in description.lower() and 'test' not in description.lower():
    # Include this tag
```

**Cause & Effect filtering:**

Current rule: Include rungs where:
- Has physical input (I:, B11:, B14:), AND
- Has physical output (O:) OR shutdown output (B3:2, B3:10, B3:0/0, B3:0/11)

```python
has_physical_input = any(tag.startswith(('I:', 'B11:', 'B14:')) for tag in rung['inputs'])
has_physical_output = any(tag.startswith(('O:')) for tag in rung['outputs'])
has_shutdown = any(tag.startswith(('B3:2', 'B3:10', 'B3:0/0', 'B3:0/11')) for tag in rung['outputs'])

if has_physical_input and (has_physical_output or has_shutdown):
    # Include this rung as an interlock
```

### Step 6: Test and Validate

1. **Run the script:**
   ```bash
   python3 parse_fire_system.py
   ```

2. **Check the output:**
   - Alarm count matches expectations
   - Interlock count covers main safety functions
   - Excel formatting looks correct

3. **Validate against PDF:**
   - Spot-check several alarms
   - Verify cause-effect relationships
   - Ensure no critical interlocks missed

4. **Review with engineering:**
   - Share output files
   - Get feedback on completeness
   - Iterate if needed

---

## Design Decisions

### Why Manual Extraction Instead of Automated PDF Parsing?

**Decision:** Manually build tag and rung dictionaries from PDF analysis

**Rationale:**
1. **Accuracy**: Fire systems are safety-critical; manual review ensures correctness
2. **PDF Variability**: Different RSLogix exports have inconsistent formatting
3. **Description Quality**: Yellow labels in PDF provide clear, human-verified descriptions
4. **One-time Effort**: Most projects are one-off conversions, not batch processing

**Trade-off:** More initial setup time, but higher confidence in results

### Why Filter Alarms by "alarm" Keyword?

**Decision:** Only include tags with "alarm" in the description

**Rationale:**
1. **Focus**: Alarm summary should show alarms, not status indicators
2. **Noise Reduction**: Excludes "Strobe Light On", "Valve Open", etc.
3. **Extensible**: Easy to add more keywords later (warning, fault, etc.)

**Examples:**
- ✅ Included: "Fire Alarm Zone 1", "FE 1 Failure Alarm"
- ❌ Excluded: "Strobe Light On", "Deluge Valve Open"

### Why Separate Tag Numbers and Descriptions in C&E Matrix?

**Decision:** Use separate rows for tag numbers and descriptions in effect columns

**Rationale:**
1. **Readability**: Easier to scan for specific tags
2. **Excel Formatting**: Avoids line-wrapping issues in cells
3. **Sorting**: Can sort by tag or description independently
4. **Standard Format**: Matches industry-standard C&E matrices

**Layout:**
```
Row 1: EFFECT | Fire Alarm Zone 1 | Fire Eye Failure | ...
Row 2: Tag No | B3:0/0            | B3:0/8           | ...
```

### Why Leave Setpoint Fields Blank?

**Decision:** Range, Pre-Trip, Trip, EU, Normal Conditions all blank

**Rationale:**
1. **Not in Ladder Logic**: RSLogix 500 ladder doesn't contain setpoint values
2. **Requires Additional Docs**: Must come from P&IDs, datasheets, or specifications
3. **Avoid Guessing**: Better to leave blank than populate with incorrect data
4. **Manual Entry Workflow**: Engineers will fill these in during review

---

## Code Structure

### Main Functions

```
extract_data_from_pdf()
    └─> Returns: (rungs, tag_descriptions)
        - Hard-coded dictionaries extracted from PDF
        - This is where you customize for each project

build_alarm_summary(tag_descriptions)
    └─> Returns: List of alarm dictionaries
        - Filters tags by "alarm" keyword
        - Formats for Excel output

build_cause_effect_matrix(rungs, tag_descriptions)
    └─> Returns: List of interlock dictionaries
        - Filters rungs with physical I/O
        - Maps inputs to outputs
        - Adds timer/counter annotations

generate_alarm_summary_excel(alarms, output_file)
    └─> Creates: Alarm_Summary_Output.xlsx
        - Yellow headers
        - Borders and formatting
        - Proper column widths

generate_cause_effect_excel(interlocks, tag_descriptions, output_file)
    └─> Creates: Cause_Effect_Output.xlsx
        - Green effect headers
        - Yellow cause headers
        - Separate tag/description rows
        - Dynamic effect columns

main()
    └─> Orchestrates the full pipeline
        - Calls all functions in sequence
        - Prints progress
```

### Key Data Structures

**Tag Descriptions:**
```python
{
    'I:0/1': 'Pull Station 2 Zone 1',
    'B3:0/0': 'Fire Alarm Zone 1',
}
```

**Ladder Rungs:**
```python
{
    'rung': '0001',
    'inputs': ['I:0/1', 'I:0/3'],
    'outputs': ['B3:0/0'],
    'description': 'Fire Alarm Zone 1',
    'timer': 'T4:1',      # Optional
    'counter': 'C5:0',    # Optional
    'logic_type': 'OR',   # Optional: OR, AND, XIO
}
```

**Alarms:**
```python
{
    'Tag No': 'B3:0/0',
    'P & ID': '',
    'Service Description': 'Fire Alarm Zone 1',
    'Range': '',
    'EU': '',
    'Normal Operating Conditions': '',
    'HH': '',
    'H': '',
    'L': '',
    'LL': '',
    'Engineering Notes': ''
}
```

**Interlocks:**
```python
{
    'Interlock No': 1,
    'Tag No': 'I:0/1',
    'Service Description': 'Pull Station 2 Zone 1',
    'Range': '',
    'Pre-Trip (H or L)': '',
    'Trip (HH or LL)': '',
    'P & ID': '',
    'Rung': '0001',
    'Effects': {
        'B3:0/0': 'X',
        'B3:2/0': 'X'
    },
    'All Inputs': ['I:0/1', 'I:0/3', 'B3:4/6'],
    'All Outputs': ['B3:0/0', 'B3:2/0']
}
```

---

## Customization Guide

### For a New RSLogix 500 Project

**Step 1: Get the PDF**
- Export ladder logic from RSLogix 500
- Ensure descriptions are visible
- Save as `test.pdf`

**Step 2: Clear Old Data**
```python
def extract_data_from_pdf():
    tag_descriptions = {}  # Start empty
    rungs = []             # Start empty
```

**Step 3: Add Your Tag Descriptions**

Go through the PDF and build the dictionary:
```python
tag_descriptions = {
    # Your inputs
    'I:0/0': 'Your first input description',
    'I:0/1': 'Your second input description',

    # Your outputs
    'O:0/0': 'Your first output description',

    # Your internal bits
    'B3:0/0': 'Your alarm description',
}
```

**Step 4: Add Your Rungs**

For each rung in the PDF:
```python
rungs.append({
    'rung': '0001',  # Rung number from PDF
    'inputs': ['I:0/0', 'I:0/1'],  # All input tags
    'outputs': ['O:0/0'],          # All output tags
    'description': 'What this rung does'
})
```

**Step 5: Adjust Filtering (if needed)**

Alarm filter keywords:
```python
# Current: only "alarm"
if 'alarm' in description.lower():

# Add more keywords:
keywords = ['alarm', 'warning', 'fault', 'trip']
if any(kw in description.lower() for kw in keywords):
```

Interlock filtering:
```python
# Current implementation:
has_physical_input = any(tag.startswith(('I:', 'B11:', 'B14:'))
                         for tag in rung['inputs'])
has_physical_output = any(tag.startswith(('O:'))
                          for tag in rung['outputs'])
has_shutdown = any(tag.startswith(('B3:2', 'B3:10'))
                   for tag in rung['outputs'])

# Modify tag prefixes for your system:
# - Change 'B11:', 'B14:' to your digital input bit files
# - Change 'B3:2', 'B3:10' to your shutdown/alarm bit files
```

**Step 6: Run and Validate**
```bash
python3 parse_fire_system.py
```

### For RSLogix 5000 / Studio 5000 Projects

**Different Approach Needed:**

RSLogix 5000 exports to `.L5X` XML files with complete tag information.

**Recommended changes:**
1. Parse the L5X file directly (XML parsing)
2. Extract tag descriptions from `<Tag>` and `<Comment>` elements
3. Parse `<RLLContent>` for ladder logic
4. Use same Excel generation functions

**Advantages:**
- No manual data entry
- More automation possible
- Preserves all tag metadata

**See:** For Studio 5000 projects, use L5X parsing instead of PDF extraction.

---

## Common Issues & Solutions

### Issue 1: "ModuleNotFoundError: No module named 'openpyxl'"

**Solution:**
```bash
pip3 install openpyxl pandas
```

### Issue 2: Excel cells showing "\n" instead of newlines

**Cause:** Using `\\n` (escaped) instead of `\n` (actual newline)

**Solution:**
```python
# Wrong:
headers = ['Tag No', 'Service Description', 'Normal Operating\\nConditions']

# Correct:
headers = ['Tag No', 'Service Description', 'Normal Operating\nConditions']
```

### Issue 3: Too many/too few alarms in output

**Cause:** Filter not matching your project's naming conventions

**Solution:** Adjust the alarm filter:
```python
# Current filter:
if 'alarm' in description.lower():

# More inclusive:
if any(word in description.lower() for word in ['alarm', 'warning', 'fault']):

# More restrictive:
if 'alarm' in description.lower() and 'test' not in description.lower():
```

### Issue 4: Interlocks missing from C&E matrix

**Cause:** Input/output tag prefixes don't match filter

**Solution:** Update the filter in `build_cause_effect_matrix()`:
```python
# Add your tag prefixes:
has_physical_input = any(tag.startswith(('I:', 'B11:', 'B14:', 'YOUR_PREFIX:'))
                         for tag in rung['inputs'])
```

### Issue 5: Excel column too narrow for descriptions

**Solution:** Adjust column widths in generation functions:
```python
# In generate_alarm_summary_excel():
ws.column_dimensions['C'].width = 60  # Increase as needed

# In generate_cause_effect_excel():
ws.column_dimensions['C'].width = 60  # Service description
for i, tag in enumerate(effect_columns):
    col_letter = chr(ord('H') + i)
    ws.column_dimensions[col_letter].width = 25  # Increase for longer descriptions
```

### Issue 6: Wrong effect columns in C&E matrix

**Cause:** Internal bits being included as effects instead of physical outputs

**Solution:** Tighten the effect filter:
```python
# Only include true physical outputs:
has_physical_output = any(tag.startswith(('O:')) for tag in rung['outputs'])
has_critical_alarm = any(tag in ['B3:0/0', 'B3:0/11'] for tag in rung['outputs'])

if has_physical_input and (has_physical_output or has_critical_alarm):
    # Include as interlock
```

---

## Testing Checklist

Before considering the conversion complete:

- [ ] All physical inputs mapped to descriptions
- [ ] All physical outputs mapped to descriptions
- [ ] All alarm bits have descriptions
- [ ] Alarm count seems reasonable (not too many/too few)
- [ ] Interlock count covers main safety functions
- [ ] Spot-check 5-10 alarms against PDF
- [ ] Spot-check 5-10 interlocks against ladder logic
- [ ] Excel formatting looks professional
- [ ] Newlines display correctly in headers
- [ ] Column widths are readable
- [ ] No "\n" showing in cells
- [ ] Green/yellow colors applied correctly
- [ ] Effect columns have both tag and description rows

---

## Next Steps

After generating the initial files:

1. **Review with Engineers**
   - Walk through the output
   - Verify critical interlocks
   - Check alarm completeness

2. **Fill in Blank Fields**
   - Add P & ID references
   - Add setpoint values from specs
   - Add operating ranges from datasheets

3. **Add Engineering Notes**
   - Document special conditions
   - Add calibration references
   - Note interlock bypass conditions

4. **Archive**
   - Keep PDF source in project folder
   - Version control the Python script
   - Save final Excel files with project number

---

## Version History

- **v1.0** (Dec 2025) - Initial implementation for Trafigura fire system
  - 13 alarms extracted
  - 8 interlocks extracted
  - Manual PDF extraction approach
  - RSLogix 500 format

---

## Contact

For questions or improvements to this process, contact the engineering documentation team.
