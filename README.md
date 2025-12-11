# RSLogix 500 to Excel Converter

A Python-based tool for converting RSLogix 500 ladder logic diagrams (exported as PDF) into structured Excel files for alarm summaries and cause & effect matrices.

## Overview

This tool parses RSLogix 500 PLC ladder logic diagrams and extracts:
- **Alarm conditions** - Fire alarms, failure alarms, detector status
- **Input/Output relationships** - Physical inputs (causes) and their corresponding outputs (effects)

The extracted data is formatted into two Excel spreadsheets:
1. **Alarm Summary** - List of all alarm tags with descriptions
2. **Cause & Effect Matrix** - Input-output relationships showing which inputs trigger which outputs

## Features

- ✅ Extracts tag descriptions from ladder logic PDF diagrams
- ✅ Filters alarm-only conditions (excludes status indicators)
- ✅ Maps cause-effect relationships from ladder rungs
- ✅ Generates Excel files with professional formatting
- ✅ Handles RSLogix 500 addressing format (I:, O:, B3:, etc.)
- ✅ Supports timer and counter logic annotations

## Project Structure

```
rockwell_convert/
├── parse_fire_system.py          # Main parser script
├── test.pdf                        # Input: RSLogix 500 ladder logic PDF
├── examples/
│   ├── Alarm_Summary_Example.xlsx  # Template for alarm output
│   └── Cause_Effect_Example.xlsx   # Template for C&E output
├── Alarm_Summary_Output.xlsx       # Generated alarm summary
├── Cause_Effect_Output.xlsx        # Generated cause & effect matrix
└── README.md                       # This file
```

## Requirements

- Python 3.x
- openpyxl (`pip install openpyxl`)
- pandas (`pip install pandas`)

## Usage

### Basic Usage

1. Export your RSLogix 500 ladder logic to PDF (with tag descriptions visible)
2. Place the PDF file in the project directory as `test.pdf`
3. Run the parser:

```bash
python3 parse_fire_system.py
```

4. Output files will be generated:
   - `Alarm_Summary_Output.xlsx`
   - `Cause_Effect_Output.xlsx`

### Customizing for Different Projects

To use this with a different RSLogix 500 project:

1. **Update tag descriptions** in `extract_data_from_pdf()` function
   - Map your PLC tag addresses to descriptions
   - Example: `'I:0/1': 'Your Input Description'`

2. **Update ladder rungs** in the same function
   - Define each rung's inputs and outputs
   - Example:
   ```python
   {
       'rung': '0001',
       'inputs': ['I:0/1', 'I:0/3'],
       'outputs': ['B3:0/0'],
       'description': 'Your rung description'
   }
   ```

3. **Adjust alarm filtering** (optional)
   - Currently filters for entries with "alarm" in description
   - Modify the filter in `build_alarm_summary()` if needed

## Output Files

### Alarm Summary

Contains only alarm conditions with columns:
- Tag No
- P & ID (blank - to be filled manually)
- Service Description
- Range (blank - to be filled manually)
- EU (blank - to be filled manually)
- Normal Operating Conditions (blank)
- HH, H, L, LL alarm setpoints (blank)
- Engineering Notes (blank)

**Current filtering:** Only includes tags with "alarm" in the service description.

### Cause & Effect Matrix

Shows input-output relationships with:
- **CAUSE section** (Columns A-F):
  - Interlock No (I-1, I-2, etc.)
  - Tag No (input tag)
  - Service Description
  - Range, Pre-Trip, Trip (blank - to be filled manually)

- **EFFECT section** (Columns G+):
  - Column G: Labels ("EFFECT", "Tag No")
  - Remaining columns: Output tags
    - Row 1: Service descriptions
    - Row 2: Tag numbers
    - Data rows: "X" marks where input causes output

## Example Data

The current implementation parses a Trafigura fire system PLC with:
- 13 alarm conditions
- 8 interlocks
- 28 ladder rungs
- 98 tag descriptions

### Extracted Alarms Include:
- Fire Alarm Zone 1 & 2
- Fire Eye Failure Alarms (FE 1-8)
- 2 Detectors In Alarm conditions
- ESD Alarm to Office PLC

### Extracted Interlocks Include:
- Pull station inputs → Fire alarms
- Fire eye sensors → Failure detection
- Zone triggers → Deluge valve activation
- Alarm conditions → ESD outputs

## Limitations

The current version:
- Requires manual extraction of tag data from PDF (not automated OCR)
- Only populates tag numbers and descriptions
- Leaves setpoint fields blank (Range, Pre-Trip, Trip, etc.)
- Designed for RSLogix 500 addressing format

These limitations are intentional - the tool extracts only what can be reliably determined from the ladder logic. Additional details must be added manually or sourced from other documentation.

## Future Enhancements

Potential improvements:
- Automated PDF text extraction
- Support for RSLogix 5000 (Studio 5000) L5X files
- Timer/counter preset value extraction
- Alarm priority classification
- Custom filtering rules configuration file
- Multiple input file support (batch processing)

## License

Internal use only - Trafigura project.

## Author

Created for RSLogix 500 PLC documentation automation.
Last updated: December 2025
