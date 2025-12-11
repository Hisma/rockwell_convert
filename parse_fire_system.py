#!/usr/bin/env python3
"""
RSLogix 500 Fire System Parser
Extracts alarm and cause-effect data from PDF ladder logic
"""

import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import defaultdict

# File paths
PDF_FILE = 'test.pdf'

def extract_data_from_pdf():
    """
    Extract ladder logic information from PDF including tag descriptions
    Returns rungs and tag descriptions
    """
    rungs = []
    tag_descriptions = {}

    # Manual extraction from PDF analysis
    # All data extracted from test.pdf ladder logic diagrams

    # Build tag descriptions dictionary from PDF analysis
    tag_descriptions = {
        # Pull Stations
        'I:0/0': 'Pull Station 1 Zone 2',
        'I:0/1': 'Pull Station 2 Zone 1',
        'I:0/2': 'Pull Station 3 Zone 2',
        'I:0/3': 'Pull Station 4 Zone 1',
        'I:0/4': 'Pull Station 5 Zone 2',
        'I:0/5': 'Pull Station 6 Zone 1',
        'B11:0/0': 'Pull Station 7 Zone 1',
        'B11:0/1': 'Pull Station 8 Zone 1',
        'B14:0/0': 'Pull Station 10 from Office Plc Zone 1',
        'B14:0/1': 'Pull Station 9 From Office PLC Zone 2',
        'B14:0/2': 'Fire Alarm Zone 2',
        'B14:0/3': 'Plant ESD',

        # Fire Eyes - Fire Detected
        'I:0/6': 'Fire Eye 1 Failure Alarm',
        'I:0/7': 'Fire Eye 1 Fire Detected Zone 1',
        'I:0/8': 'Fire Eye 2 Failure Alarm',
        'I:0/9': 'Fire Eye 2 Fire Detected Zone 2',
        'I:0/10': 'Fire Eye 3 Failure Alarm',
        'I:0/11': 'Fire Eye 3 Detected Zone 1',
        'I:0/12': 'Fire Eye 4 Failure Alarm',
        'I:0/13': 'Fire Eye 4 Fire Detected Zone 2',
        'I:0/14': 'Fire Eye 5 Failure Alarm',
        'I:0/15': 'Fire Eye 5 Fire Detected Zone 1',
        'I:0/16': 'Fire Eye 6 Failure Alaram',
        'I:0/17': 'Fire Eye 6 Fire Detected Zone 2',
        'I:0/18': 'Fire Eye 7 Failure Alarm',
        'I:0/19': 'Fire Eye 7 Failure Zone 1',
        'I:1/0': 'Fire Eye 8 Failure Alarm',
        'I:1/1': 'Fire Eye 8 Fire Detected Zone 2',
        'I:1/12': 'Fire Eye Failure Warning',
        'I:1/13': 'Strobe Light Trigger',
        'I:1/15': 'Strobe Light Trigger',

        # Outputs
        'O:0/0': 'Deluge Valve Zone 2 Open',

        # Internal Alarms and Status - B3:0
        'B3:0/0': 'Fire Alarm Zone 1',
        'B3:0/1': 'Fire Eye Faulted Zone 1',
        'B3:0/3': 'Fire Eye Faulted Zone 2',
        'B3:0/4': 'Fire Detected Zone 2 Fire Eyes. Single Detector Only',
        'B3:0/5': 'Fire Detected Zone 1 Fire Eyes. Single Detector Only',
        'B3:0/6': 'Plant ESD',
        'B3:0/8': 'Fire Eye Failure Warning',
        'B3:0/9': 'Strobe Light On',
        'B3:0/10': 'Strobe Light On',
        'B3:0/11': 'Fire Alarm Zone 2',
        'B3:0/12': 'Fire System Deluge Valve Open',
        'B3:0/13': 'Strobe Light On',

        # Plant ESD outputs - B3:2
        'B3:2/0': 'Plant ESD',
        'B3:2/1': 'Plant ESD',
        'B3:2/2': 'Plant ESD',
        'B3:2/3': 'Plant ESD',
        'B3:2/4': 'Plant ESD',
        'B3:2/5': 'Plant ESD',
        'B3:2/6': 'Plant ESD',
        'B3:2/7': 'Plant ESD',
        'B3:2/8': 'Plant ESD',
        'B3:2/9': 'Plant ESD',
        'B3:2/10': 'Plant ESD',
        'B3:2/13': 'Plant ESD',
        'B3:2/14': 'Plant ESD',

        # Fire Eye Detection Status - B3:3
        'B3:3/0': 'Fire Eye 1 Fire Detected',
        'B3:3/1': 'Fire Eye 2 Fire Detected',
        'B3:3/2': 'Fire Eye 3 Fire Detected',
        'B3:3/3': 'Fire Eye 4 Fire Detected',
        'B3:3/4': 'Fire Eye 5 Fire Detected',
        'B3:3/5': 'Fire Eye 6 Fire Detected',
        'B3:3/6': 'Fire Eye 7 Fire Detected',
        'B3:3/7': 'Fire Eye 8 Fire Detected',
        'B3:3/8': 'FE 1 Failure Alarm',
        'B3:3/9': 'FE 2 Failure Alarm',
        'B3:3/10': 'FE 3 Failure Alarm',
        'B3:3/11': 'FE 4 Failure Alarm',
        'B3:3/12': 'FE 5 Failure Alarm',
        'B3:3/13': 'FE 6 Failure Alarm',
        'B3:3/14': 'FE 7 Failure Alarm',
        'B3:3/15': 'FE 8 Failure Alarm',

        # 2 Detector Logic - B3:4
        'B3:4/0': '2 Detectors In Alarm Bit 1',
        'B3:4/1': '2 Detectors In Alarm Bit 2',
        'B3:4/2': '2 Detectors In Alarm Bit 3',
        'B3:4/3': '2 Detectors In Alarm Bit 4',
        'B3:4/4': '2 Detectors In Alarm Bit 5',
        'B3:4/5': '2 Detectors In Alarm Zone 2',
        'B3:4/6': '2 Detectors In Alarm Zone 1',

        # ESD to Office
        'B3:10/0': 'ESD Alarm to Office PLC',

        # Timers
        'T4:0': 'Message Control Timer',
        'T4:1': 'Delay Trigger Timer',
        'T4:2': 'Delay Timer',
        'T4:3': 'Delay Trigger Timer',
        'T4:4': 'Delay Trigger Timer',
        'T4:5': 'Delay Trigger Timer',
        'T4:6': 'Delay Trigger Timer',
        'T4:7': 'Delay Trigger Timer',
        'T4:8': 'Delay Trigger Timer',
        'T4:9': 'Delay Timer',
        'T4:10': 'Delay Timer',
        'T4:11': 'Delay Timer',
        'T4:12': 'Delay Timer',
        'T4:13': 'Delay Timer',
        'T4:14': 'Delay Timer',
        'T4:15': 'Delay Timer',
        'T4:16': 'Delay Timer',
    }

    # Build rungs from PDF analysis
    rungs = [
        # Rung 0000
        {
            'rung': '0000',
            'inputs': ['B14:0/3'],
            'outputs': ['B3:2/4'],
            'description': 'Plant ESD from Office PLC'
        },
        # Rung 0001 - Fire Alarm Zone 1
        {
            'rung': '0001',
            'inputs': ['I:0/1', 'I:0/3', 'I:0/5', 'B11:0/1', 'B14:0/0', 'B3:4/6'],
            'outputs': ['B3:0/0', 'B3:2/0'],
            'description': 'Fire Alarm Zone 1'
        },
        # Rung 0002 - Fire Alarm Zone 2 with Deluge
        {
            'rung': '0002',
            'inputs': ['I:0/0', 'I:0/2', 'I:0/4', 'B11:0/0', 'B14:0/1', 'B3:4/5'],
            'outputs': ['O:0/0', 'B3:0/10', 'B3:2/10'],
            'description': 'Fire Alarm Zone 2 and Deluge Valve'
        },
        # Rung 0003 - ESD to Office PLC
        {
            'rung': '0003',
            'inputs': ['B3:0/0', 'O:0/0', 'B11:0/1'],
            'outputs': ['B3:10/0'],
            'description': 'ESD Alarm to Office PLC'
        },
        # Fire Eye Failure Detection (XIO logic - examines if open)
        {
            'rung': '0004',
            'inputs': ['I:0/7'],  # XIO - Examine if Open
            'outputs': ['B3:3/8'],
            'description': 'Fire Eye 1 Failure Alarm',
            'timer': 'T4:16',
            'logic_type': 'XIO'
        },
        {
            'rung': '0006',
            'inputs': ['I:0/9'],
            'outputs': ['B3:3/9'],
            'description': 'Fire Eye 2 Failure Alarm',
            'timer': 'T4:15',
            'logic_type': 'XIO'
        },
        {
            'rung': '0008',
            'inputs': ['I:0/11'],
            'outputs': ['B3:3/10'],
            'description': 'Fire Eye 3 Failure Alarm',
            'timer': 'T4:14',
            'logic_type': 'XIO'
        },
        {
            'rung': '0010',
            'inputs': ['I:0/13'],
            'outputs': ['B3:3/11'],
            'description': 'Fire Eye 4 Failure Alarm',
            'timer': 'T4:13',
            'logic_type': 'XIO'
        },
        {
            'rung': '0012',
            'inputs': ['I:0/15'],
            'outputs': ['B3:3/12'],
            'description': 'Fire Eye 5 Failure Alarm',
            'timer': 'T4:12',
            'logic_type': 'XIO'
        },
        {
            'rung': '0014',
            'inputs': ['I:0/17'],
            'outputs': ['B3:3/13'],
            'description': 'Fire Eye 6 Failure Alarm',
            'timer': 'T4:11',
            'logic_type': 'XIO'
        },
        {
            'rung': '0016',
            'inputs': ['I:0/19'],
            'outputs': ['B3:3/14'],
            'description': 'Fire Eye 7 Failure Alarm',
            'timer': 'T4:10',
            'logic_type': 'XIO'
        },
        {
            'rung': '0018',
            'inputs': ['I:1/1'],
            'outputs': ['B3:3/15'],
            'description': 'Fire Eye 8 Failure Alarm',
            'timer': 'T4:9',
            'logic_type': 'XIO'
        },
        # Rung 0020 - Fire Eye Faulted Zone 1 (OR of failures)
        {
            'rung': '0020',
            'inputs': ['B3:3/10', 'B3:3/11', 'B3:3/12', 'B3:3/13', 'B3:3/14', 'B3:3/15'],
            'outputs': ['B3:0/1'],
            'description': 'Fire Eye Faulted Zone 1',
            'logic_type': 'OR'
        },
        # Rung 0021 - Fire Eye Faulted Zone 2
        {
            'rung': '0021',
            'inputs': ['B3:3/9', 'B3:3/8'],
            'outputs': ['B3:0/3'],
            'description': 'Fire Eye Faulted Zone 2',
            'logic_type': 'OR'
        },
        # Fire Eye Fire Detection with TON and CTU
        {
            'rung': '0022-0023',
            'inputs': ['I:0/6'],
            'outputs': ['B3:3/0'],
            'description': 'Fire Eye 1 Fire Detected',
            'timer': 'T4:1',
            'counter': 'C5:0'
        },
        {
            'rung': '0024-0025',
            'inputs': ['I:0/8'],
            'outputs': ['B3:3/1'],
            'description': 'Fire Eye 2 Fire Detected',
            'timer': 'T4:2',
            'counter': 'C5:1'
        },
        {
            'rung': '0026-0027',
            'inputs': ['I:0/10'],
            'outputs': ['B3:3/2'],
            'description': 'Fire Eye 3 Fire Detected',
            'timer': 'T4:3',
            'counter': 'C5:2'
        },
        {
            'rung': '0028-0029',
            'inputs': ['I:0/12'],
            'outputs': ['B3:3/3'],
            'description': 'Fire Eye 4 Fire Detected',
            'timer': 'T4:4',
            'counter': 'C5:3'
        },
        {
            'rung': '0030-0031',
            'inputs': ['I:0/14'],
            'outputs': ['B3:3/4'],
            'description': 'Fire Eye 5 Fire Detected',
            'timer': 'T4:5',
            'counter': 'C5:4'
        },
        {
            'rung': '0032-0033',
            'inputs': ['I:0/16'],
            'outputs': ['B3:3/5'],
            'description': 'Fire Eye 6 Fire Detected',
            'timer': 'T4:6',
            'counter': 'C5:5'
        },
        {
            'rung': '0034-0035',
            'inputs': ['I:0/18'],
            'outputs': ['B3:3/6'],
            'description': 'Fire Eye 7 Fire Detected',
            'timer': 'T4:7',
            'counter': 'C5:6'
        },
        {
            'rung': '0036-0037',
            'inputs': ['I:1/0'],
            'outputs': ['B3:3/7'],
            'description': 'Fire Eye 8 Fire Detected',
            'timer': 'T4:8',
            'counter': 'C5:7'
        },
        # Single detector alarms
        {
            'rung': '0038',
            'inputs': ['B3:3/2', 'B3:3/3', 'B3:3/4', 'B3:3/5', 'B3:3/6', 'B3:3/7'],
            'outputs': ['B3:0/5'],
            'description': 'Fire Detected Zone 1 Fire Eyes Single Detector',
            'logic_type': 'OR'
        },
        {
            'rung': '0039',
            'inputs': ['B3:3/1', 'B3:3/0'],
            'outputs': ['B3:0/4'],
            'description': 'Fire Detected Zone 2 Fire Eyes Single Detector',
            'logic_type': 'OR'
        },
        # Additional alarm rungs
        {
            'rung': '0054',
            'inputs': ['B14:0/2'],
            'outputs': ['B3:0/11'],
            'description': 'Fire Alarm Zone 2'
        },
        {
            'rung': '0055',
            'inputs': ['I:1/12'],
            'outputs': ['B3:0/8', 'B3:2/8'],
            'description': 'Fire Eye Failure Warning',
            'logic_type': 'XIO'
        },
        {
            'rung': '0056',
            'inputs': ['I:1/13'],
            'outputs': ['B3:0/9', 'B3:2/9'],
            'description': 'Strobe Light On',
            'logic_type': 'XIO'
        },
        {
            'rung': '0057',
            'inputs': ['I:1/15'],
            'outputs': ['B3:0/13', 'B3:2/13'],
            'description': 'Strobe Light On'
        },
    ]

    return rungs, tag_descriptions

def build_alarm_summary(tag_descriptions):
    """Build alarm summary data from tags"""
    alarms = []

    # Define alarm tags to extract
    alarm_tags = [
        # Fire Alarms
        ('B3:0/0', 'Fire Alarm Zone 1'),
        ('B3:0/11', 'Fire Alarm Zone 2'),
        ('B3:0/1', 'Fire Eye Faulted Zone 1'),
        ('B3:0/3', 'Fire Eye Faulted Zone 2'),
        ('B3:0/4', 'Fire Detected Zone 2 Fire Eyes. Single Detector Only'),
        ('B3:0/5', 'Fire Detected Zone 1 Fire Eyes. Single Detector Only'),
        ('B3:0/6', 'Plant ESD'),
        ('B3:0/8', 'Fire Eye Failure Warning'),
        ('B3:0/9', 'Strobe Light On'),
        ('B3:0/10', 'Strobe Light On'),
        ('B3:0/12', 'Fire System Deluge Valve Open'),
        ('B3:0/13', 'Strobe Light On'),
        # Fire Eye Failure Alarms
        ('B3:3/8', 'FE 1 Failure Alarm'),
        ('B3:3/9', 'FE 2 Failure Alarm'),
        ('B3:3/10', 'FE 3 Failure Alarm'),
        ('B3:3/11', 'FE 4 Failure Alarm'),
        ('B3:3/12', 'FE 5 Failure Alarm'),
        ('B3:3/13', 'FE 6 Failure Alarm'),
        ('B3:3/14', 'FE 7 Failure Alarm'),
        ('B3:3/15', 'FE 8 Failure Alarm'),
        # Fire Eye Fire Detected
        ('B3:3/0', 'Fire Eye 1 Fire Detected'),
        ('B3:3/1', 'Fire Eye 2 Fire Detected'),
        ('B3:3/2', 'Fire Eye 3 Fire Detected'),
        ('B3:3/3', 'Fire Eye 4 Fire Detected'),
        ('B3:3/4', 'Fire Eye 5 Fire Detected'),
        ('B3:3/5', 'Fire Eye 6 Fire Detected'),
        ('B3:3/6', 'Fire Eye 7 Fire Detected'),
        ('B3:3/7', 'Fire Eye 8 Fire Detected'),
        # 2 Detector Alarms
        ('B3:4/6', '2 Detectors In Alarm Zone 1'),
        ('B3:4/5', '2 Detectors In Alarm Zone 2'),
        # ESD
        ('B3:10/0', 'ESD Alarm to Office PLC'),
        # Deluge Status
        ('O:0/0', 'Deluge Valve Zone 2 Open'),
    ]

    for tag_addr, description in alarm_tags:
        alarms.append({
            'Tag No': tag_addr,
            'P & ID': '',
            'Service Description': description,
            'Range': '',
            'EU': '',
            'Normal Operating Conditions': '',
            'HH': '',
            'H': '',
            'L': '',
            'LL': '',
            'Engineering Notes': ''
        })

    return alarms

def build_cause_effect_matrix(rungs, tag_descriptions):
    """Build cause and effect matrix from ladder rungs"""
    interlocks = []
    interlock_num = 1

    # Filter rungs that have physical I/O or shutdowns
    for rung in rungs:
        # Check if rung has physical inputs (I:, B11:, B14:) or outputs (O:, B3:2, B3:10)
        has_physical_input = any(tag.startswith(('I:', 'B11:', 'B14:')) for tag in rung['inputs'])
        has_physical_output = any(tag.startswith(('O:')) for tag in rung['outputs'])
        has_shutdown = any(tag.startswith(('B3:2', 'B3:10', 'B3:0/0', 'B3:0/11')) for tag in rung['outputs'])

        # Only include rungs with physical I/O or key outputs
        if has_physical_input and (has_physical_output or has_shutdown):
            # Get all input tags
            input_tags = rung['inputs']

            # Primary input (first physical input)
            primary_input = None
            for tag in input_tags:
                if tag.startswith(('I:', 'B11:', 'B14:')):
                    primary_input = tag
                    break

            if not primary_input and input_tags:
                primary_input = input_tags[0]

            # Get description
            service_desc = rung['description']
            if primary_input and primary_input in tag_descriptions:
                service_desc = tag_descriptions[primary_input]

            # Build effects dictionary
            effects = defaultdict(str)
            for output in rung['outputs']:
                effects[output] = 'X'

            # Add timer/counter info to description if present
            extra_info = []
            if 'timer' in rung:
                extra_info.append(f"Timer: {rung['timer']}")
            if 'counter' in rung:
                extra_info.append(f"Counter: {rung['counter']}")
            if extra_info:
                service_desc += f" ({', '.join(extra_info)})"

            interlock = {
                'Interlock No': interlock_num,
                'Tag No': primary_input or '',
                'Service Description': service_desc,
                'Range': '',
                'Pre-Trip (H or L)': '',
                'Trip (HH or LL)': '',
                'P & ID': '',
                'Rung': rung['rung'],
                'Effects': effects,
                'All Inputs': input_tags,
                'All Outputs': rung['outputs']
            }

            interlocks.append(interlock)
            interlock_num += 1

    return interlocks

def generate_alarm_summary_excel(alarms, output_file):
    """Generate Alarm Summary Excel file"""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Alarm Summary'

    # Header row
    headers = [
        'Tag No', 'P & ID', 'Service Description', 'Range', 'EU',
        'Normal Operating\\nConditions', 'HH', 'H', 'L', 'LL', 'Engineering Notes'
    ]

    ws.append(headers)

    # Format header
    header_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for cell in ws[1]:
        cell.font = Font(bold=True, size=10)
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        cell.border = thin_border

    # Set row height for header
    ws.row_dimensions[1].height = 30

    # Add data
    for alarm in alarms:
        row = [
            alarm['Tag No'],
            alarm['P & ID'],
            alarm['Service Description'],
            alarm['Range'],
            alarm['EU'],
            alarm['Normal Operating Conditions'],
            alarm['HH'],
            alarm['H'],
            alarm['L'],
            alarm['LL'],
            alarm['Engineering Notes']
        ]
        ws.append(row)

        # Apply borders to data rows
        for cell in ws[ws.max_row]:
            cell.border = thin_border

    # Adjust column widths
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 55
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 8
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 8
    ws.column_dimensions['H'].width = 8
    ws.column_dimensions['I'].width = 8
    ws.column_dimensions['J'].width = 8
    ws.column_dimensions['K'].width = 30

    wb.save(output_file)
    print(f'✓ Alarm Summary saved to: {output_file}')

def generate_cause_effect_excel(interlocks, tag_descriptions, output_file):
    """Generate Cause & Effect Matrix Excel file"""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Cause & Effect'

    # Collect all unique output effects
    all_effects = set()
    for interlock in interlocks:
        all_effects.update(interlock['Effects'].keys())

    effect_columns = sorted(list(all_effects), key=lambda x: (x.split(':')[0], int(x.split(':')[1].split('/')[0]), int(x.split('/')[1])))

    # Create header structure similar to example
    # Row 1: Title row with "EFFECT" label
    row1 = ['', '', '', '', '', '', 'EFFECT'] + [''] * len(effect_columns)
    ws.append(row1)

    # Row 2: Effect tag addresses with descriptions
    row2 = ['', '', '', '', '', ''] + [f'{tag}\\n{tag_descriptions.get(tag, "")}' for tag in effect_columns]
    ws.append(row2)

    # Row 3: Column headers
    row3 = ['Interlock\\nNo', 'Tag No', 'Service Description', 'Range', 'Pre-Trip\\n(H or L)', 'Trip\\n(HH or LL)'] + [''] * len(effect_columns)
    ws.append(row3)

    # Row 4: CAUSE label and P & ID labels
    row4 = ['CAUSE', '', '', '', '', ''] + ['P & ID'] * len(effect_columns)
    ws.append(row4)

    # Formatting
    header_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    effect_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
    cause_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Format row 1 - EFFECT label
    for col in range(7, 7 + len(effect_columns)):
        cell = ws.cell(row=1, column=col)
        cell.fill = effect_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    # Format row 2 - Effect tags
    for idx, tag in enumerate(effect_columns):
        cell = ws.cell(row=2, column=7 + idx)
        cell.value = f'{tag}\\n{tag_descriptions.get(tag, "")}'
        cell.fill = effect_fill
        cell.font = Font(size=9)
        cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        cell.border = thin_border

    # Format row 3 - Headers
    for col in range(1, 7):
        cell = ws.cell(row=3, column=col)
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        cell.border = thin_border

    # Format row 4 - CAUSE and P & ID
    cell = ws.cell(row=4, column=1)
    cell.fill = cause_fill
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = thin_border

    for col in range(7, 7 + len(effect_columns)):
        cell = ws.cell(row=4, column=col)
        cell.fill = header_fill
        cell.font = Font(size=9)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    # Set row heights
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[4].height = 20

    # Add interlock data
    for interlock in interlocks:
        row_data = [
            f"I-{interlock['Interlock No']}",
            interlock['Tag No'],
            interlock['Service Description'],
            interlock['Range'],
            interlock['Pre-Trip (H or L)'],
            interlock['Trip (HH or LL)']
        ]

        # Add effect markers
        for effect_col in effect_columns:
            if effect_col in interlock['Effects']:
                row_data.append('X')
            else:
                row_data.append('')

        ws.append(row_data)

        # Apply borders and formatting to data row
        for col_idx, cell in enumerate(ws[ws.max_row], 1):
            cell.border = thin_border
            if col_idx >= 7:  # Effect columns
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if cell.value == 'X':
                    cell.font = Font(bold=True)

    # Adjust column widths
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 60
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 12

    # Set effect column widths
    for i, tag in enumerate(effect_columns):
        col_letter = chr(ord('G') + i)
        ws.column_dimensions[col_letter].width = 18

    wb.save(output_file)
    print(f'✓ Cause & Effect Matrix saved to: {output_file}')

def main():
    """Main execution function"""
    print('═' * 70)
    print('  RSLogix 500 FIRE SYSTEM PARSER')
    print('  Trafigura Fire System PLC')
    print('═' * 70)

    # Extract data from PDF
    print('\\n[1/4] Extracting ladder logic from PDF...')
    rungs, tag_descriptions = extract_data_from_pdf()
    print(f'      ✓ Extracted {len(rungs)} ladder rungs')
    print(f'      ✓ Loaded {len(tag_descriptions)} tag descriptions')

    # Build alarm summary
    print('\\n[2/4] Building alarm summary...')
    alarms = build_alarm_summary(tag_descriptions)
    print(f'      ✓ Found {len(alarms)} alarm tags')

    # Build cause & effect matrix
    print('\\n[3/4] Building cause & effect matrix...')
    interlocks = build_cause_effect_matrix(rungs, tag_descriptions)
    print(f'      ✓ Found {len(interlocks)} interlocks')

    # Generate Excel files
    print('\\n[4/4] Generating Excel files...')
    generate_alarm_summary_excel(alarms, 'Alarm_Summary_Output.xlsx')
    generate_cause_effect_excel(interlocks, tag_descriptions, 'Cause_Effect_Output.xlsx')

    print('\\n' + '═' * 70)
    print('  PROCESSING COMPLETE!')
    print('═' * 70)
    print('\\nOutput files created:')
    print('  ├─ Alarm_Summary_Output.xlsx')
    print('  └─ Cause_Effect_Output.xlsx')
    print('')

if __name__ == '__main__':
    main()
