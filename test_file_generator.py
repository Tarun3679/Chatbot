#!/usr/bin/env python3
"""
LibreOffice OOM Test Suite - ZERO DEPENDENCIES

This script creates test Office files and runs stress tests using ONLY
Python standard library. No pip packages required!

Creates valid PPTX/XLSX files using zipfile + XML (Office Open XML format).

Usage:
    python3 test_suite_standalone.py                     # Run full test
    python3 test_suite_standalone.py --generate-only     # Just create files
    python3 test_suite_standalone.py --test-only         # Test existing files
    python3 test_suite_standalone.py --level heavy       # More stress

Requirements:
    - Python 3.6+ (standard library only - NO PIP PACKAGES)
    - Linux operating system
    - LibreOffice installed
"""

import argparse
import io
import json
import os
import random
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Optional
from dataclasses import dataclass

# =============================================================================
# CONSOLE OUTPUT 
# =============================================================================

class Colors:
    RED = '\033[0;31m'
    GREEN = '\033[0;32m'
    YELLOW = '\033[1;33m'
    BLUE = '\033[0;34m'
    CYAN = '\033[0;36m'
    NC = '\033[0m'

if not sys.stdout.isatty():
    Colors.RED = Colors.GREEN = Colors.YELLOW = Colors.BLUE = Colors.CYAN = Colors.NC = ''

def print_header(text: str):
    print(f"\n{Colors.BLUE}{'=' * 65}{Colors.NC}")
    print(f"{Colors.BLUE}  {text}{Colors.NC}")
    print(f"{Colors.BLUE}{'=' * 65}{Colors.NC}\n")

def print_ok(text: str):
    print(f"{Colors.GREEN}[✓]{Colors.NC} {text}")

def print_warn(text: str):
    print(f"{Colors.YELLOW}[!]{Colors.NC} {text}")

def print_err(text: str):
    print(f"{Colors.RED}[✗]{Colors.NC} {text}")

def print_info(text: str):
    print(f"{Colors.CYAN}[i]{Colors.NC} {text}")

# =============================================================================
# PPTX GENERATOR (Pure Python - ZIP + XML)
# =============================================================================

def random_text(words: int = 20) -> str:
    """Generate random filler text."""
    word_list = [
        'lorem', 'ipsum', 'dolor', 'sit', 'amet', 'consectetur', 'adipiscing',
        'elit', 'sed', 'do', 'eiusmod', 'tempor', 'incididunt', 'ut', 'labore',
        'et', 'dolore', 'magna', 'aliqua', 'enim', 'ad', 'minim', 'veniam',
        'quis', 'nostrud', 'exercitation', 'ullamco', 'laboris', 'report',
        'quarterly', 'revenue', 'growth', 'market', 'sales', 'data', 'analysis'
    ]
    return ' '.join(random.choices(word_list, k=words)).capitalize() + '.'


def create_pptx(num_slides: int = 10, text_blocks: int = 3) -> bytes:
    """
    Create a valid PPTX file using only standard library.
    PPTX = ZIP file containing XML files (Office Open XML format).
    """
    buf = io.BytesIO()
    
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        # [Content_Types].xml
        slide_overrides = '\n'.join(
            f'  <Override PartName="/ppt/slides/slide{i+1}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
            for i in range(num_slides)
        )
        content_types = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
{slide_overrides}
</Types>'''
        zf.writestr('[Content_Types].xml', content_types)
        
        # _rels/.rels
        zf.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>''')
        
        # ppt/presentation.xml
        slide_ids = '\n'.join(f'    <p:sldId id="{256+i}" r:id="rId{i+2}"/>' for i in range(num_slides))
        zf.writestr('ppt/presentation.xml', f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>
{slide_ids}
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>''')
        
        # ppt/_rels/presentation.xml.rels
        slide_rels = '\n'.join(
            f'  <Relationship Id="rId{i+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>'
            for i in range(num_slides)
        )
        zf.writestr('ppt/_rels/presentation.xml.rels', f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
{slide_rels}
</Relationships>''')
        
        # ppt/slideMasters/slideMaster1.xml
        zf.writestr('ppt/slideMasters/slideMaster1.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>''')
        
        # ppt/slideMasters/_rels/slideMaster1.xml.rels
        zf.writestr('ppt/slideMasters/_rels/slideMaster1.xml.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>''')
        
        # ppt/slideLayouts/slideLayout1.xml
        zf.writestr('ppt/slideLayouts/slideLayout1.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>''')
        
        # ppt/slideLayouts/_rels/slideLayout1.xml.rels
        zf.writestr('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>''')
        
        # Create each slide
        for slide_num in range(num_slides):
            # Generate text content for this slide
            text_shapes = []
            for block_idx in range(text_blocks):
                x_pos = 500000 + (block_idx % 2) * 4000000
                y_pos = 1500000 + (block_idx // 2) * 1500000
                text_content = random_text(30 + random.randint(0, 20))
                
                text_shapes.append(f'''
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="{block_idx + 2}" name="TextBox {block_idx + 1}"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="{x_pos}" y="{y_pos}"/><a:ext cx="3500000" cy="1000000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square"/>
        <a:lstStyle/>
        <a:p><a:r><a:rPr lang="en-US" sz="1200"/><a:t>{text_content}</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>''')
            
            slide_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="500000" y="300000"/><a:ext cx="8000000" cy="800000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:r><a:rPr lang="en-US" sz="2800" b="1"/><a:t>Slide {slide_num + 1}: {random_text(5)}</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      {''.join(text_shapes)}
    </p:spTree>
  </p:cSld>
</p:sld>'''
            zf.writestr(f'ppt/slides/slide{slide_num + 1}.xml', slide_xml)
            
            # Slide rels
            zf.writestr(f'ppt/slides/_rels/slide{slide_num + 1}.xml.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>''')
    
    return buf.getvalue()


# =============================================================================
# XLSX GENERATOR (Pure Python - ZIP + XML)
# =============================================================================

def create_xlsx(num_sheets: int = 3, rows_per_sheet: int = 100, cols: int = 10) -> bytes:
    """
    Create a valid XLSX file using only standard library.
    XLSX = ZIP file containing XML files (Office Open XML format).
    """
    buf = io.BytesIO()
    
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        # [Content_Types].xml
        sheet_overrides = '\n'.join(
            f'  <Override PartName="/xl/worksheets/sheet{i+1}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(num_sheets)
        )
        zf.writestr('[Content_Types].xml', f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
{sheet_overrides}
</Types>''')
        
        # _rels/.rels
        zf.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>''')
        
        # xl/workbook.xml
        sheet_defs = '\n'.join(
            f'    <sheet name="Sheet{i+1}" sheetId="{i+1}" r:id="rId{i+1}"/>'
            for i in range(num_sheets)
        )
        zf.writestr('xl/workbook.xml', f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
{sheet_defs}
  </sheets>
</workbook>''')
        
        # xl/_rels/workbook.xml.rels
        sheet_rels = '\n'.join(
            f'  <Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>'
            for i in range(num_sheets)
        )
        zf.writestr('xl/_rels/workbook.xml.rels', f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{sheet_rels}
  <Relationship Id="rId{num_sheets+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId{num_sheets+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>''')
        
        # xl/styles.xml (minimal)
        zf.writestr('xl/styles.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="1"><xf/></cellXfs>
</styleSheet>''')
        
        # Collect all shared strings
        shared_strings = []
        shared_string_map = {}
        
        def get_string_index(s: str) -> int:
            if s not in shared_string_map:
                shared_string_map[s] = len(shared_strings)
                shared_strings.append(s)
            return shared_string_map[s]
        
        # Create worksheets
        for sheet_idx in range(num_sheets):
            rows_xml = []
            
            # Header row
            header_cells = []
            for c in range(cols):
                col_letter = chr(65 + c) if c < 26 else f"A{chr(65 + c - 26)}"
                header_text = f"Column_{col_letter}"
                str_idx = get_string_index(header_text)
                header_cells.append(f'<c r="{col_letter}1" t="s"><v>{str_idx}</v></c>')
            rows_xml.append(f'    <row r="1">{"".join(header_cells)}</row>')
            
            # Data rows
            for row_num in range(2, rows_per_sheet + 2):
                cells = []
                for c in range(cols):
                    col_letter = chr(65 + c) if c < 26 else f"A{chr(65 + c - 26)}"
                    cell_ref = f"{col_letter}{row_num}"
                    
                    # Vary data types
                    if c % 3 == 0:
                        # Number
                        value = random.randint(1, 10000)
                        cells.append(f'<c r="{cell_ref}"><v>{value}</v></c>')
                    elif c % 3 == 1:
                        # Decimal
                        value = round(random.uniform(0, 1000), 2)
                        cells.append(f'<c r="{cell_ref}"><v>{value}</v></c>')
                    else:
                        # Text
                        text = random_text(3)
                        str_idx = get_string_index(text)
                        cells.append(f'<c r="{cell_ref}" t="s"><v>{str_idx}</v></c>')
                
                rows_xml.append(f'    <row r="{row_num}">{"".join(cells)}</row>')
            
            # Write worksheet
            last_col = chr(65 + cols - 1) if cols <= 26 else f"A{chr(65 + cols - 27)}"
            zf.writestr(f'xl/worksheets/sheet{sheet_idx + 1}.xml', f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:{last_col}{rows_per_sheet + 1}"/>
  <sheetData>
{"chr(10)".join(rows_xml)}
  </sheetData>
</worksheet>''')
        
        # xl/sharedStrings.xml
        ss_items = '\n'.join(f'  <si><t>{s}</t></si>' for s in shared_strings)
        zf.writestr('xl/sharedStrings.xml', f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">
{ss_items}
</sst>''')
    
    return buf.getvalue()


# =============================================================================
# SYSTEM UTILITIES
# =============================================================================

def get_memory_info() -> Dict[str, float]:
    """Get memory info from /proc/meminfo (Linux)."""
    try:
        with open('/proc/meminfo', 'r') as f:
            info = {}
            for line in f:
                parts = line.split()
                if len(parts) >= 2:
                    info[parts[0].rstrip(':')] = int(parts[1])
            
            total = info.get('MemTotal', 0) / 1024
            avail = info.get('MemAvailable', info.get('MemFree', 0)) / 1024
            used = total - avail
            
            return {
                'total_mb': total,
                'used_mb': used,
                'available_mb': avail,
                'percent': round(used / total * 100, 1) if total > 0 else 0
            }
    except:
        return {'total_mb': 0, 'used_mb': 0, 'available_mb': 0, 'percent': 0}


def find_libreoffice() -> Optional[str]:
    """Find LibreOffice executable."""
    paths = [
        '/usr/bin/soffice',
        '/usr/bin/libreoffice', 
        '/usr/local/bin/soffice',
        '/opt/libreoffice/program/soffice',
        shutil.which('soffice'),
        shutil.which('libreoffice'),
    ]
    for p in paths:
        if p and os.path.isfile(p) and os.access(p, os.X_OK):
            return p
    return None


def get_optimized_env() -> Dict[str, str]:
    """Environment variables to reduce LibreOffice memory usage."""
    env = os.environ.copy()
    env.update({
        'SAL_DISABLE_OPENCL': '1',
        'SAL_DISABLEGL': '1',
        'SAL_DISABLESKIA': '1',
        'SAL_USE_VCLPLUGIN': 'gen',
        'SAL_NO_CRASHREPORT': '1',
        'JAVA_TOOL_OPTIONS': '-Xmx256m',
    })
    return env


# =============================================================================
# CONVERSION & TESTING
# =============================================================================

@dataclass
class ConversionResult:
    filename: str
    input_size_kb: float
    output_size_kb: float
    status: str
    duration_sec: float
    error: str = ""


def convert_to_pdf(
    input_path: Path,
    soffice: str,
    timeout: int = 120,
    use_unique_profile: bool = True,
    use_optimized_env: bool = True
) -> ConversionResult:
    """Convert a file to PDF using LibreOffice."""
    
    input_size = input_path.stat().st_size / 1024
    start = time.time()
    
    with tempfile.TemporaryDirectory(prefix='lo_') as tmp:
        tmp_path = Path(tmp)
        
        # Copy input to temp
        tmp_input = tmp_path / input_path.name
        shutil.copy(input_path, tmp_input)
        
        # Build command
        cmd = [
            soffice,
            '--headless',
            '--invisible',
            '--nodefault',
            '--nofirststartwizard',
            '--nolockcheck',
            '--nologo',
            '--norestore',
        ]
        
        if use_unique_profile:
            profile = tmp_path / f'profile_{uuid.uuid4().hex[:8]}'
            profile.mkdir()
            cmd.append(f'-env:UserInstallation=file://{profile}')
        
        cmd.extend(['--convert-to', 'pdf', '--outdir', str(tmp_path), str(tmp_input)])
        
        env = get_optimized_env() if use_optimized_env else os.environ.copy()
        
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, env=env)
            duration = time.time() - start
            
            # Find output PDF
            pdfs = list(tmp_path.glob('*.pdf'))
            if pdfs:
                output_size = pdfs[0].stat().st_size / 1024
                return ConversionResult(
                    filename=input_path.name,
                    input_size_kb=input_size,
                    output_size_kb=output_size,
                    status='success',
                    duration_sec=round(duration, 2)
                )
            else:
                return ConversionResult(
                    filename=input_path.name,
                    input_size_kb=input_size,
                    output_size_kb=0,
                    status='failed',
                    duration_sec=round(duration, 2),
                    error=result.stderr[:200] if result.stderr else 'No PDF output'
                )
        
        except subprocess.TimeoutExpired:
            return ConversionResult(
                filename=input_path.name,
                input_size_kb=input_size,
                output_size_kb=0,
                status='timeout',
                duration_sec=timeout,
                error=f'Timed out after {timeout}s'
            )
        except Exception as e:
            return ConversionResult(
                filename=input_path.name,
                input_size_kb=input_size,
                output_size_kb=0,
                status='error',
                duration_sec=time.time() - start,
                error=str(e)[:200]
            )


# =============================================================================
# TEST CONFIGURATIONS
# =============================================================================

TEST_CONFIGS = {
    'light': {
        'pptx': [
            {'name': 'pptx_small.pptx', 'slides': 5, 'text_blocks': 2},
            {'name': 'pptx_medium.pptx', 'slides': 10, 'text_blocks': 3},
        ],
        'xlsx': [
            {'name': 'xlsx_small.xlsx', 'sheets': 2, 'rows': 100, 'cols': 8},
            {'name': 'xlsx_medium.xlsx', 'sheets': 3, 'rows': 500, 'cols': 10},
        ]
    },
    'medium': {
        'pptx': [
            {'name': 'pptx_01.pptx', 'slides': 15, 'text_blocks': 4},
            {'name': 'pptx_02.pptx', 'slides': 25, 'text_blocks': 5},
            {'name': 'pptx_03.pptx', 'slides': 30, 'text_blocks': 4},
        ],
        'xlsx': [
            {'name': 'xlsx_01.xlsx', 'sheets': 3, 'rows': 1000, 'cols': 12},
            {'name': 'xlsx_02.xlsx', 'sheets': 5, 'rows': 2000, 'cols': 15},
            {'name': 'xlsx_03.xlsx', 'sheets': 4, 'rows': 3000, 'cols': 10},
        ]
    },
    'heavy': {
        'pptx': [
            {'name': 'pptx_heavy_01.pptx', 'slides': 50, 'text_blocks': 6},
            {'name': 'pptx_heavy_02.pptx', 'slides': 75, 'text_blocks': 5},
            {'name': 'pptx_heavy_03.pptx', 'slides': 60, 'text_blocks': 7},
        ],
        'xlsx': [
            {'name': 'xlsx_heavy_01.xlsx', 'sheets': 8, 'rows': 5000, 'cols': 20},
            {'name': 'xlsx_heavy_02.xlsx', 'sheets': 10, 'rows': 8000, 'cols': 15},
            {'name': 'xlsx_heavy_03.xlsx', 'sheets': 6, 'rows': 10000, 'cols': 25},
        ]
    },
}


# =============================================================================
# MAIN FUNCTIONS
# =============================================================================

def generate_test_files(output_dir: Path, level: str = 'medium', verbose: bool = True) -> List[Path]:
    """Generate test files for the specified stress level."""
    
    config = TEST_CONFIGS.get(level, TEST_CONFIGS['medium'])
    output_dir.mkdir(parents=True, exist_ok=True)
    
    created = []
    
    if verbose:
        print_info(f"Generating {level} test files in {output_dir}")
    
    # Create PPTX files
    for cfg in config['pptx']:
        if verbose:
            print(f"  Creating {cfg['name']} ({cfg['slides']} slides)...", end=' ', flush=True)
        
        data = create_pptx(num_slides=cfg['slides'], text_blocks=cfg['text_blocks'])
        path = output_dir / cfg['name']
        path.write_bytes(data)
        created.append(path)
        
        if verbose:
            print(f"{len(data) / 1024:.1f} KB")
    
    # Create XLSX files
    for cfg in config['xlsx']:
        if verbose:
            print(f"  Creating {cfg['name']} ({cfg['rows']} rows x {cfg['sheets']} sheets)...", end=' ', flush=True)
        
        data = create_xlsx(num_sheets=cfg['sheets'], rows_per_sheet=cfg['rows'], cols=cfg['cols'])
        path = output_dir / cfg['name']
        path.write_bytes(data)
        created.append(path)
        
        if verbose:
            print(f"{len(data) / 1024:.1f} KB")
    
    return created


def run_stress_test(
    test_dir: Path,
    concurrent: int = 4,
    rounds: int = 2,
    timeout: int = 120,
    use_fixes: bool = True,
    verbose: bool = True
) -> Dict:
    """Run stress test on files in test_dir."""
    
    soffice = find_libreoffice()
    if not soffice:
        print_err("LibreOffice not found!")
        sys.exit(1)
    
    # Find test files
    files = list(test_dir.glob('*.pptx')) + list(test_dir.glob('*.xlsx'))
    if not files:
        print_err(f"No test files found in {test_dir}")
        sys.exit(1)
    
    if verbose:
        print_info(f"Found {len(files)} test files")
        print_info(f"Running {rounds} round(s) with {concurrent} concurrent conversions")
        print_info(f"Using fixes: {use_fixes}")
        print()
    
    # Build task list
    tasks = [(f, r) for r in range(rounds) for f in files]
    results: List[ConversionResult] = []
    
    peak_memory = 0
    start_time = time.time()
    
    with ThreadPoolExecutor(max_workers=concurrent) as executor:
        futures = {
            executor.submit(
                convert_to_pdf, f, soffice, timeout,
                use_unique_profile=use_fixes,
                use_optimized_env=use_fixes
            ): (f, r) for f, r in tasks
        }
        
        done = 0
        for future in as_completed(futures):
            done += 1
            f, r = futures[future]
            
            # Check memory
            mem = get_memory_info()
            if mem['used_mb'] > peak_memory:
                peak_memory = mem['used_mb']
            
            try:
                result = future.result()
                results.append(result)
                
                if verbose:
                    icon = '✓' if result.status == 'success' else '✗'
                    print(f"[{done}/{len(tasks)}] {icon} {result.filename} "
                          f"({result.duration_sec}s) | Mem: {mem['used_mb']/1024:.1f}GB ({mem['percent']}%)")
            except Exception as e:
                if verbose:
                    print(f"[{done}/{len(tasks)}] ✗ {f.name}: {e}")
    
    total_time = time.time() - start_time
    
    # Calculate summary
    success = sum(1 for r in results if r.status == 'success')
    failed = sum(1 for r in results if r.status == 'failed')
    timeouts = sum(1 for r in results if r.status == 'timeout')
    
    return {
        'total': len(results),
        'success': success,
        'failed': failed,
        'timeouts': timeouts,
        'success_rate': round(success / len(results) * 100, 1) if results else 0,
        'total_time_sec': round(total_time, 1),
        'avg_duration_sec': round(sum(r.duration_sec for r in results) / len(results), 2) if results else 0,
        'peak_memory_gb': round(peak_memory / 1024, 2),
        'results': [r.__dict__ for r in results]
    }


def main():
    parser = argparse.ArgumentParser(
        description='LibreOffice OOM Test Suite (Zero Dependencies)',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument('--output-dir', '-o', default='./test_files',
                        help='Output directory for test files')
    parser.add_argument('--level', '-l', choices=['light', 'medium', 'heavy'], default='medium',
                        help='Stress level')
    parser.add_argument('--concurrent', '-c', type=int, default=4,
                        help='Concurrent conversions')
    parser.add_argument('--rounds', '-r', type=int, default=2,
                        help='Number of test rounds')
    parser.add_argument('--timeout', '-t', type=int, default=120,
                        help='Timeout per conversion (seconds)')
    parser.add_argument('--generate-only', action='store_true',
                        help='Only generate test files')
    parser.add_argument('--test-only', action='store_true',
                        help='Only run tests (files must exist)')
    parser.add_argument('--compare', action='store_true',
                        help='Run both with and without fixes and compare')
    parser.add_argument('--quiet', '-q', action='store_true',
                        help='Less output')
    
    args = parser.parse_args()
    output_dir = Path(args.output_dir)
    verbose = not args.quiet
    
    # Header
    if verbose:
        print_header("LibreOffice OOM Test Suite (Zero Dependencies)")
    
    # Check LibreOffice
    soffice = find_libreoffice()
    if soffice:
        print_ok(f"LibreOffice found: {soffice}")
    else:
        print_err("LibreOffice not found! Please install it.")
        sys.exit(1)
    
    # Show system info
    mem = get_memory_info()
    print_info(f"System: {os.cpu_count()} CPUs, {mem['total_mb']/1024:.1f}GB RAM")
    
    # Generate files
    if not args.test_only:
        print_header(f"Generating Test Files (Level: {args.level})")
        generate_test_files(output_dir, args.level, verbose)
    
    if args.generate_only:
        print_ok("Test files generated. Use --test-only to run tests.")
        return
    
    # Run tests
    if args.compare:
        # Run WITHOUT fixes first
        print_header("Test 1: WITHOUT Fixes (Baseline)")
        baseline = run_stress_test(
            output_dir, args.concurrent, args.rounds, args.timeout,
            use_fixes=False, verbose=verbose
        )
        
        print("\nWaiting 5 seconds...\n")
        time.sleep(5)
        
        # Run WITH fixes
        print_header("Test 2: WITH Fixes")
        with_fixes = run_stress_test(
            output_dir, args.concurrent, args.rounds, args.timeout,
            use_fixes=True, verbose=verbose
        )
        
        # Compare
        print_header("Comparison Results")
        print(f"{'Metric':<25} {'Baseline':>12} {'With Fixes':>12} {'Change':>10}")
        print("-" * 60)
        
        for key, label in [
            ('success', 'Successful'),
            ('failed', 'Failed'),
            ('timeouts', 'Timeouts'),
            ('success_rate', 'Success Rate (%)'),
            ('avg_duration_sec', 'Avg Duration (s)'),
            ('peak_memory_gb', 'Peak Memory (GB)')
        ]:
            b = baseline[key]
            f = with_fixes[key]
            diff = f - b
            sign = '+' if diff > 0 else ''
            print(f"{label:<25} {b:>12} {f:>12} {sign}{diff:>9.1f}")
        
        print("-" * 60)
        
        if with_fixes['success_rate'] > baseline['success_rate']:
            print_ok(f"Success rate improved by {with_fixes['success_rate'] - baseline['success_rate']:.1f}%")
        if with_fixes['peak_memory_gb'] < baseline['peak_memory_gb']:
            reduction = (baseline['peak_memory_gb'] - with_fixes['peak_memory_gb']) / baseline['peak_memory_gb'] * 100
            print_ok(f"Peak memory reduced by {reduction:.1f}%")
    
    else:
        # Single test run
        print_header("Running Stress Test")
        results = run_stress_test(
            output_dir, args.concurrent, args.rounds, args.timeout,
            use_fixes=True, verbose=verbose
        )
        
        print_header("Results")
        print(f"Total conversions: {results['total']}")
        print(f"  Successful: {results['success']}")
        print(f"  Failed: {results['failed']}")
        print(f"  Timeouts: {results['timeouts']}")
        print(f"Success rate: {results['success_rate']}%")
        print(f"Peak memory: {results['peak_memory_gb']} GB")
        print(f"Total time: {results['total_time_sec']}s")


if __name__ == '__main__':
    main()
