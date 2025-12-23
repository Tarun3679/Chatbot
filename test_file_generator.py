#!/usr/bin/env python3
"""
LibreOffice OOM Stress Tester (2.5GB Limit Simulation)
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, Optional
from dataclasses import dataclass

# =============================================================================
# SYSTEM UTILITIES
# =============================================================================

def get_memory_info() -> Dict[str, float]:
    """Reads system memory from /proc/meminfo"""
    try:
        with open('/proc/meminfo', 'r') as f:
            info = {line.split()[0].rstrip(':'): int(line.split()[1]) for line in f}
        total = info.get('MemTotal', 0) / 1024
        avail = info.get('MemAvailable', info.get('MemFree', 0)) / 1024
        used = total - avail
        return {'total_mb': total, 'used_mb': used, 'percent': round(used / total * 100, 1)}
    except:
        return {'total_mb': 0, 'used_mb': 0, 'percent': 0}

def find_libreoffice() -> Optional[str]:
    paths = ['/usr/bin/soffice', '/usr/bin/libreoffice', '/opt/libreoffice/program/soffice', shutil.which('soffice')]
    for p in paths:
        if p and os.path.isfile(p) and os.access(p, os.X_OK): return p
    return None

# =============================================================================
# CONVERSION ENGINE (FIXED)
# =============================================================================

@dataclass
class ConversionResult:
    filename: str
    status: str
    duration: float
    exit_code: int
    error: str = ""

def convert_to_pdf(input_path: Path, soffice: str, use_fixes: bool = True) -> ConversionResult:
    start = time.time()
    work_dir = Path(tempfile.gettempdir()) / f'lo_test_{uuid.uuid4().hex[:8]}'
    work_dir.mkdir(exist_ok=True)
    
    try:
        tmp_input = work_dir / input_path.name
        shutil.copy(input_path, tmp_input)
        
        # Build command with headless best practices
        cmd = [
            soffice, '--headless', '--nologo', '--nodefault',
            '--convert-to', 'pdf', '--outdir', str(work_dir)
        ]
        
        # Profile Isolation is key for high concurrency
        if use_fixes:
            profile_dir = (work_dir / 'profile').absolute()
            profile_dir.mkdir(exist_ok=True)
            cmd.append(f'-env:UserInstallation=file://{profile_dir}')
        
        cmd.append(str(tmp_input))
        
        # CRITICAL: Bypasses X11 "Can't open display"
        env = os.environ.copy()
        if use_fixes:
            env['DISPLAY'] = ''
            env['SAL_USE_VCLPLUGIN'] = 'gen'
        
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=180, env=env, cwd=str(work_dir))
        
        # Detect OOM Kill (Exit Code 137)
        if result.returncode == 137:
            return ConversionResult(input_path.name, 'OOM_KILLED', time.time()-start, 137, "Process hit memory limit")
            
        if (work_dir / (tmp_input.stem + '.pdf')).exists():
            return ConversionResult(input_path.name, 'SUCCESS', round(time.time()-start, 2), 0)
        
        return ConversionResult(input_path.name, 'FAILED', round(time.time()-start, 2), result.returncode, result.stderr[:100])
    
    except Exception as e:
        return ConversionResult(input_path.name, 'ERROR', time.time()-start, -1, str(e)[:100])
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

# =============================================================================
# TEST RUNNER
# =============================================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--concurrent', type=int, default=6, help="Increase this to trigger OOM")
    args = parser.parse_args()
    
    soffice = find_libreoffice()
    test_files = list(Path('./test_files').glob('*.*'))
    
    print(f"\n--- Starting OOM Stress Test (Target Limit: 2.5GB) ---")
    print(f"Concurrent workers: {args.concurrent}")
    
    with ThreadPoolExecutor(max_workers=args.concurrent) as exe:
        futures = {exe.submit(convert_to_pdf, f, soffice): f for f in test_files}
        for future in as_completed(futures):
            res = future.result()
            mem = get_memory_info()
            status_color = "[✓]" if res.status == 'SUCCESS' else "[✗]"
            print(f"{status_color} {res.filename:<20} | Status: {res.status:<10} | RAM Used: {mem['used_mb']/1024:.2f}GB")

if __name__ == '__main__':
    main()
