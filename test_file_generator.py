#!/usr/bin/env python3
"""
LibreOffice OOM Test with PID Memory Tracking and 2000 MiB Cap
Auto-escalates concurrency until memory reaches 2000 MiB or OOM occurs
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
from typing import List, Dict, Optional
from dataclasses import dataclass
import threading

# Check for psutil
try:
    import psutil
except ImportError:
    print("ERROR: psutil not installed")
    print("Install with: pip install psutil --break-system-packages")
    sys.exit(1)

# Global memory tracking
memory_cap_reached = threading.Event()
MEMORY_CAP_MIB = 2000

# =============================================================================
# MEMORY TRACKING
# =============================================================================

def get_all_libreoffice_pids() -> List[int]:
    """Find all LibreOffice/soffice process PIDs"""
    pids = []
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            name = proc.info['name'].lower()
            cmdline = ' '.join(proc.info['cmdline'] or []).lower()
            if 'soffice' in name or 'libreoffice' in name or 'soffice' in cmdline:
                pids.append(proc.info['pid'])
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return pids

def get_process_memory_mib(pid: int) -> float:
    """Get memory usage in MiB for a specific PID"""
    try:
        proc = psutil.Process(pid)
        mem_info = proc.memory_info()
        return mem_info.rss / (1024 * 1024)
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return 0.0

def get_total_libreoffice_memory() -> Dict[str, float]:
    """Get total memory used by all LibreOffice processes"""
    pids = get_all_libreoffice_pids()
    total_rss_mib = 0.0
    process_details = []
    
    for pid in pids:
        mem_mib = get_process_memory_mib(pid)
        if mem_mib > 0:
            total_rss_mib += mem_mib
            process_details.append({'pid': pid, 'memory_mib': mem_mib})
    
    return {
        'total_mib': total_rss_mib,
        'process_count': len(pids),
        'processes': process_details
    }

def memory_monitor_thread(interval: float = 0.5, verbose: bool = False):
    """Background thread to monitor memory and trigger cap alert"""
    global memory_cap_reached
    
    while not memory_cap_reached.is_set():
        mem_stats = get_total_libreoffice_memory()
        total_mib = mem_stats['total_mib']
        
        if verbose and mem_stats['process_count'] > 0:
            print(f"[MEMORY MONITOR] Total: {total_mib:.1f} MiB across {mem_stats['process_count']} processes")
        
        if total_mib >= MEMORY_CAP_MIB:
            memory_cap_reached.set()
            print("\n" + "="*80)
            print("MEMORY CAP REACHED: {:.1f} MiB >= {} MiB".format(total_mib, MEMORY_CAP_MIB))
            print("="*80)
            print("Process breakdown:")
            for proc in mem_stats['processes']:
                print("  PID {}: {:.1f} MiB".format(proc['pid'], proc['memory_mib']))
            print("="*80)
            print("Stopping test to prevent OOM kill")
            print("="*80)
            break
        
        time.sleep(interval)

# =============================================================================
# LIBREOFFICE UTILITIES
# =============================================================================

def find_libreoffice() -> Optional[str]:
    """Find LibreOffice executable"""
    paths = [
        '/usr/bin/soffice',
        '/usr/bin/libreoffice',
        '/opt/libreoffice/program/soffice',
        shutil.which('soffice')
    ]
    for p in paths:
        if p and os.path.isfile(p) and os.access(p, os.X_OK):
            return p
    return None

# =============================================================================
# CONVERSION ENGINE
# =============================================================================

@dataclass
class ConversionResult:
    filename: str
    status: str
    duration: float
    exit_code: int
    memory_before_mib: float
    memory_after_mib: float
    error: str = ""

def convert_to_pdf(input_path: Path, soffice: str) -> ConversionResult:
    """Convert document to PDF with memory tracking"""
    start = time.time()
    
    # Check if cap already reached
    if memory_cap_reached.is_set():
        return ConversionResult(
            input_path.name, 'SKIPPED', 0.0, -2, 0.0, 0.0,
            error="Memory cap reached before conversion started"
        )
    
    mem_before = get_total_libreoffice_memory()
    
    work_dir = Path(tempfile.gettempdir()) / f'lo_test_{uuid.uuid4().hex[:8]}'
    work_dir.mkdir(exist_ok=True)
    
    try:
        tmp_input = work_dir / input_path.name
        shutil.copy(input_path, tmp_input)
        
        # Profile isolation
        profile_dir = (work_dir / 'profile').absolute()
        profile_dir.mkdir(exist_ok=True)
        
        cmd = [
            soffice,
            '--headless',
            '--nologo',
            '--nodefault',
            '--convert-to', 'pdf',
            '--outdir', str(work_dir),
            f'-env:UserInstallation=file://{profile_dir}',
            str(tmp_input)
        ]
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=180,
            cwd=str(work_dir)
        )
        
        mem_after = get_total_libreoffice_memory()
        duration = time.time() - start
        
        # Check for OOM kill (exit 137)
        if result.returncode == 137:
            return ConversionResult(
                input_path.name,
                'OOM_KILLED',
                duration,
                137,
                mem_before['total_mib'],
                mem_after['total_mib'],
                error="Process killed by OOM killer"
            )
        
        # Check for success
        expected_pdf = work_dir / (tmp_input.stem + '.pdf')
        if expected_pdf.exists():
            return ConversionResult(
                input_path.name,
                'SUCCESS',
                round(duration, 2),
                0,
                mem_before['total_mib'],
                mem_after['total_mib']
            )
        
        # Failed
        return ConversionResult(
            input_path.name,
            'FAILED',
            round(duration, 2),
            result.returncode,
            mem_before['total_mib'],
            mem_after['total_mib'],
            error=result.stderr[:100] if result.stderr else "No PDF output"
        )
    
    except subprocess.TimeoutExpired:
        mem_after = get_total_libreoffice_memory()
        return ConversionResult(
            input_path.name,
            'TIMEOUT',
            time.time() - start,
            -1,
            mem_before['total_mib'],
            mem_after['total_mib'],
            error="Conversion timeout (>180s)"
        )
    except Exception as e:
        mem_after = get_total_libreoffice_memory()
        return ConversionResult(
            input_path.name,
            'ERROR',
            time.time() - start,
            -1,
            mem_before['total_mib'],
            mem_after['total_mib'],
            error=str(e)[:100]
        )
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

# =============================================================================
# TEST RUNNER
# =============================================================================

def run_concurrent_test(concurrent: int, test_files: List[Path], soffice: str, verbose: bool = False) -> Dict:
    """Run conversions at specified concurrency level"""
    
    print("\n" + "="*80)
    print("RUNNING TEST: {} concurrent conversions".format(concurrent))
    print("="*80)
    
    mem_start = get_total_libreoffice_memory()
    print("Memory at start: {:.1f} MiB ({} processes)".format(
        mem_start['total_mib'], mem_start['process_count']))
    print()
    
    results = []
    oom_triggered = False
    cap_hit = False
    
    with ThreadPoolExecutor(max_workers=concurrent) as executor:
        futures = {executor.submit(convert_to_pdf, f, soffice): f for f in test_files}
        
        for future in as_completed(futures):
            # Check if cap was reached
            if memory_cap_reached.is_set():
                cap_hit = True
                # Cancel remaining futures
                for f in futures:
                    f.cancel()
                break
            
            res = future.result()
            results.append(res)
            
            mem_current = get_total_libreoffice_memory()
            
            status_symbol = "[SUCCESS]" if res.status == 'SUCCESS' else "[FAILED]"
            if res.status == 'OOM_KILLED':
                status_symbol = "[OOM]"
                oom_triggered = True
            
            print("{} {} | {:.2f}s | Mem: {:.1f} -> {:.1f} MiB | Total: {:.1f} MiB | Exit: {}".format(
                status_symbol,
                res.filename[:20].ljust(20),
                res.duration,
                res.memory_before_mib,
                res.memory_after_mib,
                mem_current['total_mib'],
                res.exit_code
            ))
            
            if verbose and res.error:
                print("  Error: {}".format(res.error))
    
    # Summary
    success = sum(1 for r in results if r.status == 'SUCCESS')
    failed = sum(1 for r in results if r.status == 'FAILED')
    oom = sum(1 for r in results if 'OOM' in r.status)
    skipped = sum(1 for r in results if r.status == 'SKIPPED')
    
    mem_final = get_total_libreoffice_memory()
    
    print("\nResults: {} success, {} failed, {} OOM, {} skipped".format(
        success, failed, oom, skipped))
    print("Final memory: {:.1f} MiB ({} processes)".format(
        mem_final['total_mib'], mem_final['process_count']))
    
    if verbose and mem_final['process_count'] > 0:
        print("\nProcess details:")
        for proc in mem_final['processes']:
            print("  PID {}: {:.1f} MiB".format(proc['pid'], proc['memory_mib']))
    
    return {
        'oom_triggered': oom_triggered,
        'cap_reached': cap_hit,
        'success': success,
        'failed': failed,
        'oom': oom,
        'final_memory_mib': mem_final['total_mib']
    }

# =============================================================================
# MAIN
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='LibreOffice OOM Test with Memory Cap and PID Tracking'
    )
    parser.add_argument('--concurrent', type=int, help='Fixed concurrency level')
    parser.add_argument('--auto-escalate', action='store_true',
                       help='Auto-escalate concurrency until cap/OOM')
    parser.add_argument('--start', type=int, default=5,
                       help='Starting concurrency (default: 5)')
    parser.add_argument('--step', type=int, default=5,
                       help='Concurrency increment (default: 5)')
    parser.add_argument('--max', type=int, default=50,
                       help='Maximum concurrency (default: 50)')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Verbose output with process details')
    parser.add_argument('--monitor-interval', type=float, default=0.5,
                       help='Memory monitoring interval in seconds (default: 0.5)')
    args = parser.parse_args()
    
    # Find LibreOffice
    soffice = find_libreoffice()
    if not soffice:
        print("ERROR: LibreOffice not found")
        sys.exit(1)
    
    # Find test files
    test_files_dir = Path('./test_files')
    if not test_files_dir.exists():
        print("ERROR: {} directory not found".format(test_files_dir))
        sys.exit(1)
    
    test_files = list(test_files_dir.glob('*.*'))
    if not test_files:
        print("ERROR: No files found in {}".format(test_files_dir))
        sys.exit(1)
    
    # Display configuration
    print("\n" + "="*80)
    print("LIBREOFFICE OOM TEST WITH MEMORY CAP")
    print("="*80)
    print("LibreOffice: {}".format(soffice))
    print("Test files: {}".format(len(test_files)))
    for f in test_files:
        size_mb = f.stat().st_size / (1024 * 1024)
        print("  - {} ({:.2f} MB)".format(f.name, size_mb))
    print("\nMemory cap: {} MiB".format(MEMORY_CAP_MIB))
    print("Target: Reach {} MiB or trigger OOM (exit 137)".format(MEMORY_CAP_MIB))
    
    # Start memory monitor thread
    monitor = threading.Thread(
        target=memory_monitor_thread,
        args=(args.monitor_interval, args.verbose),
        daemon=True
    )
    monitor.start()
    
    # Run tests
    if args.concurrent:
        # Fixed concurrency mode
        print("\nMode: Fixed concurrency ({} workers)".format(args.concurrent))
        result = run_concurrent_test(args.concurrent, test_files, soffice, args.verbose)
        
    elif args.auto_escalate:
        # Auto-escalation mode
        print("\nMode: Auto-escalate ({} -> {}, step {})".format(
            args.start, args.max, args.step))
        print("Strategy: Increase concurrency until {} MiB cap or OOM".format(MEMORY_CAP_MIB))
        
        current = args.start
        
        while current <= args.max:
            if memory_cap_reached.is_set():
                print("\nMemory cap reached. Stopping escalation.")
                break
            
            result = run_concurrent_test(current, test_files, soffice, args.verbose)
            
            if result['oom_triggered']:
                print("\n" + "="*80)
                print("OOM TRIGGERED at concurrency level: {}".format(current))
                print("="*80)
                break
            
            if result['cap_reached']:
                print("\n" + "="*80)
                print("Memory cap ({} MiB) reached at concurrency: {}".format(
                    MEMORY_CAP_MIB, current))
                print("Final memory: {:.1f} MiB".format(result['final_memory_mib']))
                print("="*80)
                break
            
            if current < args.max:
                current += args.step
                print("\nIncreasing concurrency to {}...".format(current))
                time.sleep(2)
        
        if current > args.max and not result['oom_triggered'] and not result['cap_reached']:
            print("\n" + "="*80)
            print("Reached max concurrency ({}) without hitting cap or OOM".format(args.max))
            print("Final memory: {:.1f} MiB".format(result['final_memory_mib']))
            print("="*80)
    
    else:
        # Default mode
        print("\nMode: Default (10 concurrent conversions)")
        result = run_concurrent_test(10, test_files, soffice, args.verbose)
    
    # Wait for monitor thread to finish
    memory_cap_reached.set()
    monitor.join(timeout=1)
    
    print("\n" + "="*80)
    print("TEST COMPLETE")
    print("="*80)

if __name__ == '__main__':
    main()
