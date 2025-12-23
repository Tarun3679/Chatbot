#!/usr/bin/env python3
"""
LibreOffice OOM Reproducer for Kubernetes Pods
Target: Exceed 2500 MiB pod limit to trigger OOM kill

Pod Configuration:
- Memory Request: 1500 MiB (guaranteed)
- Memory Limit:   2500 MiB (hard limit → OOM if exceeded)
- Host Memory:    30-40 GB (irrelevant)

This script progressively increases concurrent conversions until the pod
exceeds 2500 MiB and gets OOM killed.
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
# CONTAINER MEMORY TRACKING (POD LIMITS)
# =============================================================================

def get_pod_memory_limit() -> Optional[int]:
    """
    Detect the pod's cgroup memory limit in bytes.
    This should be 2500 MiB = 2621440000 bytes
    """
    cgroup_paths = [
        '/sys/fs/cgroup/memory.max',                          # cgroup v2
        '/sys/fs/cgroup/memory/memory.limit_in_bytes',        # cgroup v1
    ]
    
    for path in cgroup_paths:
        if Path(path).exists():
            try:
                limit = Path(path).read_text().strip()
                if limit == 'max':
                    continue
                limit_bytes = int(limit)
                # Only accept limits < 100GB (pod limits are much smaller)
                if limit_bytes < 100 * 1024 * 1024 * 1024:
                    return limit_bytes
            except:
                continue
    return None

def get_pod_memory_usage() -> Optional[int]:
    """
    Get current pod memory usage in bytes from cgroup.
    """
    cgroup_paths = [
        '/sys/fs/cgroup/memory.current',                      # cgroup v2
        '/sys/fs/cgroup/memory/memory.usage_in_bytes',        # cgroup v1
    ]
    
    for path in cgroup_paths:
        if Path(path).exists():
            try:
                usage = Path(path).read_text().strip()
                return int(usage)
            except:
                continue
    return None

def get_memory_pressure() -> Dict[str, any]:
    """
    Get pod memory statistics relative to the 2500 MiB limit.
    """
    limit_bytes = get_pod_memory_limit()
    usage_bytes = get_pod_memory_usage()
    
    if limit_bytes and usage_bytes:
        limit_mib = limit_bytes / (1024 * 1024)
        usage_mib = usage_bytes / (1024 * 1024)
        available_mib = (limit_bytes - usage_bytes) / (1024 * 1024)
        percent = (usage_bytes / limit_bytes) * 100
        
        # Calculate pressure zones
        if percent < 60:  # < 1500 MiB
            zone = "SAFE"
            color = "🟢"
        elif percent < 80:  # 1500-2000 MiB
            zone = "WARNING"
            color = "🟡"
        elif percent < 95:  # 2000-2375 MiB
            zone = "DANGER"
            color = "🟠"
        else:  # > 2375 MiB
            zone = "CRITICAL"
            color = "🔴"
        
        return {
            'limit_mib': limit_mib,
            'usage_mib': usage_mib,
            'available_mib': available_mib,
            'percent': percent,
            'zone': zone,
            'color': color,
            'has_limit': True
        }
    else:
        # Fallback - no cgroup limit detected
        return {
            'limit_mib': 0,
            'usage_mib': 0,
            'available_mib': 0,
            'percent': 0,
            'zone': 'UNKNOWN',
            'color': '⚪',
            'has_limit': False
        }

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
    """
    Convert document to PDF with memory tracking.
    Uses default environment (no special fixes) since those broke your setup.
    """
    start = time.time()
    mem_before = get_memory_pressure()
    
    work_dir = Path(tempfile.gettempdir()) / f'lo_oom_{uuid.uuid4().hex[:8]}'
    work_dir.mkdir(exist_ok=True)
    
    try:
        # Copy file to work directory
        tmp_input = work_dir / input_path.name
        shutil.copy(input_path, tmp_input)
        
        # Build command with profile isolation (prevents "already running" errors)
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
        
        # Run conversion
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=180,
            cwd=str(work_dir)
        )
        
        mem_after = get_memory_pressure()
        duration = time.time() - start
        
        # Check for OOM kill (exit code 137)
        if result.returncode == 137:
            return ConversionResult(
                input_path.name,
                'OOM_KILLED',
                duration,
                137,
                mem_before['usage_mib'],
                mem_after['usage_mib'],
                error="⚠️ KILLED BY OOM - Pod exceeded 2500 MiB limit"
            )
        
        # Check for success
        expected_pdf = work_dir / (tmp_input.stem + '.pdf')
        if expected_pdf.exists():
            return ConversionResult(
                input_path.name,
                'SUCCESS',
                round(duration, 2),
                0,
                mem_before['usage_mib'],
                mem_after['usage_mib']
            )
        
        # Failed
        return ConversionResult(
            input_path.name,
            'FAILED',
            round(duration, 2),
            result.returncode,
            mem_before['usage_mib'],
            mem_after['usage_mib'],
            error=result.stderr[:150] if result.stderr else "No PDF output"
        )
    
    except subprocess.TimeoutExpired:
        mem_after = get_memory_pressure()
        return ConversionResult(
            input_path.name,
            'TIMEOUT',
            time.time() - start,
            -1,
            mem_before['usage_mib'],
            mem_after['usage_mib'],
            error="Timeout (>180s)"
        )
    except Exception as e:
        mem_after = get_memory_pressure()
        return ConversionResult(
            input_path.name,
            'ERROR',
            time.time() - start,
            -1,
            mem_before['usage_mib'],
            mem_after['usage_mib'],
            error=str(e)[:150]
        )
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

# =============================================================================
# TEST RUNNER
# =============================================================================

def print_memory_bar(percent: float, width: int = 50):
    """Print visual memory usage bar"""
    filled = int((percent / 100) * width)
    bar = '█' * filled + '░' * (width - filled)
    return f"[{bar}] {percent:.1f}%"

def run_oom_test(concurrent: int, test_files: List[Path], soffice: str, verbose: bool = False):
    """
    Run OOM stress test with specified concurrency level.
    Returns True if OOM was triggered.
    """
    mem = get_memory_pressure()
    
    print(f"\n{'='*80}")
    print(f"🔄 RUNNING TEST: {concurrent} concurrent conversions")
    print(f"{'='*80}")
    print(f"{mem['color']} Memory: {mem['usage_mib']:.0f}/{mem['limit_mib']:.0f} MiB "
          f"({mem['percent']:.1f}%) - {mem['zone']}")
    print(f"   Available: {mem['available_mib']:.0f} MiB")
    print(f"   {print_memory_bar(mem['percent'])}")
    print()
    
    results = []
    oom_triggered = False
    
    with ThreadPoolExecutor(max_workers=concurrent) as executor:
        # Submit all conversions
        futures = {executor.submit(convert_to_pdf, f, soffice): f for f in test_files}
        
        # Process results as they complete
        for future in as_completed(futures):
            res = future.result()
            results.append(res)
            
            # Update memory stats
            mem = get_memory_pressure()
            
            # Format output
            status_icon = "✓" if res.status == 'SUCCESS' else "✗"
            if res.status == 'OOM_KILLED':
                status_icon = "💀"
                oom_triggered = True
            
            print(f"{status_icon} {res.filename:<25} {res.status:<12} "
                  f"{res.duration:>6.2f}s | "
                  f"{mem['color']} {mem['usage_mib']:>7.0f} MiB ({mem['percent']:>5.1f}%) | "
                  f"Exit: {res.exit_code}")
            
            if verbose and res.error:
                print(f"   └─ {res.error}")
    
    # Summary
    success = sum(1 for r in results if r.status == 'SUCCESS')
    failed = sum(1 for r in results if r.status == 'FAILED')
    oom = sum(1 for r in results if 'OOM' in r.status)
    
    print(f"\n📊 Results: {success} success, {failed} failed, {oom} OOM killed")
    
    final_mem = get_memory_pressure()
    print(f"{final_mem['color']} Final Memory: {final_mem['usage_mib']:.0f}/{final_mem['limit_mib']:.0f} MiB "
          f"({final_mem['percent']:.1f}%)")
    
    return oom_triggered

# =============================================================================
# MAIN
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='LibreOffice OOM Reproducer for K8s Pods (2500 MiB limit)'
    )
    parser.add_argument('--concurrent', type=int, help="Fixed concurrency level")
    parser.add_argument('--auto-escalate', action='store_true',
                       help="Auto-escalate concurrency until OOM is triggered")
    parser.add_argument('--start', type=int, default=2,
                       help="Starting concurrency for auto-escalate (default: 2)")
    parser.add_argument('--step', type=int, default=2,
                       help="Concurrency increment for auto-escalate (default: 2)")
    parser.add_argument('--max', type=int, default=30,
                       help="Maximum concurrency for auto-escalate (default: 30)")
    parser.add_argument('--verbose', '-v', action='store_true',
                       help="Show detailed error output")
    args = parser.parse_args()
    
    # Find LibreOffice
    soffice = find_libreoffice()
    if not soffice:
        print("❌ ERROR: LibreOffice not found!")
        sys.exit(1)
    
    # Find test files
    test_files_dir = Path('./test_files')
    if not test_files_dir.exists():
        print(f"❌ ERROR: {test_files_dir} directory not found!")
        sys.exit(1)
    
    test_files = list(test_files_dir.glob('*.*'))
    if not test_files:
        print(f"❌ ERROR: No files found in {test_files_dir}")
        sys.exit(1)
    
    # Check pod memory limit
    mem = get_memory_pressure()
    
    print("\n" + "="*80)
    print("🎯 LIBREOFFICE OOM STRESS TESTER")
    print("="*80)
    print(f"📦 Pod Memory Configuration:")
    
    if mem['has_limit']:
        print(f"   ✓ Memory Limit Detected: {mem['limit_mib']:.0f} MiB")
        print(f"   Current Usage: {mem['usage_mib']:.0f} MiB ({mem['percent']:.1f}%)")
        print(f"   Available: {mem['available_mib']:.0f} MiB")
        print(f"   Target: Exceed {mem['limit_mib']:.0f} MiB to trigger OOM")
    else:
        print(f"   ⚠️  WARNING: No cgroup memory limit detected!")
        print(f"   This might not be running in a memory-limited container.")
        print(f"   Continuing anyway...")
    
    print(f"\n📂 Test Configuration:")
    print(f"   LibreOffice: {soffice}")
    print(f"   Test Files: {len(test_files)} files")
    for f in test_files:
        size_mb = f.stat().st_size / (1024 * 1024)
        print(f"      • {f.name} ({size_mb:.2f} MB)")
    
    # Run test(s)
    if args.concurrent:
        # Fixed concurrency mode
        print(f"\n🎯 Mode: Fixed concurrency ({args.concurrent} workers)")
        run_oom_test(args.concurrent, test_files, soffice, args.verbose)
    
    elif args.auto_escalate:
        # Auto-escalation mode
        print(f"\n🎯 Mode: Auto-escalate ({args.start} → {args.max}, step {args.step})")
        print(f"   Strategy: Increase concurrency until OOM is triggered")
        
        oom_triggered = False
        current = args.start
        
        while current <= args.max and not oom_triggered:
            oom_triggered = run_oom_test(current, test_files, soffice, args.verbose)
            
            if oom_triggered:
                print(f"\n🎉 SUCCESS! OOM triggered at concurrency level: {current}")
                print(f"   This means ~{current} concurrent LibreOffice processes")
                print(f"   exceeded the 2500 MiB pod limit.")
                break
            
            if current < args.max:
                current += args.step
                print(f"\n⬆️  Increasing concurrency to {current}...")
                time.sleep(2)  # Brief pause between escalations
        
        if not oom_triggered:
            print(f"\n⚠️  OOM not triggered even at {args.max} concurrent conversions.")
            print(f"   Either the limit is higher than expected, or files are too small.")
    
    else:
        # Default: single run with 6 concurrent
        print(f"\n🎯 Mode: Default (6 concurrent conversions)")
        print(f"   Tip: Use --auto-escalate to find OOM threshold automatically")
        run_oom_test(6, test_files, soffice, args.verbose)
    
    print("\n" + "="*80)
    print("✅ TEST COMPLETE")
    print("="*80)

if __name__ == '__main__':
    main()
