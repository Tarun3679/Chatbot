#!/usr/bin/env python3
"""
LibreOffice OOM Stress Tester (2.5GB Limit Simulation) - FIXED VERSION
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
# SYSTEM UTILITIES (FIXED FOR DOCKER)
# =============================================================================

def get_memory_info() -> Dict[str, float]:
    """Reads memory info - tries cgroup limits first (Docker), falls back to /proc/meminfo"""
    try:
        # Try Docker cgroup v2 first (modern Docker)
        cgroup_limit = Path('/sys/fs/cgroup/memory.max')
        cgroup_usage = Path('/sys/fs/cgroup/memory.current')
        
        if cgroup_limit.exists() and cgroup_usage.exists():
            limit = int(cgroup_limit.read_text().strip())
            usage = int(cgroup_usage.read_text().strip())
            if limit != 9223372036854771712:  # Not "max" (unlimited)
                limit_mb = limit / (1024 * 1024)
                usage_mb = usage / (1024 * 1024)
                return {
                    'total_mb': limit_mb,
                    'used_mb': usage_mb,
                    'percent': round(usage_mb / limit_mb * 100, 1),
                    'source': 'cgroup_v2'
                }
        
        # Try Docker cgroup v1 (older Docker)
        cgroup_limit_v1 = Path('/sys/fs/cgroup/memory/memory.limit_in_bytes')
        cgroup_usage_v1 = Path('/sys/fs/cgroup/memory/memory.usage_in_bytes')
        
        if cgroup_limit_v1.exists() and cgroup_usage_v1.exists():
            limit = int(cgroup_limit_v1.read_text().strip())
            usage = int(cgroup_usage_v1.read_text().strip())
            if limit < 9223372036854771712:  # Not unlimited
                limit_mb = limit / (1024 * 1024)
                usage_mb = usage / (1024 * 1024)
                return {
                    'total_mb': limit_mb,
                    'used_mb': usage_mb,
                    'percent': round(usage_mb / limit_mb * 100, 1),
                    'source': 'cgroup_v1'
                }
        
        # Fallback to /proc/meminfo (host memory - WARNING!)
        with open('/proc/meminfo', 'r') as f:
            info = {line.split()[0].rstrip(':'): int(line.split()[1]) for line in f}
        total = info.get('MemTotal', 0) / 1024
        avail = info.get('MemAvailable', info.get('MemFree', 0)) / 1024
        used = total - avail
        return {
            'total_mb': total,
            'used_mb': used,
            'percent': round(used / total * 100, 1),
            'source': 'proc_meminfo_WARNING_HOST_MEMORY'
        }
    except Exception as e:
        return {'total_mb': 0, 'used_mb': 0, 'percent': 0, 'source': f'error: {e}'}

def find_libreoffice() -> Optional[str]:
    paths = ['/usr/bin/soffice', '/usr/bin/libreoffice', '/opt/libreoffice/program/soffice', shutil.which('soffice')]
    for p in paths:
        if p and os.path.isfile(p) and os.access(p, os.X_OK): return p
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
    error: str = ""
    stdout: str = ""
    stderr: str = ""

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
            return ConversionResult(
                input_path.name, 'OOM_KILLED', time.time()-start, 137,
                error="Process hit memory limit",
                stdout=result.stdout,
                stderr=result.stderr
            )
        
        # Check for dmesg OOM kill evidence
        try:
            dmesg_check = subprocess.run(['dmesg', '-T'], capture_output=True, text=True, timeout=5)
            if 'oom-kill' in dmesg_check.stdout.lower() or 'out of memory' in dmesg_check.stdout.lower():
                recent_oom = [line for line in dmesg_check.stdout.split('\n')[-50:] 
                             if 'oom' in line.lower() or 'killed process' in line.lower()]
                if recent_oom and 'soffice' in '\n'.join(recent_oom):
                    return ConversionResult(
                        input_path.name, 'OOM_KILLED_DMESG', time.time()-start, result.returncode,
                        error=f"OOM detected in dmesg: {recent_oom[-1][:100]}",
                        stdout=result.stdout,
                        stderr=result.stderr
                    )
        except:
            pass
            
        if (work_dir / (tmp_input.stem + '.pdf')).exists():
            return ConversionResult(
                input_path.name, 'SUCCESS', round(time.time()-start, 2), 0,
                stdout=result.stdout,
                stderr=result.stderr
            )
        
        return ConversionResult(
            input_path.name, 'FAILED', round(time.time()-start, 2), result.returncode,
            error=result.stderr[:200] if result.stderr else "No PDF output",
            stdout=result.stdout[:200],
            stderr=result.stderr[:200]
        )
    
    except subprocess.TimeoutExpired:
        return ConversionResult(
            input_path.name, 'TIMEOUT', time.time()-start, -1,
            error="Conversion took >180s"
        )
    except Exception as e:
        return ConversionResult(
            input_path.name, 'ERROR', time.time()-start, -1,
            error=str(e)[:200]
        )
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

# =============================================================================
# TEST RUNNER
# =============================================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--concurrent', type=int, default=6, help="Increase this to trigger OOM")
    parser.add_argument('--no-fixes', action='store_true', help="Disable LibreOffice fixes")
    parser.add_argument('--verbose', '-v', action='store_true', help="Show detailed error output")
    args = parser.parse_args()
    
    use_fixes = not args.no_fixes
    
    soffice = find_libreoffice()
    if not soffice:
        print("ERROR: LibreOffice not found!")
        sys.exit(1)
    
    test_files_dir = Path('./test_files')
    if not test_files_dir.exists():
        print(f"ERROR: {test_files_dir} directory not found!")
        print("Creating sample directory structure...")
        test_files_dir.mkdir(exist_ok=True)
        sys.exit(1)
    
    test_files = list(test_files_dir.glob('*.*'))
    if not test_files:
        print(f"ERROR: No files found in {test_files_dir}")
        sys.exit(1)
    
    # Show initial memory state
    initial_mem = get_memory_info()
    print(f"\n--- Starting OOM Stress Test (Target Limit: 2.5GB) ---")
    print(f"LibreOffice: {soffice}")
    print(f"Concurrent workers: {args.concurrent}")
    print(f"Using fixes: {use_fixes}")
    print(f"Test files: {len(test_files)}")
    print(f"Memory source: {initial_mem.get('source', 'unknown')}")
    print(f"Memory limit: {initial_mem['total_mb']:.0f}MB ({initial_mem['total_mb']/1024:.2f}GB)")
    print(f"Memory used (start): {initial_mem['used_mb']:.0f}MB ({initial_mem['used_mb']/1024:.2f}GB)")
    print()
    
    results = []
    with ThreadPoolExecutor(max_workers=args.concurrent) as exe:
        futures = {exe.submit(convert_to_pdf, f, soffice, use_fixes): f for f in test_files}
        for future in as_completed(futures):
            res = future.result()
            results.append(res)
            mem = get_memory_info()
            
            status_color = "[✓]" if res.status == 'SUCCESS' else "[✗]"
            print(f"{status_color} {res.filename:<25} | Status: {res.status:<15} | "
                  f"RAM Used: {mem['used_mb']/1024:.2f}GB | Exit: {res.exit_code}")
            
            if args.verbose and res.status != 'SUCCESS':
                print(f"    Duration: {res.duration:.2f}s")
                if res.error:
                    print(f"    Error: {res.error}")
                if res.stderr:
                    print(f"    Stderr: {res.stderr}")
                if res.stdout:
                    print(f"    Stdout: {res.stdout}")
                print()
    
    # Summary
    print("\n--- Summary ---")
    success = sum(1 for r in results if r.status == 'SUCCESS')
    failed = sum(1 for r in results if r.status == 'FAILED')
    oom = sum(1 for r in results if 'OOM' in r.status)
    print(f"Success: {success}, Failed: {failed}, OOM Killed: {oom}, Total: {len(results)}")
    
    final_mem = get_memory_info()
    print(f"Final memory: {final_mem['used_mb']/1024:.2f}GB ({final_mem['percent']}%)")

if __name__ == '__main__':
    main()
