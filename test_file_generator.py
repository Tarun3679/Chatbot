#!/usr/bin/env python3
"""
LibreOffice OOM Stress Tester - WORKING VERSION
Fixed X11 issues and proper environment detection
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
    """Reads memory info from best available source"""
    # Try all cgroup variants
    cgroup_checks = [
        ('/sys/fs/cgroup/memory.max', '/sys/fs/cgroup/memory.current', 'cgroup_v2'),
        ('/sys/fs/cgroup/memory/memory.limit_in_bytes', '/sys/fs/cgroup/memory/memory.usage_in_bytes', 'cgroup_v1'),
    ]
    
    for limit_path, usage_path, name in cgroup_checks:
        if Path(limit_path).exists() and Path(usage_path).exists():
            try:
                limit = Path(limit_path).read_text().strip()
                usage = Path(usage_path).read_text().strip()
                
                # Handle "max" value in cgroup v2
                if limit == 'max':
                    continue
                    
                limit_bytes = int(limit)
                usage_bytes = int(usage)
                
                # Sanity check - if limit is absurdly high, it's probably unlimited
                if limit_bytes > 100 * (1024**3):  # > 100GB
                    continue
                
                limit_mb = limit_bytes / (1024 * 1024)
                usage_mb = usage_bytes / (1024 * 1024)
                return {
                    'total_mb': limit_mb,
                    'used_mb': usage_mb,
                    'percent': round(usage_mb / limit_mb * 100, 1),
                    'source': name
                }
            except:
                continue
    
    # Fallback to /proc/meminfo (host memory)
    try:
        with open('/proc/meminfo', 'r') as f:
            info = {line.split()[0].rstrip(':'): int(line.split()[1]) for line in f}
        total = info.get('MemTotal', 0) / 1024
        avail = info.get('MemAvailable', info.get('MemFree', 0)) / 1024
        used = total - avail
        return {
            'total_mb': total,
            'used_mb': used,
            'percent': round(used / total * 100, 1),
            'source': 'proc_meminfo_HOST'
        }
    except:
        return {'total_mb': 0, 'used_mb': 0, 'percent': 0, 'source': 'error'}

def find_libreoffice() -> Optional[str]:
    paths = ['/usr/bin/soffice', '/usr/bin/libreoffice', '/opt/libreoffice/program/soffice', shutil.which('soffice')]
    for p in paths:
        if p and os.path.isfile(p) and os.access(p, os.X_OK): 
            return p
    return None

def detect_best_display_config() -> Dict[str, str]:
    """
    Auto-detect the best DISPLAY configuration for this environment.
    Different LibreOffice setups need different configurations.
    """
    # Test configurations to try (in order of preference)
    configs = [
        {'name': 'default', 'env': {}},  # Use existing environment
        {'name': 'xvfb_display', 'env': {'DISPLAY': ':99'}},  # Virtual X server
        {'name': 'empty_display', 'env': {'DISPLAY': ''}},  # Force no display
        {'name': 'gen_plugin', 'env': {'DISPLAY': '', 'SAL_USE_VCLPLUGIN': 'gen'}},  # Generic VCL
        {'name': 'svp_plugin', 'env': {'SAL_USE_VCLPLUGIN': 'svp'}},  # Server pages (headless)
    ]
    
    # Quick test with a minimal command
    soffice = find_libreoffice()
    if not soffice:
        return configs[0]
    
    for config in configs:
        try:
            env = os.environ.copy()
            env.update(config['env'])
            
            # Test if soffice can start with this config
            result = subprocess.run(
                [soffice, '--headless', '--version'],
                capture_output=True,
                text=True,
                timeout=5,
                env=env
            )
            
            # If no X11 error, this config works
            if 'X11 error' not in result.stderr and result.returncode == 0:
                print(f"   ✓ Best config: {config['name']}")
                return config
        except:
            continue
    
    # Default fallback
    print(f"   ⚠ Using default config (auto-detection failed)")
    return configs[0]

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

def convert_to_pdf(input_path: Path, soffice: str, env_config: Dict[str, str], use_profile_isolation: bool = True) -> ConversionResult:
    start = time.time()
    work_dir = Path(tempfile.gettempdir()) / f'lo_test_{uuid.uuid4().hex[:8]}'
    work_dir.mkdir(exist_ok=True)
    
    try:
        tmp_input = work_dir / input_path.name
        shutil.copy(input_path, tmp_input)
        
        # Build command
        cmd = [
            soffice, '--headless', '--nologo', '--nodefault',
            '--convert-to', 'pdf', '--outdir', str(work_dir)
        ]
        
        # Profile isolation (prevents "already running" errors)
        if use_profile_isolation:
            profile_dir = (work_dir / 'profile').absolute()
            profile_dir.mkdir(exist_ok=True)
            cmd.append(f'-env:UserInstallation=file://{profile_dir}')
        
        cmd.append(str(tmp_input))
        
        # Apply environment configuration
        env = os.environ.copy()
        env.update(env_config)
        
        result = subprocess.run(
            cmd, 
            capture_output=True, 
            text=True, 
            timeout=180, 
            env=env, 
            cwd=str(work_dir)
        )
        
        # Check for OOM kill
        if result.returncode == 137:
            return ConversionResult(
                input_path.name, 'OOM_KILLED', time.time()-start, 137,
                error="Process killed by OOM (exit 137)"
            )
        
        # Check for successful PDF creation
        expected_pdf = work_dir / (tmp_input.stem + '.pdf')
        if expected_pdf.exists():
            return ConversionResult(
                input_path.name, 'SUCCESS', round(time.time()-start, 2), 0
            )
        
        # Failed - return error details
        error_msg = result.stderr[:200] if result.stderr else "No PDF output"
        return ConversionResult(
            input_path.name, 'FAILED', round(time.time()-start, 2), 
            result.returncode, error_msg, result.stdout[:200], result.stderr[:200]
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
    parser = argparse.ArgumentParser(description='LibreOffice OOM Stress Tester')
    parser.add_argument('--concurrent', type=int, default=1, help="Number of concurrent workers")
    parser.add_argument('--no-profile-isolation', action='store_true', help="Disable profile isolation")
    parser.add_argument('--display-config', choices=['default', 'xvfb', 'empty', 'gen', 'svp', 'auto'], 
                       default='auto', help="X11 display configuration")
    parser.add_argument('--verbose', '-v', action='store_true', help="Show detailed error output")
    args = parser.parse_args()
    
    # Find LibreOffice
    soffice = find_libreoffice()
    if not soffice:
        print("ERROR: LibreOffice not found!")
        sys.exit(1)
    
    # Find test files
    test_files_dir = Path('./test_files')
    if not test_files_dir.exists():
        print(f"ERROR: {test_files_dir} directory not found!")
        sys.exit(1)
    
    test_files = list(test_files_dir.glob('*.*'))
    if not test_files:
        print(f"ERROR: No files found in {test_files_dir}")
        sys.exit(1)
    
    # Determine environment configuration
    print("\n--- Environment Detection ---")
    if args.display_config == 'auto':
        env_config_obj = detect_best_display_config()
        env_config = env_config_obj['env']
        config_name = env_config_obj['name']
    else:
        config_map = {
            'default': {},
            'xvfb': {'DISPLAY': ':99'},
            'empty': {'DISPLAY': ''},
            'gen': {'DISPLAY': '', 'SAL_USE_VCLPLUGIN': 'gen'},
            'svp': {'SAL_USE_VCLPLUGIN': 'svp'},
        }
        env_config = config_map[args.display_config]
        config_name = args.display_config
    
    # Show configuration
    mem = get_memory_info()
    print(f"\n--- Starting OOM Stress Test ---")
    print(f"LibreOffice: {soffice}")
    print(f"Concurrent workers: {args.concurrent}")
    print(f"Profile isolation: {not args.no_profile_isolation}")
    print(f"Display config: {config_name}")
    print(f"Test files: {len(test_files)}")
    print(f"Memory source: {mem.get('source', 'unknown')}")
    print(f"Memory limit: {mem['total_mb']:.0f}MB ({mem['total_mb']/1024:.2f}GB)")
    print(f"Memory used: {mem['used_mb']:.0f}MB ({mem['used_mb']/1024:.2f}GB)\n")
    
    # Run conversions
    results = []
    use_profile = not args.no_profile_isolation
    
    with ThreadPoolExecutor(max_workers=args.concurrent) as exe:
        futures = {exe.submit(convert_to_pdf, f, soffice, env_config, use_profile): f for f in test_files}
        for future in as_completed(futures):
            res = future.result()
            results.append(res)
            mem = get_memory_info()
            
            status_color = "[✓]" if res.status == 'SUCCESS' else "[✗]"
            print(f"{status_color} {res.filename:<25} | {res.status:<15} | "
                  f"RAM: {mem['used_mb']/1024:.2f}GB | Exit: {res.exit_code}")
            
            if args.verbose and res.status != 'SUCCESS':
                print(f"    Duration: {res.duration:.2f}s")
                if res.error:
                    print(f"    Error: {res.error}")
                if res.stderr:
                    print(f"    Stderr: {res.stderr[:150]}")
                print()
    
    # Summary
    print("\n--- Summary ---")
    success = sum(1 for r in results if r.status == 'SUCCESS')
    failed = sum(1 for r in results if r.status == 'FAILED')
    oom = sum(1 for r in results if 'OOM' in r.status)
    print(f"Success: {success}, Failed: {failed}, OOM: {oom}, Total: {len(results)}")
    
    final_mem = get_memory_info()
    print(f"Final memory: {final_mem['used_mb']/1024:.2f}GB ({final_mem['percent']}%)")

if __name__ == '__main__':
    main()
