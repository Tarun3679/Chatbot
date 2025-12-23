#!/usr/bin/env python3
"""
LibreOffice OOM Stress Test Runner for Linux

This script runs stress tests on LibreOffice PDF conversions to reproduce
OOM (Out of Memory) errors and test fixes.

Features:
- Real-time memory monitoring
- Concurrent conversion testing
- dmesg OOM detection
- Detailed reporting

Usage:
    python stress_test_runner.py [--test-dir ./test_files] [--concurrent 4]

Requirements:
    - Linux operating system
    - LibreOffice installed (soffice command available)
    - Python packages: psutil
"""

import argparse
import json
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time
import uuid
from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional, Callable

# Check for psutil
try:
    import psutil
except ImportError:
    print("ERROR: psutil is required. Install with: pip install psutil")
    sys.exit(1)


@dataclass
class ConversionResult:
    """Result of a single conversion attempt."""
    filename: str
    input_size_mb: float
    output_size_kb: float
    status: str  # 'success', 'failed', 'timeout', 'oom'
    duration_seconds: float
    error_message: str = ""
    memory_peak_mb: float = 0
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())


@dataclass 
class StressTestReport:
    """Complete stress test report."""
    start_time: str
    end_time: str
    duration_seconds: float
    system_info: Dict
    test_config: Dict
    results: List[ConversionResult]
    summary: Dict
    oom_events: List[str]


def get_system_info() -> Dict:
    """Gather system information for the report."""
    mem = psutil.virtual_memory()
    
    # Get LibreOffice version
    lo_version = "unknown"
    try:
        result = subprocess.run(
            ["soffice", "--version"],
            capture_output=True, text=True, timeout=10
        )
        lo_version = result.stdout.strip() or result.stderr.strip()
    except Exception:
        pass
    
    return {
        "hostname": os.uname().nodename,
        "kernel": os.uname().release,
        "cpu_count": psutil.cpu_count(),
        "cpu_count_physical": psutil.cpu_count(logical=False),
        "total_memory_gb": round(mem.total / (1024**3), 2),
        "available_memory_gb": round(mem.available / (1024**3), 2),
        "swap_total_gb": round(psutil.swap_memory().total / (1024**3), 2),
        "libreoffice_version": lo_version,
        "python_version": sys.version.split()[0],
    }


def find_libreoffice_path() -> str:
    """Find the LibreOffice soffice executable."""
    possible_paths = [
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/usr/local/bin/soffice",
        "/opt/libreoffice/program/soffice",
        "/snap/bin/libreoffice",
        shutil.which("soffice"),
        shutil.which("libreoffice"),
    ]
    
    for path in possible_paths:
        if path and os.path.isfile(path) and os.access(path, os.X_OK):
            return path
    
    raise RuntimeError(
        "LibreOffice not found. Please install it:\n"
        "  Ubuntu/Debian: sudo apt install libreoffice\n"
        "  RHEL/CentOS:   sudo dnf install libreoffice\n"
        "  Arch:          sudo pacman -S libreoffice-fresh"
    )


def get_optimized_env() -> Dict[str, str]:
    """Get environment variables optimized for LibreOffice headless conversion."""
    env = os.environ.copy()
    
    # Disable GPU/hardware acceleration
    env["SAL_DISABLE_OPENCL"] = "1"
    env["SAL_DISABLEGL"] = "1"
    env["SAL_DISABLESKIA"] = "1"
    
    # Use generic VCL plugin
    env["SAL_USE_VCLPLUGIN"] = "gen"
    
    # Prevent crash dialogs
    env["SAL_NO_CRASHREPORT"] = "1"
    
    # Limit Java heap
    env["JAVA_TOOL_OPTIONS"] = "-Xmx256m"
    
    # Set HOME to temp to avoid profile conflicts
    # (will be overridden per-conversion with unique profile)
    
    return env


class MemoryMonitor:
    """Background memory monitoring thread."""
    
    def __init__(self, interval: float = 0.5):
        self.interval = interval
        self.peak_memory_mb = 0
        self.peak_percent = 0
        self.running = False
        self._thread: Optional[threading.Thread] = None
        self._lock = threading.Lock()
        self.readings: List[Dict] = []
    
    def start(self):
        """Start monitoring in background."""
        self.running = True
        self.peak_memory_mb = 0
        self.readings = []
        self._thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self._thread.start()
    
    def stop(self) -> float:
        """Stop monitoring and return peak memory usage."""
        self.running = False
        if self._thread:
            self._thread.join(timeout=2)
        return self.peak_memory_mb
    
    def _monitor_loop(self):
        """Internal monitoring loop."""
        while self.running:
            try:
                mem = psutil.virtual_memory()
                used_mb = mem.used / (1024**2)
                
                with self._lock:
                    if used_mb > self.peak_memory_mb:
                        self.peak_memory_mb = used_mb
                        self.peak_percent = mem.percent
                    
                    self.readings.append({
                        "timestamp": time.time(),
                        "used_mb": used_mb,
                        "percent": mem.percent,
                        "available_mb": mem.available / (1024**2)
                    })
                
                time.sleep(self.interval)
            except Exception:
                pass
    
    def get_current(self) -> Dict:
        """Get current memory stats."""
        mem = psutil.virtual_memory()
        return {
            "used_mb": mem.used / (1024**2),
            "percent": mem.percent,
            "available_mb": mem.available / (1024**2),
            "peak_mb": self.peak_memory_mb
        }


def check_dmesg_for_oom(since_timestamp: float) -> List[str]:
    """Check dmesg for OOM killer events since a given timestamp."""
    oom_events = []
    
    try:
        # Try to read dmesg (may require sudo)
        result = subprocess.run(
            ["dmesg", "--time-format=iso", "-l", "err,warn"],
            capture_output=True, text=True, timeout=5
        )
        
        for line in result.stdout.splitlines():
            line_lower = line.lower()
            if any(term in line_lower for term in ["out of memory", "oom-killer", "killed process", "memory cgroup"]):
                oom_events.append(line.strip())
    except subprocess.TimeoutExpired:
        pass
    except PermissionError:
        # Try with journalctl as fallback
        try:
            result = subprocess.run(
                ["journalctl", "-k", "--since", f"@{int(since_timestamp)}", "--no-pager"],
                capture_output=True, text=True, timeout=10
            )
            for line in result.stdout.splitlines():
                line_lower = line.lower()
                if any(term in line_lower for term in ["out of memory", "oom", "killed"]):
                    oom_events.append(line.strip())
        except Exception:
            pass
    except Exception:
        pass
    
    return oom_events


def convert_file_to_pdf(
    input_path: Path,
    soffice_path: str,
    timeout: int = 120,
    use_unique_profile: bool = True,
    use_optimized_env: bool = True,
) -> ConversionResult:
    """
    Convert a single file to PDF using LibreOffice.
    
    Args:
        input_path: Path to input file
        soffice_path: Path to soffice executable
        timeout: Conversion timeout in seconds
        use_unique_profile: Use unique user profile (recommended)
        use_optimized_env: Use optimized environment variables
    
    Returns:
        ConversionResult with status and details
    """
    input_size_mb = input_path.stat().st_size / (1024**2)
    start_time = time.time()
    
    # Create temporary directory for this conversion
    with tempfile.TemporaryDirectory(prefix="lo_convert_") as temp_dir:
        temp_path = Path(temp_dir)
        
        # Copy input file to temp dir
        temp_input = temp_path / input_path.name
        shutil.copy(input_path, temp_input)
        
        # Expected output path
        output_name = temp_input.stem + ".pdf"
        expected_output = temp_path / output_name
        
        # Build command
        cmd = [
            soffice_path,
            "--headless",
            "--invisible",
            "--nodefault",
            "--nofirststartwizard",
            "--nolockcheck",
            "--nologo",
            "--norestore",
        ]
        
        # Add unique profile if enabled
        if use_unique_profile:
            profile_dir = temp_path / f"profile_{uuid.uuid4().hex[:8]}"
            profile_dir.mkdir()
            cmd.append(f"-env:UserInstallation=file://{profile_dir}")
        
        cmd.extend([
            "--convert-to", "pdf",
            "--outdir", str(temp_path),
            str(temp_input)
        ])
        
        # Environment
        env = get_optimized_env() if use_optimized_env else os.environ.copy()
        
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=timeout,
                env=env,
            )
            
            duration = time.time() - start_time
            
            # Check if PDF was created
            if expected_output.exists():
                output_size = expected_output.stat().st_size
                return ConversionResult(
                    filename=input_path.name,
                    input_size_mb=input_size_mb,
                    output_size_kb=output_size / 1024,
                    status="success",
                    duration_seconds=duration,
                )
            else:
                # Check for PDF with different name (LibreOffice quirk)
                pdf_files = list(temp_path.glob("*.pdf"))
                if pdf_files:
                    output_size = pdf_files[0].stat().st_size
                    return ConversionResult(
                        filename=input_path.name,
                        input_size_mb=input_size_mb,
                        output_size_kb=output_size / 1024,
                        status="success",
                        duration_seconds=duration,
                    )
                
                return ConversionResult(
                    filename=input_path.name,
                    input_size_mb=input_size_mb,
                    output_size_kb=0,
                    status="failed",
                    duration_seconds=duration,
                    error_message=f"No PDF output. stderr: {result.stderr[:500]}"
                )
                
        except subprocess.TimeoutExpired:
            duration = time.time() - start_time
            # Kill any lingering soffice processes
            kill_soffice_processes()
            return ConversionResult(
                filename=input_path.name,
                input_size_mb=input_size_mb,
                output_size_kb=0,
                status="timeout",
                duration_seconds=duration,
                error_message=f"Conversion timed out after {timeout}s"
            )
            
        except Exception as e:
            duration = time.time() - start_time
            return ConversionResult(
                filename=input_path.name,
                input_size_mb=input_size_mb,
                output_size_kb=0,
                status="failed",
                duration_seconds=duration,
                error_message=str(e)[:500]
            )


def kill_soffice_processes():
    """Kill any orphaned soffice processes."""
    try:
        for proc in psutil.process_iter(['name', 'cmdline']):
            try:
                if proc.info['name'] in ('soffice', 'soffice.bin'):
                    proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    except Exception:
        pass


def run_stress_test(
    test_files_dir: str,
    concurrent: int = 4,
    rounds: int = 2,
    timeout: int = 120,
    use_unique_profile: bool = True,
    use_optimized_env: bool = True,
    verbose: bool = True,
) -> StressTestReport:
    """
    Run a complete stress test suite.
    
    Args:
        test_files_dir: Directory containing test files
        concurrent: Number of concurrent conversions
        rounds: Number of rounds to repeat all files
        timeout: Timeout per conversion in seconds
        use_unique_profile: Use unique LibreOffice profile per conversion
        use_optimized_env: Use optimized environment variables
        verbose: Print progress output
    
    Returns:
        StressTestReport with all results
    """
    test_dir = Path(test_files_dir)
    
    if not test_dir.exists():
        raise ValueError(f"Test directory does not exist: {test_dir}")
    
    # Find test files
    test_files = (
        list(test_dir.glob("*.pptx")) +
        list(test_dir.glob("*.ppt")) +
        list(test_dir.glob("*.xlsx")) +
        list(test_dir.glob("*.xls")) +
        list(test_dir.glob("*.docx")) +
        list(test_dir.glob("*.doc"))
    )
    
    if not test_files:
        raise ValueError(f"No test files found in {test_dir}")
    
    # Find LibreOffice
    soffice_path = find_libreoffice_path()
    
    # Gather system info
    system_info = get_system_info()
    
    start_timestamp = time.time()
    start_time = datetime.now().isoformat()
    
    if verbose:
        print("=" * 70)
        print("LIBREOFFICE OOM STRESS TEST")
        print("=" * 70)
        print(f"System: {system_info['hostname']} ({system_info['cpu_count']} CPUs, "
              f"{system_info['total_memory_gb']} GB RAM)")
        print(f"LibreOffice: {system_info['libreoffice_version']}")
        print(f"Test files: {len(test_files)} files × {rounds} rounds = {len(test_files) * rounds} conversions")
        print(f"Concurrency: {concurrent} parallel conversions")
        print(f"Unique profiles: {'Yes' if use_unique_profile else 'No'}")
        print(f"Optimized env: {'Yes' if use_optimized_env else 'No'}")
        print("=" * 70)
    
    # Start memory monitor
    monitor = MemoryMonitor(interval=0.5)
    monitor.start()
    
    results: List[ConversionResult] = []
    
    # Build task queue
    tasks = []
    for round_num in range(rounds):
        for file_path in test_files:
            tasks.append((round_num, file_path))
    
    # Process with thread pool
    from concurrent.futures import ThreadPoolExecutor, as_completed
    
    completed = 0
    total_tasks = len(tasks)
    
    with ThreadPoolExecutor(max_workers=concurrent) as executor:
        futures = {}
        
        for round_num, file_path in tasks:
            future = executor.submit(
                convert_file_to_pdf,
                file_path,
                soffice_path,
                timeout,
                use_unique_profile,
                use_optimized_env,
            )
            futures[future] = (round_num, file_path)
        
        for future in as_completed(futures):
            round_num, file_path = futures[future]
            completed += 1
            
            try:
                result = future.result(timeout=timeout + 30)
                result.memory_peak_mb = monitor.peak_memory_mb
                results.append(result)
                
                if verbose:
                    status_icon = "✓" if result.status == "success" else "✗"
                    mem_stats = monitor.get_current()
                    print(f"[{completed}/{total_tasks}] {status_icon} {result.filename} "
                          f"({result.duration_seconds:.1f}s) | "
                          f"Mem: {mem_stats['used_mb']/1024:.1f}GB ({mem_stats['percent']}%) "
                          f"Peak: {mem_stats['peak_mb']/1024:.1f}GB")
                    
            except Exception as e:
                if verbose:
                    print(f"[{completed}/{total_tasks}] ✗ {file_path.name}: {e}")
                results.append(ConversionResult(
                    filename=file_path.name,
                    input_size_mb=file_path.stat().st_size / (1024**2),
                    output_size_kb=0,
                    status="failed",
                    duration_seconds=0,
                    error_message=str(e)[:500]
                ))
    
    # Stop monitor
    peak_memory = monitor.stop()
    
    # Check for OOM events
    oom_events = check_dmesg_for_oom(start_timestamp)
    
    end_time = datetime.now().isoformat()
    duration = time.time() - start_timestamp
    
    # Calculate summary
    successful = sum(1 for r in results if r.status == "success")
    failed = sum(1 for r in results if r.status == "failed")
    timeouts = sum(1 for r in results if r.status == "timeout")
    oom_count = sum(1 for r in results if r.status == "oom")
    
    avg_duration = sum(r.duration_seconds for r in results) / len(results) if results else 0
    
    summary = {
        "total_conversions": len(results),
        "successful": successful,
        "failed": failed,
        "timeouts": timeouts,
        "oom_killed": oom_count,
        "success_rate": round(successful / len(results) * 100, 2) if results else 0,
        "average_duration_seconds": round(avg_duration, 2),
        "peak_memory_gb": round(peak_memory / 1024, 2),
        "oom_events_detected": len(oom_events),
    }
    
    report = StressTestReport(
        start_time=start_time,
        end_time=end_time,
        duration_seconds=round(duration, 2),
        system_info=system_info,
        test_config={
            "test_files_dir": str(test_dir.absolute()),
            "file_count": len(test_files),
            "rounds": rounds,
            "concurrent": concurrent,
            "timeout": timeout,
            "use_unique_profile": use_unique_profile,
            "use_optimized_env": use_optimized_env,
        },
        results=[asdict(r) for r in results],
        summary=summary,
        oom_events=oom_events,
    )
    
    if verbose:
        print("\n" + "=" * 70)
        print("STRESS TEST RESULTS")
        print("=" * 70)
        print(f"Duration: {duration:.1f} seconds")
        print(f"Total conversions: {summary['total_conversions']}")
        print(f"  ✓ Successful: {summary['successful']}")
        print(f"  ✗ Failed: {summary['failed']}")
        print(f"  ⏱ Timeouts: {summary['timeouts']}")
        print(f"  💀 OOM killed: {summary['oom_killed']}")
        print(f"Success rate: {summary['success_rate']}%")
        print(f"Average duration: {summary['average_duration_seconds']}s")
        print(f"Peak memory: {summary['peak_memory_gb']} GB")
        
        if oom_events:
            print(f"\n⚠️  OOM EVENTS DETECTED ({len(oom_events)}):")
            for event in oom_events[:5]:
                print(f"  {event[:100]}")
        
        print("=" * 70)
    
    return report


def save_report(report: StressTestReport, output_path: str):
    """Save the stress test report to a JSON file."""
    with open(output_path, 'w') as f:
        json.dump(asdict(report), f, indent=2, default=str)
    print(f"Report saved to: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Run LibreOffice PDF conversion stress tests",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --test-dir ./test_files --concurrent 2
  %(prog)s --test-dir ./test_files --concurrent 4 --rounds 3
  %(prog)s --test-dir ./test_files --no-unique-profile  # Test without fix
  %(prog)s --test-dir ./test_files --output report.json

Compare with and without fixes:
  %(prog)s --test-dir ./test_files --no-unique-profile --no-optimized-env --output before_fix.json
  %(prog)s --test-dir ./test_files --output after_fix.json
        """
    )
    
    parser.add_argument(
        "--test-dir", "-d",
        default="./test_files",
        help="Directory containing test files (default: ./test_files)"
    )
    parser.add_argument(
        "--concurrent", "-c",
        type=int, default=4,
        help="Number of concurrent conversions (default: 4)"
    )
    parser.add_argument(
        "--rounds", "-r",
        type=int, default=2,
        help="Number of rounds to repeat test files (default: 2)"
    )
    parser.add_argument(
        "--timeout", "-t",
        type=int, default=120,
        help="Timeout per conversion in seconds (default: 120)"
    )
    parser.add_argument(
        "--no-unique-profile",
        action="store_true",
        help="Disable unique profile per conversion (to test without fix)"
    )
    parser.add_argument(
        "--no-optimized-env",
        action="store_true",
        help="Disable optimized environment variables (to test without fix)"
    )
    parser.add_argument(
        "--output", "-o",
        help="Output path for JSON report (optional)"
    )
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Suppress progress output"
    )
    
    args = parser.parse_args()
    
    try:
        report = run_stress_test(
            test_files_dir=args.test_dir,
            concurrent=args.concurrent,
            rounds=args.rounds,
            timeout=args.timeout,
            use_unique_profile=not args.no_unique_profile,
            use_optimized_env=not args.no_optimized_env,
            verbose=not args.quiet,
        )
        
        if args.output:
            save_report(report, args.output)
        
        # Exit with error code if there were failures
        if report.summary["failed"] > 0 or report.summary["oom_killed"] > 0:
            sys.exit(1)
            
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
