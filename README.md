# LibreOffice OOM Test Suite

A comprehensive test suite for reproducing and fixing LibreOffice Out-of-Memory (OOM) errors during PDF conversion.

## Overview

This test suite helps you:

1. **Generate test files** - Create PowerPoint and Excel files of varying sizes to stress-test LibreOffice
2. **Run stress tests** - Execute concurrent conversions to reproduce OOM errors
3. **Compare results** - See the difference before and after applying fixes
4. **Monitor memory** - Track memory usage in real-time during tests

## Quick Start

```bash
# 1. Install dependencies
pip install psutil python-pptx openpyxl pillow numpy

# 2. Make the test script executable
chmod +x run_tests.sh

# 3. Run the complete test suite
./run_tests.sh medium
```

## Files Included

| File | Description |
|------|-------------|
| `run_tests.sh` | Main script that runs the complete test workflow |
| `test_file_generator.py` | Generates test PowerPoint and Excel files |
| `stress_test_runner.py` | Runs stress tests with memory monitoring |
| `conversion_with_fixes.py` | Conversion module with all OOM fixes applied |

## Stress Levels

| Level | PowerPoint | Excel | Use Case |
|-------|------------|-------|----------|
| `light` | 5-10 slides, few images | 500-1000 rows | Quick validation |
| `medium` | 20-30 slides, 1-2 images each | 5000-8000 rows | Normal testing |
| `heavy` | 50-75 slides, 3+ images | 15000-20000 rows | Serious stress test |
| `extreme` | 100-150 slides, 4+ images | 40000-50000 rows | Maximum stress (may cause OOM) |

## Usage Examples

### Generate Test Files Only

```bash
# Generate medium-complexity test files
python test_file_generator.py --level medium --output-dir ./test_files

# Generate extreme test files (warning: may be large)
python test_file_generator.py --level extreme --output-dir ./test_files
```

### Run Stress Tests

```bash
# Test WITHOUT fixes (to establish baseline)
python stress_test_runner.py \
    --test-dir ./test_files \
    --concurrent 4 \
    --no-unique-profile \
    --no-optimized-env \
    --output baseline.json

# Test WITH fixes
python stress_test_runner.py \
    --test-dir ./test_files \
    --concurrent 4 \
    --output with_fixes.json
```

### Convert a Single File

```bash
# Using the fixed conversion module
python conversion_with_fixes.py presentation.pptx
```

## The OOM Fixes

The `conversion_with_fixes.py` module implements these key fixes:

### 1. Unique User Profile Per Conversion

```python
# Each conversion gets its own profile directory
profile_dir = temp_path / f"profile_{uuid.uuid4().hex}"
cmd.append(f"-env:UserInstallation=file://{profile_dir}")
```

This prevents:
- Profile corruption from concurrent access
- Memory leaks from shared profile data
- Lock file conflicts

### 2. Optimized Command-Line Flags

```python
cmd = [
    soffice_path,
    "--headless",           # No GUI
    "--invisible",          # No window
    "--nodefault",          # No default document
    "--nofirststartwizard", # Skip wizard
    "--nolockcheck",        # No lock checking
    "--nologo",             # No splash
    "--norestore",          # No crash recovery
]
```

### 3. Memory-Optimized Environment Variables

```python
env = {
    "SAL_DISABLE_OPENCL": "1",    # Disable OpenCL GPU
    "SAL_DISABLEGL": "1",         # Disable OpenGL
    "SAL_DISABLESKIA": "1",       # Disable Skia
    "SAL_USE_VCLPLUGIN": "gen",   # Generic (headless) VCL
    "SAL_NO_CRASHREPORT": "1",    # No crash dialogs
    "JAVA_TOOL_OPTIONS": "-Xmx256m",  # Limit Java heap
}
```

### 4. Optional: Hard Memory Limits with systemd-run

```python
# Wrap LibreOffice with systemd memory limits
cmd = [
    "systemd-run", "--scope", "--user",
    "-p", "MemoryMax=1024M",
    "-p", "MemoryHigh=800M",
    *original_cmd
]
```

## Interpreting Results

After running tests, compare these metrics:

| Metric | Good Sign |
|--------|-----------|
| Success Rate | Higher is better |
| Failed | Lower is better |
| Timeouts | Lower is better |
| OOM Killed | Should be 0 |
| Peak Memory | Lower is better |
| Avg Duration | Lower is better |

## Troubleshooting

### OOM Errors Still Occurring

1. **Reduce concurrency**:
   ```bash
   python stress_test_runner.py --concurrent 2
   ```

2. **Enable memory limits**:
   ```python
   convert_powerpoint_to_pdf(
       data,
       use_memory_limit=True,
       memory_limit_mb=1024
   )
   ```

3. **Increase swap space**:
   ```bash
   sudo fallocate -l 4G /swapfile
   sudo chmod 600 /swapfile
   sudo mkswap /swapfile
   sudo swapon /swapfile
   ```

### Permission Errors

```bash
# If dmesg requires root:
sudo dmesg | grep -i oom

# Or use journalctl:
journalctl -k --since "1 hour ago" | grep -i oom
```

### LibreOffice Not Found

```bash
# Ubuntu/Debian
sudo apt update && sudo apt install libreoffice

# RHEL/CentOS/Fedora
sudo dnf install libreoffice

# Arch
sudo pacman -S libreoffice-fresh
```

## Alternative: Gotenberg

If LibreOffice continues to cause issues, consider [Gotenberg](https://gotenberg.dev/):

```bash
# Run Gotenberg container
docker run --rm -p 3000:3000 gotenberg/gotenberg:8

# Convert via API
curl --request POST \
  --url http://localhost:3000/forms/libreoffice/convert \
  --form files=@document.pptx \
  -o output.pdf
```

## System Requirements

- **OS**: Linux (Ubuntu 18.04+, Debian 10+, RHEL 7+, etc.)
- **RAM**: Minimum 4GB, recommended 8GB+
- **Python**: 3.8+
- **LibreOffice**: 6.0+ (7.0+ recommended)

## License

MIT License - Use freely for testing and production.
