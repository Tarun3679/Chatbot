#!/bin/bash
#
# LibreOffice OOM Test Suite Runner
# 
# This script automates the complete test workflow:
# 1. Check dependencies
# 2. Generate test files
# 3. Run stress tests (before and after fixes)
# 4. Compare results
#
# Usage:
#   ./run_tests.sh [light|medium|heavy|extreme]
#
# Default stress level: medium
#

set -e  # Exit on error

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Configuration
STRESS_LEVEL="${1:-medium}"
TEST_DIR="./test_files"
REPORTS_DIR="./reports"
CONCURRENT=4
ROUNDS=2
TIMEOUT=120

# Print colored header
print_header() {
    echo ""
    echo -e "${BLUE}======================================================================${NC}"
    echo -e "${BLUE}  $1${NC}"
    echo -e "${BLUE}======================================================================${NC}"
    echo ""
}

# Print status message
print_status() {
    echo -e "${GREEN}[✓]${NC} $1"
}

# Print warning
print_warning() {
    echo -e "${YELLOW}[!]${NC} $1"
}

# Print error
print_error() {
    echo -e "${RED}[✗]${NC} $1"
}

# Check if command exists
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# =============================================================================
# STEP 1: Check System Requirements
# =============================================================================
print_header "STEP 1: Checking System Requirements"

# Check Python
if command_exists python3; then
    PYTHON=python3
elif command_exists python; then
    PYTHON=python
else
    print_error "Python not found. Please install Python 3."
    exit 1
fi
print_status "Python: $($PYTHON --version)"

# Check LibreOffice
if command_exists soffice; then
    LO_VERSION=$(soffice --version 2>/dev/null || echo "unknown")
    print_status "LibreOffice: $LO_VERSION"
elif command_exists libreoffice; then
    LO_VERSION=$(libreoffice --version 2>/dev/null || echo "unknown")
    print_status "LibreOffice: $LO_VERSION"
else
    print_error "LibreOffice not found. Please install it:"
    echo "  Ubuntu/Debian: sudo apt install libreoffice"
    echo "  RHEL/CentOS:   sudo dnf install libreoffice"
    echo "  Arch:          sudo pacman -S libreoffice-fresh"
    exit 1
fi

# Check/Install Python dependencies
print_status "Checking Python dependencies..."

$PYTHON << 'EOF'
import sys
missing = []

try:
    import psutil
except ImportError:
    missing.append("psutil")

try:
    from pptx import Presentation
except ImportError:
    missing.append("python-pptx")

try:
    from openpyxl import Workbook
except ImportError:
    missing.append("openpyxl")

try:
    from PIL import Image
except ImportError:
    missing.append("pillow")

try:
    import numpy
except ImportError:
    missing.append("numpy")

if missing:
    print("MISSING:" + ",".join(missing))
    sys.exit(1)
else:
    print("OK")
    sys.exit(0)
EOF

if [ $? -ne 0 ]; then
    print_warning "Installing missing Python dependencies..."
    $PYTHON -m pip install psutil python-pptx openpyxl pillow numpy --quiet
    print_status "Dependencies installed"
fi

# Show system info
echo ""
echo "System Information:"
echo "  - OS: $(uname -s) $(uname -r)"
echo "  - CPU: $(nproc) cores"
echo "  - RAM: $(free -h | awk '/^Mem:/ {print $2}')"
echo "  - Swap: $(free -h | awk '/^Swap:/ {print $2}')"

# =============================================================================
# STEP 2: Generate Test Files
# =============================================================================
print_header "STEP 2: Generating Test Files (Level: $STRESS_LEVEL)"

mkdir -p "$TEST_DIR"
mkdir -p "$REPORTS_DIR"

$PYTHON test_file_generator.py --output-dir "$TEST_DIR" --level "$STRESS_LEVEL"

# Show generated files
echo ""
echo "Generated files:"
ls -lh "$TEST_DIR"/*.pptx "$TEST_DIR"/*.xlsx 2>/dev/null | awk '{print "  " $9 " (" $5 ")"}'

# =============================================================================
# STEP 3: Run Stress Test WITHOUT Fixes (Baseline)
# =============================================================================
print_header "STEP 3: Running Stress Test WITHOUT Fixes (Baseline)"

print_warning "This test intentionally uses the OLD configuration to establish a baseline"
print_warning "You may see failures or OOM errors - this is expected!"
echo ""

BASELINE_REPORT="$REPORTS_DIR/baseline_$(date +%Y%m%d_%H%M%S).json"

$PYTHON stress_test_runner.py \
    --test-dir "$TEST_DIR" \
    --concurrent "$CONCURRENT" \
    --rounds "$ROUNDS" \
    --timeout "$TIMEOUT" \
    --no-unique-profile \
    --no-optimized-env \
    --output "$BASELINE_REPORT" \
    || true  # Don't exit on failure

print_status "Baseline report saved: $BASELINE_REPORT"

# Wait a bit for system to recover
echo ""
print_status "Waiting 10 seconds for system to stabilize..."
sleep 10

# =============================================================================
# STEP 4: Run Stress Test WITH Fixes
# =============================================================================
print_header "STEP 4: Running Stress Test WITH Fixes"

print_status "This test uses all OOM fixes applied"
echo ""

FIXED_REPORT="$REPORTS_DIR/with_fixes_$(date +%Y%m%d_%H%M%S).json"

$PYTHON stress_test_runner.py \
    --test-dir "$TEST_DIR" \
    --concurrent "$CONCURRENT" \
    --rounds "$ROUNDS" \
    --timeout "$TIMEOUT" \
    --output "$FIXED_REPORT" \
    || true

print_status "Fixed version report saved: $FIXED_REPORT"

# =============================================================================
# STEP 5: Compare Results
# =============================================================================
print_header "STEP 5: Comparing Results"

$PYTHON << EOF
import json
from pathlib import Path

def load_report(path):
    with open(path) as f:
        return json.load(f)

baseline = load_report("$BASELINE_REPORT")
fixed = load_report("$FIXED_REPORT")

print("=" * 60)
print(f"{'Metric':<30} {'Baseline':>12} {'With Fixes':>12}")
print("=" * 60)

metrics = [
    ("Total Conversions", "total_conversions"),
    ("Successful", "successful"),
    ("Failed", "failed"),
    ("Timeouts", "timeouts"),
    ("OOM Killed", "oom_killed"),
    ("Success Rate (%)", "success_rate"),
    ("Avg Duration (s)", "average_duration_seconds"),
    ("Peak Memory (GB)", "peak_memory_gb"),
]

for label, key in metrics:
    b_val = baseline["summary"].get(key, 0)
    f_val = fixed["summary"].get(key, 0)
    
    # Color coding for improvement
    if key in ["successful", "success_rate"]:
        improved = f_val > b_val
    elif key in ["failed", "timeouts", "oom_killed", "average_duration_seconds", "peak_memory_gb"]:
        improved = f_val < b_val
    else:
        improved = None
    
    if improved is True:
        indicator = "✓"
    elif improved is False:
        indicator = "✗"
    else:
        indicator = " "
    
    print(f"{label:<30} {str(b_val):>12} {str(f_val):>12} {indicator}")

print("=" * 60)

# OOM Events
baseline_oom = len(baseline.get("oom_events", []))
fixed_oom = len(fixed.get("oom_events", []))
print(f"{'OOM Events in dmesg':<30} {baseline_oom:>12} {fixed_oom:>12}")

print("=" * 60)

# Summary
b_success = baseline["summary"]["success_rate"]
f_success = fixed["summary"]["success_rate"]
improvement = f_success - b_success

if improvement > 0:
    print(f"\n✓ SUCCESS RATE IMPROVED by {improvement:.1f}%")
elif improvement < 0:
    print(f"\n✗ Success rate decreased by {abs(improvement):.1f}%")
else:
    print(f"\n○ Success rate unchanged")

b_mem = baseline["summary"]["peak_memory_gb"]
f_mem = fixed["summary"]["peak_memory_gb"]
mem_reduction = ((b_mem - f_mem) / b_mem * 100) if b_mem > 0 else 0

if mem_reduction > 0:
    print(f"✓ PEAK MEMORY REDUCED by {mem_reduction:.1f}%")
elif mem_reduction < 0:
    print(f"✗ Peak memory increased by {abs(mem_reduction):.1f}%")

EOF

# =============================================================================
# STEP 6: Check dmesg for OOM Events
# =============================================================================
print_header "STEP 6: Checking for OOM Events in System Logs"

echo "Recent OOM-related messages (if any):"
dmesg 2>/dev/null | grep -i -E "(out of memory|oom|killed process)" | tail -20 || \
    journalctl -k --since "1 hour ago" 2>/dev/null | grep -i -E "(oom|killed)" | tail -20 || \
    echo "  (No OOM events found or insufficient permissions to read logs)"

# =============================================================================
# DONE
# =============================================================================
print_header "TESTING COMPLETE"

echo "Reports saved in: $REPORTS_DIR/"
ls -la "$REPORTS_DIR"/*.json 2>/dev/null

echo ""
echo "Next steps:"
echo "  1. Review the comparison above"
echo "  2. Check detailed reports in $REPORTS_DIR/"
echo "  3. If OOM errors persist, try:"
echo "     - Reducing --concurrent (currently: $CONCURRENT)"
echo "     - Using systemd-run memory limits in conversion_with_fixes.py"
echo "     - Consider using Gotenberg instead of direct LibreOffice"
echo ""
