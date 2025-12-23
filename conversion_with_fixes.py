#!/usr/bin/env python3
"""
LibreOffice PDF Conversion Module with OOM Fixes

This module provides optimized functions for converting PowerPoint and Excel
files to PDF using LibreOffice in headless mode, with all known fixes for
OOM (Out of Memory) errors.

Key optimizations:
1. Unique user profile per conversion (prevents conflicts)
2. Optimized environment variables (reduces memory)
3. Proper cleanup of temporary files and profiles
4. Configurable memory limits via systemd-run (optional)

Usage:
    from conversion_with_fixes import convert_powerpoint_to_pdf, convert_excel_to_pdf
    
    pdf_bytes = convert_powerpoint_to_pdf(pptx_bytes)
    pdf_bytes = convert_excel_to_pdf(xlsx_bytes)
"""

import logging
import os
import shutil
import subprocess
import tempfile
import uuid
from pathlib import Path
from typing import Optional, Dict

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def find_libreoffice_path() -> str:
    """
    Find the LibreOffice soffice executable.
    
    Returns:
        Path to soffice executable
        
    Raises:
        RuntimeError: If LibreOffice is not found
    """
    possible_paths = [
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/usr/local/bin/soffice",
        "/opt/libreoffice/program/soffice",
        "/opt/libreoffice7.6/program/soffice",
        "/opt/libreoffice24.2/program/soffice",
        "/snap/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS
    ]
    
    # Also check PATH
    which_result = shutil.which("soffice") or shutil.which("libreoffice")
    if which_result:
        possible_paths.insert(0, which_result)
    
    for path in possible_paths:
        if path and os.path.isfile(path) and os.access(path, os.X_OK):
            return path
    
    raise RuntimeError(
        "LibreOffice not found. Please install it:\n"
        "  Ubuntu/Debian: sudo apt install libreoffice\n"
        "  RHEL/CentOS:   sudo dnf install libreoffice\n"
        "  Arch:          sudo pacman -S libreoffice-fresh\n"
        "  macOS:         brew install --cask libreoffice"
    )


def get_optimized_env_for_libreoffice() -> Dict[str, str]:
    """
    Get environment variables optimized for LibreOffice headless conversion.
    
    These settings disable various features that consume memory but aren't
    needed for headless PDF conversion.
    
    Returns:
        Dictionary of environment variables
    """
    env = os.environ.copy()
    
    # ===========================================
    # CRITICAL: Disable GPU/Hardware Acceleration
    # ===========================================
    # These are major memory consumers in headless mode
    env["SAL_DISABLE_OPENCL"] = "1"      # Disable OpenCL GPU acceleration
    env["SAL_DISABLEGL"] = "1"           # Disable OpenGL
    env["SAL_DISABLESKIA"] = "1"         # Disable Skia graphics library
    
    # ===========================================
    # VCL Plugin Configuration
    # ===========================================
    # Use generic (headless) VCL plugin - no X11/Wayland dependencies
    env["SAL_USE_VCLPLUGIN"] = "gen"
    
    # ===========================================
    # Disable Crash Reporting & Recovery
    # ===========================================
    env["SAL_NO_CRASHREPORT"] = "1"      # No crash report dialogs
    env["SAL_DISABLE_WATCHDOG"] = "1"    # Disable watchdog timer
    
    # ===========================================
    # Java Configuration (if Java is used)
    # ===========================================
    # Limit Java heap size to prevent JVM from consuming too much memory
    env["JAVA_TOOL_OPTIONS"] = "-Xmx256m -Xms64m"
    
    # ===========================================
    # Additional Memory Optimizations
    # ===========================================
    # Disable file locking (not needed for temp files)
    env["SAL_ENABLE_FILE_LOCKING"] = "0"
    
    # Reduce font cache
    env["SAL_FONTCONFIG_CACHE_DISABLE"] = "1"
    
    return env


def cleanup_soffice_processes(profile_dir: Optional[str] = None):
    """
    Clean up any orphaned soffice processes.
    
    Args:
        profile_dir: If provided, only kill processes using this profile
    """
    try:
        import psutil
        for proc in psutil.process_iter(['name', 'cmdline']):
            try:
                if proc.info['name'] in ('soffice', 'soffice.bin'):
                    cmdline = proc.info.get('cmdline', [])
                    if profile_dir:
                        # Only kill if using our profile
                        if any(profile_dir in arg for arg in cmdline if arg):
                            proc.kill()
                    # Don't kill all soffice processes - might be user's
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    except ImportError:
        # psutil not available, try pkill as fallback
        if profile_dir:
            try:
                subprocess.run(
                    ["pkill", "-f", profile_dir],
                    capture_output=True,
                    timeout=5
                )
            except Exception:
                pass


def convert_document_to_pdf(
    document_bytes: bytes,
    input_extension: str,
    libreoffice_path: Optional[str] = None,
    timeout: int = 120,
    use_unique_profile: bool = True,
    use_optimized_env: bool = True,
    use_memory_limit: bool = False,
    memory_limit_mb: int = 1024,
) -> bytes:
    """
    Convert a document to PDF using LibreOffice.
    
    This is the core conversion function with all OOM fixes applied.
    
    Args:
        document_bytes: The document content as bytes
        input_extension: File extension (e.g., 'pptx', 'xlsx', 'docx')
        libreoffice_path: Custom path to soffice (optional)
        timeout: Conversion timeout in seconds
        use_unique_profile: Use unique user profile per conversion (recommended)
        use_optimized_env: Use memory-optimized environment variables (recommended)
        use_memory_limit: Use systemd-run for hard memory limits (Linux only)
        memory_limit_mb: Memory limit in MB when use_memory_limit is True
    
    Returns:
        PDF content as bytes
        
    Raises:
        RuntimeError: If conversion fails
    """
    soffice_path = libreoffice_path or find_libreoffice_path()
    
    # Normalize extension
    ext = input_extension.lower().lstrip('.')
    
    # Generate unique ID for this conversion
    conversion_id = uuid.uuid4().hex[:12]
    
    logger.info(f"Starting conversion [{conversion_id}]: {ext} -> PDF ({len(document_bytes)} bytes)")
    
    try:
        with tempfile.TemporaryDirectory(prefix=f"lo_convert_{conversion_id}_") as temp_dir:
            temp_path = Path(temp_dir)
            
            # Input and output paths
            input_file = temp_path / f"input.{ext}"
            expected_output = temp_path / "input.pdf"
            
            # Write input file
            input_file.write_bytes(document_bytes)
            
            # Create unique user profile directory
            profile_dir = None
            if use_unique_profile:
                profile_dir = temp_path / f"profile_{conversion_id}"
                profile_dir.mkdir()
                logger.debug(f"Using unique profile: {profile_dir}")
            
            # Build command with all recommended flags
            cmd = [
                soffice_path,
                "--headless",           # No GUI
                "--invisible",          # No window at all
                "--nodefault",          # Don't create default document
                "--nofirststartwizard", # Skip first-run wizard
                "--nolockcheck",        # Don't check for lock files
                "--nologo",             # No splash screen
                "--norestore",          # Don't restore crashed documents
            ]
            
            # Add unique profile if enabled
            if use_unique_profile and profile_dir:
                cmd.append(f"-env:UserInstallation=file://{profile_dir}")
            
            # Add conversion parameters
            cmd.extend([
                "--convert-to", "pdf",
                "--outdir", str(temp_path),
                str(input_file)
            ])
            
            # Wrap with systemd-run for memory limiting (Linux only)
            if use_memory_limit:
                systemd_cmd = [
                    "systemd-run",
                    "--scope",
                    "--user",  # Run as user, not system
                    "-p", f"MemoryMax={memory_limit_mb}M",
                    "-p", f"MemoryHigh={int(memory_limit_mb * 0.8)}M",  # Soft limit at 80%
                ]
                cmd = systemd_cmd + cmd
            
            # Environment
            env = get_optimized_env_for_libreoffice() if use_optimized_env else os.environ.copy()
            
            logger.debug(f"Executing: {' '.join(cmd[:10])}...")
            
            # Run conversion
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=timeout,
                env=env,
                cwd=temp_dir,  # Set working directory
            )
            
            # Log any warnings/errors
            if result.returncode != 0:
                logger.warning(f"LibreOffice returned code {result.returncode}")
                if result.stderr:
                    logger.warning(f"stderr: {result.stderr[:500]}")
            
            # Check for output file
            if not expected_output.exists():
                # LibreOffice sometimes names output differently
                pdf_files = list(temp_path.glob("*.pdf"))
                if pdf_files:
                    expected_output = pdf_files[0]
                else:
                    raise RuntimeError(
                        f"PDF output file not generated. "
                        f"returncode={result.returncode}, "
                        f"stderr={result.stderr[:500] if result.stderr else 'none'}"
                    )
            
            # Read output
            pdf_bytes = expected_output.read_bytes()
            
            if not pdf_bytes:
                raise RuntimeError("Generated PDF is empty")
            
            if not pdf_bytes.startswith(b'%PDF'):
                logger.warning("Generated file may not be a valid PDF (missing header)")
            
            logger.info(f"Conversion successful [{conversion_id}]: {len(pdf_bytes)} bytes PDF")
            
            return pdf_bytes
            
    except subprocess.TimeoutExpired:
        logger.error(f"Conversion timed out after {timeout}s [{conversion_id}]")
        # Cleanup orphaned processes
        if profile_dir:
            cleanup_soffice_processes(str(profile_dir))
        raise RuntimeError(f"Conversion timed out after {timeout} seconds")
        
    except Exception as e:
        logger.error(f"Conversion failed [{conversion_id}]: {e}")
        raise


def convert_powerpoint_to_pdf(
    document_bytes: bytes,
    libreoffice_path: Optional[str] = None,
    timeout: int = 120,
    **kwargs
) -> bytes:
    """
    Convert PowerPoint (PPTX/PPT) to PDF.
    
    Args:
        document_bytes: PowerPoint file content as bytes
        libreoffice_path: Custom path to LibreOffice (optional)
        timeout: Conversion timeout in seconds
        **kwargs: Additional arguments passed to convert_document_to_pdf
    
    Returns:
        PDF content as bytes
    """
    # Detect format from magic bytes
    if document_bytes[:4] == b'PK\x03\x04':
        ext = 'pptx'  # OOXML format
    else:
        ext = 'ppt'   # Legacy format
    
    return convert_document_to_pdf(
        document_bytes=document_bytes,
        input_extension=ext,
        libreoffice_path=libreoffice_path,
        timeout=timeout,
        **kwargs
    )


def convert_excel_to_pdf(
    document_bytes: bytes,
    libreoffice_path: Optional[str] = None,
    timeout: int = 120,
    **kwargs
) -> bytes:
    """
    Convert Excel (XLSX/XLS) to PDF.
    
    Args:
        document_bytes: Excel file content as bytes
        libreoffice_path: Custom path to LibreOffice (optional)
        timeout: Conversion timeout in seconds
        **kwargs: Additional arguments passed to convert_document_to_pdf
    
    Returns:
        PDF content as bytes
    """
    # Detect format from magic bytes
    if document_bytes[:4] == b'PK\x03\x04':
        ext = 'xlsx'  # OOXML format
    else:
        ext = 'xls'   # Legacy format
    
    return convert_document_to_pdf(
        document_bytes=document_bytes,
        input_extension=ext,
        libreoffice_path=libreoffice_path,
        timeout=timeout,
        **kwargs
    )


def convert_word_to_pdf(
    document_bytes: bytes,
    libreoffice_path: Optional[str] = None,
    timeout: int = 120,
    **kwargs
) -> bytes:
    """
    Convert Word (DOCX/DOC) to PDF.
    
    Args:
        document_bytes: Word file content as bytes
        libreoffice_path: Custom path to LibreOffice (optional)
        timeout: Conversion timeout in seconds
        **kwargs: Additional arguments passed to convert_document_to_pdf
    
    Returns:
        PDF content as bytes
    """
    # Detect format from magic bytes
    if document_bytes[:4] == b'PK\x03\x04':
        ext = 'docx'  # OOXML format
    else:
        ext = 'doc'   # Legacy format
    
    return convert_document_to_pdf(
        document_bytes=document_bytes,
        input_extension=ext,
        libreoffice_path=libreoffice_path,
        timeout=timeout,
        **kwargs
    )


# =============================================================================
# EXAMPLE USAGE AND TESTING
# =============================================================================

if __name__ == "__main__":
    import sys
    
    print("LibreOffice Conversion Module with OOM Fixes")
    print("=" * 50)
    
    # Check LibreOffice installation
    try:
        lo_path = find_libreoffice_path()
        print(f"✓ LibreOffice found: {lo_path}")
    except RuntimeError as e:
        print(f"✗ {e}")
        sys.exit(1)
    
    # Show optimized environment
    print("\nOptimized environment variables:")
    env = get_optimized_env_for_libreoffice()
    for key in sorted(env.keys()):
        if key.startswith("SAL_") or key.startswith("JAVA"):
            print(f"  {key}={env[key]}")
    
    # Test with a file if provided
    if len(sys.argv) > 1:
        input_file = Path(sys.argv[1])
        if input_file.exists():
            print(f"\nConverting: {input_file}")
            
            file_bytes = input_file.read_bytes()
            ext = input_file.suffix.lower()
            
            try:
                pdf_bytes = convert_document_to_pdf(
                    document_bytes=file_bytes,
                    input_extension=ext,
                    timeout=120,
                    use_unique_profile=True,
                    use_optimized_env=True,
                )
                
                output_file = input_file.with_suffix('.pdf')
                output_file.write_bytes(pdf_bytes)
                print(f"✓ Output saved: {output_file} ({len(pdf_bytes)} bytes)")
                
            except Exception as e:
                print(f"✗ Conversion failed: {e}")
                sys.exit(1)
        else:
            print(f"File not found: {input_file}")
            sys.exit(1)
    else:
        print("\nUsage: python conversion_with_fixes.py <input_file>")
        print("Example: python conversion_with_fixes.py presentation.pptx")
