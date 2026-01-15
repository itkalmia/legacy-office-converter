import os
import subprocess
import tempfile
import uuid
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

app = FastAPI(
    title="Legacy Office Converter",
    version="1.0.0",
    description="Converts legacy Office files to modern formats."
)

# LibreOffice binary path (configurable via environment variable)
LIBREOFFICE_BINARY = os.getenv("LIBREOFFICE_BINARY", "soffice")

# Mapping of legacy formats to modern formats
FORMAT_MAPPING = {
    ".doc": ".docx",
    ".xls": ".xlsx",
    ".ppt": ".pptx",
}

# LibreOffice format filters
LIBREOFFICE_FILTERS = {
    ".docx": "docx",
    ".xlsx": "xlsx",
    ".pptx": "pptx",
}


def is_libreoffice_available() -> bool:
    """Check if LibreOffice is available and working by calling help."""
    try:
        result = subprocess.run(
            [LIBREOFFICE_BINARY, "--help"],
            capture_output=True,
            timeout=5,
        )
        return result.returncode == 0
    except (subprocess.TimeoutExpired, FileNotFoundError, Exception):
        print(Exception, "LibreOffice not available or not working")
        return False


def convert_with_libreoffice(input_path: Path, output_dir: Path, output_format: str) -> Optional[Path]:
    """
    Convert a file using LibreOffice in headless mode.

    Args:
        input_path: Path to the input file
        output_dir: Directory where the output file will be created
        output_format: Target format (e.g., 'docx', 'xlsx', 'pptx')

    Returns:
        Path to the converted file or None if conversion failed
    """
    try:
        # Run LibreOffice in headless mode
        cmd = [
            LIBREOFFICE_BINARY,
            "--headless",
            "--convert-to",
            output_format,
            "--outdir",
            str(output_dir),
            str(input_path),
        ]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60,
        )

        if result.returncode != 0:
            print(f"LibreOffice conversion failed: {result.stderr}")
            return None

        # Find the converted file
        expected_name = input_path.stem + "." + output_format
        output_path = output_dir / expected_name

        if output_path.exists():
            return output_path

        return None

    except subprocess.TimeoutExpired:
        print("LibreOffice conversion timed out")
        return None
    except Exception as e:
        print(f"Error during conversion: {str(e)}")
        return None


@app.get("/")
async def root():
    libreoffice_ok = is_libreoffice_available()
    return {
        "status": "healthy" if libreoffice_ok else "unhealthy",
        "service": "Legacy Office Converter",
        "supported_formats": list(FORMAT_MAPPING.keys()),
    }


@app.post("/convert")
async def convert_file(file: UploadFile = File(...)):
    """
    Convert a legacy Office file to its modern equivalent.

    Supported conversions:
    - .doc → .docx
    - .xls → .xlsx
    - .ppt → .pptx
    """
    # Get file extension
    file_ext = Path(file.filename).suffix.lower()

    # Validate file format
    if file_ext not in FORMAT_MAPPING:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported file format. Supported formats: {', '.join(FORMAT_MAPPING.keys())}",
        )

    # Determine output format
    output_ext = FORMAT_MAPPING[file_ext]
    output_format = LIBREOFFICE_FILTERS[output_ext]

    # Create temporary directory for processing
    temp_dir = Path(tempfile.mkdtemp())

    try:
        # Generate unique filename to avoid conflicts
        unique_id = str(uuid.uuid4())
        input_filename = f"{unique_id}{file_ext}"
        input_path = temp_dir / input_filename

        # Save uploaded file
        content = await file.read()
        input_path.write_bytes(content)

        # Convert file
        output_path = convert_with_libreoffice(input_path, temp_dir, output_format)

        if output_path is None or not output_path.exists():
            raise HTTPException(
                status_code=500,
                detail="Conversion failed. Please check if the file is valid.",
            )

        # Determine output filename
        original_stem = Path(file.filename).stem
        output_filename = f"{original_stem}{output_ext}"

        # Return the converted file
        return FileResponse(
            path=str(output_path),
            filename=output_filename,
            media_type="application/octet-stream",
            background=None,
        )

    except HTTPException:
        # Clean up temp directory
        for f in temp_dir.glob("*"):
            f.unlink()
        temp_dir.rmdir()
        raise
    except Exception as e:
        # Clean up temp directory
        for f in temp_dir.glob("*"):
            f.unlink()
        temp_dir.rmdir()
        raise HTTPException(status_code=500, detail=f"An error occurred: {str(e)}")


@app.get("/health")
async def health():
    """Health check endpoint for monitoring and container orchestration."""
    libreoffice_ok = is_libreoffice_available()
    if not libreoffice_ok:
        raise HTTPException(
            status_code=503,
            detail={"status": "unhealthy", "reason": "LibreOffice binary not responding"}
        )
    return {"status": "healthy"}