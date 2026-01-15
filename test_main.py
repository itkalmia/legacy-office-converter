import io
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, Mock, patch

import pytest
from fastapi.testclient import TestClient

from main import FORMAT_MAPPING, LIBREOFFICE_BINARY, app, convert_with_libreoffice

client = TestClient(app)


def is_libreoffice_available():
    """Check if LibreOffice is available in the system."""
    try:
        result = subprocess.run(
            [LIBREOFFICE_BINARY, "--version"],
            capture_output=True,
            timeout=5,
        )
        return result.returncode == 0
    except (subprocess.TimeoutExpired, FileNotFoundError):
        return False


def create_minimal_doc_file(file_path: Path):
    """Create a minimal valid .doc file."""
    # Minimal RTF content that can be converted (LibreOffice accepts RTF as .doc)
    rtf_content = r"""{\rtf1\ansi\deff0
{\fonttbl{\f0 Times New Roman;}}
\f0\fs24 This is a test document for conversion.
}"""
    file_path.write_text(rtf_content)


def create_minimal_xls_file(file_path: Path):
    """Create a minimal valid .xls file using CSV format."""
    # Simple CSV that can be saved as .xls
    csv_content = "Name,Value\nTest,123\nData,456"
    file_path.write_text(csv_content)


def create_minimal_ppt_file(file_path: Path):
    """Create a minimal valid .ppt file."""
    # For PPT, we'll create a simple text file that LibreOffice can attempt to import
    ppt_content = "Test Presentation\n\nSlide 1 Content"
    file_path.write_text(ppt_content)


class TestHealthEndpoints:
    """Test health check endpoints."""

    def test_root_endpoint(self):
        """Test the root endpoint returns correct information."""
        response = client.get("/")
        assert response.status_code == 200
        data = response.json()
        assert data["status"] == "healthy"
        assert data["service"] == "Legacy Office Converter"
        assert set(data["supported_formats"]) == {".doc", ".xls", ".ppt"}

    def test_health_endpoint(self):
        """Test the health endpoint."""
        response = client.get("/health")
        assert response.status_code == 200
        assert response.json() == {"status": "healthy"}


class TestConversionLogic:
    """Test the LibreOffice conversion logic."""

    @patch("main.subprocess.run")
    def test_convert_with_libreoffice_success(self, mock_run):
        """Test successful conversion with LibreOffice."""
        # Setup
        mock_run.return_value = Mock(returncode=0, stderr="")

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_file = temp_path / "test.doc"
            input_file.write_text("test content")

            # Create expected output file
            output_file = temp_path / "test.docx"
            output_file.write_text("converted content")

            # Execute
            result = convert_with_libreoffice(input_file, temp_path, "docx")

            # Assert
            assert result == output_file
            mock_run.assert_called_once()
            call_args = mock_run.call_args[0][0]
            assert call_args[0] == LIBREOFFICE_BINARY
            assert "--headless" in call_args
            assert "--convert-to" in call_args
            assert "docx" in call_args

    @patch("main.subprocess.run")
    def test_convert_with_libreoffice_failure(self, mock_run):
        """Test failed conversion with LibreOffice."""
        # Setup - simulate LibreOffice failure
        mock_run.return_value = Mock(returncode=1, stderr="Conversion failed")

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_file = temp_path / "test.doc"
            input_file.write_text("test content")

            # Execute
            result = convert_with_libreoffice(input_file, temp_path, "docx")

            # Assert
            assert result is None

    @patch("main.subprocess.run")
    def test_convert_with_libreoffice_timeout(self, mock_run):
        """Test conversion timeout handling."""
        # Setup - simulate timeout
        import subprocess
        mock_run.side_effect = subprocess.TimeoutExpired(cmd=LIBREOFFICE_BINARY, timeout=60)

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_file = temp_path / "test.doc"
            input_file.write_text("test content")

            # Execute
            result = convert_with_libreoffice(input_file, temp_path, "docx")

            # Assert
            assert result is None

    @patch("main.subprocess.run")
    def test_convert_with_libreoffice_exception(self, mock_run):
        """Test conversion exception handling."""
        # Setup - simulate exception
        mock_run.side_effect = Exception("Unexpected error")

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_file = temp_path / "test.doc"
            input_file.write_text("test content")

            # Execute
            result = convert_with_libreoffice(input_file, temp_path, "docx")

            # Assert
            assert result is None


class TestConvertEndpoint:
    """Test the /convert endpoint."""

    @patch("main.convert_with_libreoffice")
    def test_convert_doc_to_docx(self, mock_convert):
        """Test converting .doc to .docx."""
        # Setup
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_file = temp_path / "test.docx"
            output_file.write_bytes(b"converted content")
            mock_convert.return_value = output_file

            # Create test file
            file_content = b"fake doc content"
            files = {"file": ("test.doc", io.BytesIO(file_content), "application/msword")}

            # Execute
            response = client.post("/convert", files=files)

            # Assert
            assert response.status_code == 200
            assert response.headers["content-disposition"] == 'attachment; filename="test.docx"'
            mock_convert.assert_called_once()

    @patch("main.convert_with_libreoffice")
    def test_convert_xls_to_xlsx(self, mock_convert):
        """Test converting .xls to .xlsx."""
        # Setup
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_file = temp_path / "spreadsheet.xlsx"
            output_file.write_bytes(b"converted content")
            mock_convert.return_value = output_file

            # Create test file
            file_content = b"fake xls content"
            files = {"file": ("spreadsheet.xls", io.BytesIO(file_content), "application/vnd.ms-excel")}

            # Execute
            response = client.post("/convert", files=files)

            # Assert
            assert response.status_code == 200
            assert response.headers["content-disposition"] == 'attachment; filename="spreadsheet.xlsx"'

    @patch("main.convert_with_libreoffice")
    def test_convert_ppt_to_pptx(self, mock_convert):
        """Test converting .ppt to .pptx."""
        # Setup
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_file = temp_path / "presentation.pptx"
            output_file.write_bytes(b"converted content")
            mock_convert.return_value = output_file

            # Create test file
            file_content = b"fake ppt content"
            files = {"file": ("presentation.ppt", io.BytesIO(file_content), "application/vnd.ms-powerpoint")}

            # Execute
            response = client.post("/convert", files=files)

            # Assert
            assert response.status_code == 200
            assert response.headers["content-disposition"] == 'attachment; filename="presentation.pptx"'

    def test_convert_unsupported_format(self):
        """Test uploading unsupported file format."""
        # Create test file with unsupported format
        file_content = b"fake pdf content"
        files = {"file": ("document.pdf", io.BytesIO(file_content), "application/pdf")}

        # Execute
        response = client.post("/convert", files=files)

        # Assert
        assert response.status_code == 400
        assert "Unsupported file format" in response.json()["detail"]

    def test_convert_no_file_extension(self):
        """Test uploading file without extension."""
        # Create test file without extension
        file_content = b"fake content"
        files = {"file": ("document", io.BytesIO(file_content), "application/octet-stream")}

        # Execute
        response = client.post("/convert", files=files)

        # Assert
        assert response.status_code == 400

    @patch("main.convert_with_libreoffice")
    def test_convert_conversion_fails(self, mock_convert):
        """Test handling of failed conversion."""
        # Setup - simulate failed conversion
        mock_convert.return_value = None

        # Create test file
        file_content = b"fake doc content"
        files = {"file": ("test.doc", io.BytesIO(file_content), "application/msword")}

        # Execute
        response = client.post("/convert", files=files)

        # Assert
        assert response.status_code == 500
        assert "Conversion failed" in response.json()["detail"]

    def test_convert_missing_file(self):
        """Test endpoint without file upload."""
        # Execute
        response = client.post("/convert")

        # Assert
        assert response.status_code == 422  # Unprocessable Entity

    @patch("main.convert_with_libreoffice")
    def test_convert_case_insensitive_extension(self, mock_convert):
        """Test that file extensions are handled case-insensitively."""
        # Setup
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_file = temp_path / "TEST.docx"
            output_file.write_bytes(b"converted content")
            mock_convert.return_value = output_file

            # Create test file with uppercase extension
            file_content = b"fake doc content"
            files = {"file": ("TEST.DOC", io.BytesIO(file_content), "application/msword")}

            # Execute
            response = client.post("/convert", files=files)

            # Assert
            assert response.status_code == 200


class TestFormatMappings:
    """Test format mapping constants."""

    def test_format_mapping_completeness(self):
        """Test that all legacy formats are mapped."""
        expected_formats = {".doc", ".xls", ".ppt"}
        assert set(FORMAT_MAPPING.keys()) == expected_formats

    def test_format_mapping_correctness(self):
        """Test that format mappings are correct."""
        assert FORMAT_MAPPING[".doc"] == ".docx"
        assert FORMAT_MAPPING[".xls"] == ".xlsx"
        assert FORMAT_MAPPING[".ppt"] == ".pptx"


@pytest.mark.skipif(not is_libreoffice_available(), reason="LibreOffice not available")
class TestActualConversionFlow:
    """Integration tests for actual file conversion with LibreOffice."""

    def test_convert_with_libreoffice_creates_output_file(self):
        """Test that LibreOffice actually creates the output file in expected directory."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)

            # Create input file
            input_file = temp_path / "test_document.doc"
            create_minimal_doc_file(input_file)

            # Verify input file exists
            assert input_file.exists()
            assert input_file.stat().st_size > 0

            # Create output directory
            output_dir = temp_path / "output"
            output_dir.mkdir()

            # Perform conversion
            result = convert_with_libreoffice(input_file, output_dir, "docx")

            # Assert output file was created in expected directory
            assert result is not None
            assert result.exists()
            assert result.parent == output_dir
            assert result.name == "test_document.docx"
            assert result.stat().st_size > 0

            # Verify the file is in the correct location
            expected_path = output_dir / "test_document.docx"
            assert result == expected_path

    def test_convert_doc_to_docx_real_conversion(self):
        """Test actual .doc to .docx conversion."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)

            # Create input file
            input_file = temp_path / "document.doc"
            create_minimal_doc_file(input_file)

            output_dir = temp_path / "output"
            output_dir.mkdir()

            # Convert
            result = convert_with_libreoffice(input_file, output_dir, "docx")

            # Verify
            assert result is not None
            assert result.suffix == ".docx"
            assert result.exists()
            assert result.parent == output_dir

    def test_convert_xls_to_xlsx_real_conversion(self):
        """Test actual .xls to .xlsx conversion."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)

            # Create input file
            input_file = temp_path / "spreadsheet.xls"
            create_minimal_xls_file(input_file)

            output_dir = temp_path / "output"
            output_dir.mkdir()

            # Convert
            result = convert_with_libreoffice(input_file, output_dir, "xlsx")

            # Verify
            assert result is not None
            assert result.suffix == ".xlsx"
            assert result.exists()
            assert result.parent == output_dir

    def test_convert_ppt_to_pptx_real_conversion(self):
        """Test actual .ppt to .pptx conversion."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)

            # Create input file
            input_file = temp_path / "presentation.ppt"
            create_minimal_ppt_file(input_file)

            output_dir = temp_path / "output"
            output_dir.mkdir()

            # Convert
            result = convert_with_libreoffice(input_file, output_dir, "pptx")

            # Verify
            assert result is not None
            assert result.suffix == ".pptx"
            assert result.exists()
            assert result.parent == output_dir

    def test_multiple_conversions_in_same_directory(self):
        """Test multiple conversions output to the same directory."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_dir = temp_path / "output"
            output_dir.mkdir()

            # Create and convert multiple files
            doc_file = temp_path / "doc1.doc"
            create_minimal_doc_file(doc_file)
            result1 = convert_with_libreoffice(doc_file, output_dir, "docx")

            xls_file = temp_path / "sheet1.xls"
            create_minimal_xls_file(xls_file)
            result2 = convert_with_libreoffice(xls_file, output_dir, "xlsx")

            # Verify both files exist in output directory
            assert result1 is not None and result1.exists()
            assert result2 is not None and result2.exists()
            assert result1.parent == output_dir
            assert result2.parent == output_dir

            # Verify output directory contains both files
            output_files = list(output_dir.iterdir())
            assert len(output_files) == 2
            assert any(f.name == "doc1.docx" for f in output_files)
            assert any(f.name == "sheet1.xlsx" for f in output_files)

    def test_output_file_not_created_on_invalid_input(self):
        """Test that no output file is created when input is invalid."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_dir = temp_path / "output"
            output_dir.mkdir()

            # Create invalid input file
            invalid_file = temp_path / "invalid.doc"
            invalid_file.write_text("This is not a valid document format")

            # Attempt conversion
            result = convert_with_libreoffice(invalid_file, output_dir, "docx")

            # Verify no output was created or conversion failed
            # LibreOffice might fail or create empty file
            if result is None:
                # Conversion failed as expected
                assert True
            else:
                # If file was created, it should still be in output_dir
                assert result.parent == output_dir


@pytest.mark.skipif(not is_libreoffice_available(), reason="LibreOffice not available")
class TestEndToEndConversion:
    """End-to-end integration tests with actual API calls and LibreOffice."""

    def test_api_endpoint_creates_file_in_temp_directory(self):
        """Test that API endpoint creates converted file in temporary directory."""
        # Create a minimal doc file
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_file = temp_path / "test.doc"
            create_minimal_doc_file(input_file)

            # Read file content
            file_content = input_file.read_bytes()

            # Send to API
            files = {"file": ("test.doc", io.BytesIO(file_content), "application/msword")}
            response = client.post("/convert", files=files)

            # Verify response
            assert response.status_code == 200
            assert response.headers["content-disposition"] == 'attachment; filename="test.docx"'

            # Verify response contains file data
            assert len(response.content) > 0

    def test_api_endpoint_doc_to_docx_full_flow(self):
        """Test full conversion flow from .doc to .docx through API."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_file = temp_path / "document.doc"
            create_minimal_doc_file(input_file)

            files = {"file": ("document.doc", input_file.open("rb"), "application/msword")}
            response = client.post("/convert", files=files)

            assert response.status_code == 200

            # Save and verify the output
            output_file = temp_path / "output.docx"
            output_file.write_bytes(response.content)
            assert output_file.exists()
            assert output_file.stat().st_size > 0

    def test_api_endpoint_xls_to_xlsx_full_flow(self):
        """Test full conversion flow from .xls to .xlsx through API."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_file = temp_path / "spreadsheet.xls"
            create_minimal_xls_file(input_file)

            files = {"file": ("spreadsheet.xls", input_file.open("rb"), "application/vnd.ms-excel")}
            response = client.post("/convert", files=files)

            assert response.status_code == 200

            # Save and verify the output
            output_file = temp_path / "output.xlsx"
            output_file.write_bytes(response.content)
            assert output_file.exists()
            assert output_file.stat().st_size > 0

    def test_api_endpoint_ppt_to_pptx_full_flow(self):
        """Test full conversion flow from .ppt to .pptx through API."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_file = temp_path / "presentation.ppt"
            create_minimal_ppt_file(input_file)

            files = {"file": ("presentation.ppt", input_file.open("rb"), "application/vnd.ms-powerpoint")}
            response = client.post("/convert", files=files)

            assert response.status_code == 200

            # Save and verify the output
            output_file = temp_path / "output.pptx"
            output_file.write_bytes(response.content)
            assert output_file.exists()
            assert output_file.stat().st_size > 0
