# Legacy Office Converter

A lightweight FastAPI microservice that converts legacy Microsoft Office files to modern formats using LibreOffice in headless mode.

## Description

This service provides a REST API for converting older Office document formats to their modern equivalents:

- `.doc` → `.docx` (Word documents)
- `.xls` → `.xlsx` (Excel spreadsheets)
- `.ppt` → `.pptx` (PowerPoint presentations)

The service leverages LibreOffice's powerful conversion engine running in headless mode, making it ideal for server-side batch processing, automated workflows, and containerized deployments.

## Features

- **Simple REST API** - Upload a file and receive the converted version
- **Health checks** - Built-in endpoints for monitoring and orchestration
- **Docker ready** - Containerized deployment with all dependencies
- **Temporary file handling** - Secure processing with automatic cleanup
- **Timeout protection** - 60-second conversion timeout to prevent hanging

## Requirements

- Python 3.11+
- LibreOffice (automatically installed in Docker)
- FastAPI dependencies (see `requirements.txt`)

## Installation

### Using Docker (Recommended)

```bash
docker build -t legacy-office-converter .
docker run -p 9800:9800 legacy-office-converter
```

### Local Development

```bash
# Install LibreOffice
sudo apt-get install libreoffice libreoffice-writer libreoffice-calc libreoffice-impress

# Install Python dependencies
pip install -r requirements.txt

# Run the service
uvicorn main:app --host 0.0.0.0 --port 9800
```

## API Endpoints

### `GET /`

Root endpoint with service status and supported formats.

**Response:**
```json
{
  "status": "healthy",
  "service": "Legacy Office Converter",
  "supported_formats": [".doc", ".xls", ".ppt"]
}
```

### `POST /convert`

Convert a legacy Office file to modern format.

**Request:**
- Method: `POST`
- Content-Type: `multipart/form-data`
- Body: File upload with key `file`

**Response:**
- Returns the converted file as a download
- Filename preserves original name with updated extension

**Example using curl:**
```bash
curl -X POST http://localhost:9800/convert \
  -F "file=@/path/to/document.doc" \
  -o converted_document.docx
```

**Example using Python:**
```python
import requests

with open("document.doc", "rb") as f:
    response = requests.post(
        "http://localhost:9800/convert",
        files={"file": f}
    )

with open("document.docx", "wb") as f:
    f.write(response.content)
```

**Example using JavaScript/fetch:**
```javascript
const formData = new FormData();
formData.append('file', fileInput.files[0]);

const response = await fetch('http://localhost:9800/convert', {
  method: 'POST',
  body: formData
});

const blob = await response.blob();
const url = window.URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'converted_document.docx';
a.click();
```

### `GET /health`

Health check endpoint for monitoring and container orchestration.

**Response (Healthy):**
```json
{
  "status": "healthy"
}
```

**Response (Unhealthy):**
- Status Code: `503`
```json
{
  "detail": {
    "status": "unhealthy",
    "reason": "LibreOffice binary not responding"
  }
}
```

## Error Responses

### 400 Bad Request
Unsupported file format provided.

```json
{
  "detail": "Unsupported file format. Supported formats: .doc, .xls, .ppt"
}
```

### 500 Internal Server Error
Conversion failed or internal error occurred.

```json
{
  "detail": "Conversion failed. Please check if the file is valid."
}
```

## Configuration

### Environment Variables

- `LIBREOFFICE_BINARY` - Path to LibreOffice binary (default: `soffice`)

```bash
export LIBREOFFICE_BINARY=/usr/bin/soffice
uvicorn main:app --host 0.0.0.0 --port 9800
```

## Use Cases

- **Legacy system migration** - Bulk convert old Office files to modern formats
- **Document processing pipelines** - Integrate into automated workflows
- **Cloud storage integration** - Convert files on upload
- **Archive modernization** - Update document archives to current standards
- **Cross-platform compatibility** - Ensure files work with modern Office suites

## Testing

```bash
# Run tests
pytest

# Run tests with coverage
pytest --cov=main --cov-report=html
```

## License

This project is open source and available under the MIT License.
