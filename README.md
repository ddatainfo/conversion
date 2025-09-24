# Industrial Automation Conversion API

A FastAPI application for processing and merging industrial automation data from PDF and text files.

## Overview

This application provides a REST API for:
- Processing PDF files containing industrial measurement data
- Extracting measurements from text files
- Merging and converting data between different formats
- Generating reports for industrial automation systems

## Project Structure

```
conversion/
├── api/
│   ├── main.py                 # FastAPI application entry point
│   ├── routes/
│   │   └── file_routes.py      # File processing endpoints
│   ├── services/
│   │   └── merge_service.py    # Data merging logic
│   └── utils/
│       ├── excel_extraction.py # Excel data extraction utilities
│       ├── extract_measurements.py # Measurement extraction utilities
│       └── merge_data.py       # Data merging utilities
├── PDF/                        # Source PDF files
├── TXT/                        # Source text files
├── REPORT/                     # Generated reports
└── requirements.txt            # Python dependencies
```

## Prerequisites

- Python 3.8 or higher
- pip (Python package installer)

## Installation

1. Clone or download the project to your local machine

2. Navigate to the project directory:
   ```bash
   cd conversion
   ```

3. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   ```

4. Activate the virtual environment:
   - On Windows:
     ```bash
     venv\Scripts\activate
     ```
   - On macOS/Linux:
     ```bash
     source venv/bin/activate
     ```

5. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

### Using Uvicorn (Recommended)

To run the application using uvicorn, use one of the following commands:

#### Basic Usage
```bash
uvicorn api.main:app --reload
```

#### With Custom Host and Port
```bash
uvicorn api.main:app --host 0.0.0.0 --port 8000 --reload
```

#### Production Mode (without auto-reload)
```bash
uvicorn api.main:app --host 0.0.0.0 --port 8000
```

### Command Line Options

- `--reload`: Enable auto-reload for development (restarts server on code changes)
- `--host`: Bind socket to this host (default: 127.0.0.1)
- `--port`: Bind socket to this port (default: 8000)
- `--workers`: Number of worker processes (for production)

### Development Mode
For development with auto-reload:
```bash
uvicorn api.main:app --reload --host 127.0.0.1 --port 8000
```

### Production Mode
For production deployment:
```bash
uvicorn api.main:app --host 0.0.0.0 --port 8000 --workers 4
```

## Accessing the Application

Once the server is running, you can access:

- **API Root**: http://localhost:8000/
- **Interactive API Documentation (Swagger UI)**: http://localhost:8000/docs
- **Alternative API Documentation (ReDoc)**: http://localhost:8000/redoc

## API Endpoints

- `GET /`: Welcome message and API status
- `POST /files/*`: File processing endpoints (see `/docs` for detailed documentation)

## Usage Examples

### Basic Health Check
```bash
curl http://localhost:8000/
```

### Using the Interactive Documentation
1. Start the server with uvicorn
2. Open your browser and navigate to http://localhost:8000/docs
3. Explore and test the available endpoints directly from the browser

## Development

### Running in Development Mode
```bash
uvicorn api.main:app --reload --log-level debug
```

### Environment Variables
You can configure the application using environment variables:
- `HOST`: Server host (default: 127.0.0.1)
- `PORT`: Server port (default: 8000)

Example:
```bash
export HOST=0.0.0.0
export PORT=8080
uvicorn api.main:app --host $HOST --port $PORT --reload
```

## Troubleshooting

### Common Issues

1. **Port already in use**: If port 8000 is busy, use a different port:
   ```bash
   uvicorn api.main:app --port 8001 --reload
   ```

2. **Module not found errors**: Ensure you're running from the project root directory and have activated your virtual environment.

3. **Permission errors**: On some systems, you might need to use `sudo` or run as administrator for certain ports.

## Dependencies

The application uses the following main dependencies:
- **FastAPI**: Modern web framework for building APIs
- **Uvicorn**: ASGI server for running FastAPI applications
- **Pandas**: Data manipulation and analysis
- **NumPy**: Numerical computing
- **OpenPyXL**: Excel file processing
- **python-multipart**: File upload support

## Contributing

1. Ensure all dependencies are installed
2. Run the application in development mode with `--reload`
3. Test your changes using the interactive documentation at `/docs`
4. Follow Python PEP 8 style guidelines
