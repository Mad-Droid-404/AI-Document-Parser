# AI Document Parser

An AI outlook plugin document that parses, analyses and creates summary based on the contents of the mail.

## Overview

This project consists of two main components:
- **Frontend**: Office Add-in with HTTPS server for the task pane interface
- **Backend**: Python API server for document processing and AI analysis

## Prerequisites

Before starting, ensure you have the following installed:
- [Node.js](https://nodejs.org/) (version 14 or higher)
- [Python](https://python.org/) (version 3.8 or higher)
- pip (Python package manager)

## Quick Start

### 1. Initial Setup

Clone the repository and install dependencies:

```bash
# Install Node.js dependencies
npm install express office-addin-dev-certs office-addin-manifest http-server

# Generate SSL certificates (REQUIRED for Office Add-ins)
npx office-addin-dev-certs install --machine
```

### 2. Backend Setup (Python)

```bash
# Create and activate virtual environment (recommended)
python -m venv venv

# Activate virtual environment
source venv/bin/activate  # Linux/macOS
# OR
venv\Scripts\activate     # Windows

# Install Python dependencies (add your requirements.txt)
pip install -r requirements.txt

# Start Python backend server
python app.py
```

The backend will be available at `http://localhost:5000`

### 3. Frontend Setup (Office Add-in)

```bash
# Start the HTTPS development server
node server.js

# Alternative method using http-server:
# npx http-server . -p 3000 --ssl -c-1
```

The frontend will be available at `https://localhost:3000`

## Development & Testing

### Validate Manifest

Before sideloading the add-in, validate your manifest file:

```bash
npx office-addin-manifest validate manifest.xml
```

### Health Checks

Verify both servers are running correctly:

```bash
# Test frontend accessibility
curl -k https://localhost:3000/taskpane.html

# Test backend health endpoint
curl -X POST http://localhost:5000/api/health
```

## Project Structure

```
ai-document-parser/
├── manifest.xml          # Office Add-in manifest
├── taskpane.html         # Main add-in interface
├── server.js             # HTTPS development server
├── app.py                # Python backend API
├── requirements.txt      # Python dependencies
├── package.json          # Node.js dependencies
└── README.md            # This file
```

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/health` | POST | Health check for backend |
| `/api/parse` | POST | Parse document content |
| `/api/analyze` | POST | AI analysis of document |