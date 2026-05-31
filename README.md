# ppt-to-pdf-automation

# PowerPoint to PDF Automation

## Overview

This project automates the conversion of PowerPoint presentations (.pptx) into PDF documents. The solution scans an input directory, processes all PowerPoint files, converts them into PDFs, and saves the output automatically.

It is designed to eliminate repetitive manual export tasks and streamline reporting and document publishing workflows.

## Features

- Automatic PowerPoint-to-PDF conversion
- Batch processing of multiple presentations
- Automatic output folder creation
- Preserves original slide formatting
- Bulk document publishing
- Simple folder-based workflow

## Technologies Used

- Python
- comtypes
- Microsoft PowerPoint COM Automation

## Workflow

1. Scan input directory for PowerPoint files.
2. Launch Microsoft PowerPoint programmatically.
3. Open each presentation.
4. Convert presentation to PDF.
5. Save generated PDF to output folder.
6. Close presentation.
7. Repeat for all files automatically.

## Project Structure

```text
PPT-to-PDF-Automation/
│
├── input/
│   ├── presentation1.pptx
│   ├── presentation2.pptx
│
├── output/
│   ├── presentation1.pdf
│   ├── presentation2.pdf
│
└── main.py
```

## Installation

```bash
pip install comtypes
```

### Requirements

- Windows Operating System
- Microsoft PowerPoint Installed
- Python 3.x

## Run

```bash
python main.py
```

## Output

The generated PDF files are automatically stored inside the `output` folder while preserving the original presentation layout.

## Use Cases

- Financial Reporting
- Business Presentations
- Executive Reports
- Client Deliverables
- Document Publishing
- Presentation Archiving
- Automated Reporting Workflows

## Benefits

- Eliminates manual export effort
- Reduces processing time
- Supports bulk conversions
- Maintains presentation quality
- Improves reporting efficiency
