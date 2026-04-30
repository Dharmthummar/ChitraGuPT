# ChitraGuPT | ExcelVision App

**An AI-powered document extraction agent that automates data entry from receipts, invoices, screenshots, and other documents directly into local Excel spreadsheets.**

[![AI Powered](https://img.shields.io/badge/AI-Powered-blueviolet?style=for-the-badge)](#)
[![Built with Flask](https://img.shields.io/badge/Backend-Flask-black?style=for-the-badge&logo=flask)](https://flask.palletsprojects.com/)
[![Creator](https://img.shields.io/badge/Creator-Dharm%20Thummar-blue?style=for-the-badge)](https://github.com/dharmthummar)

---

## Overview

**ChitraGuPT** transforms manual data entry workflows by extracting structured information from documents and automatically appending it to your existing Excel workbooks.

Whether you are working with receipts, invoices, screenshots, or scanned documents, ChitraGuPT intelligently reads the content, maps the extracted data to your spreadsheet columns, and inserts a new row without disturbing your existing workbook formatting.

The application is designed for local-first productivity, combining AI-powered extraction with a simple desktop workflow and optional mobile capture support.

---

## Key Features

### AI-Powered Data Extraction

ChitraGuPT uses an AI extraction engine to parse unstructured documents and convert them into structured, spreadsheet-ready data.

### Excel Workbook Integration

The app detects worksheet headers, maps extracted values to the correct columns, and appends new rows directly to your existing `.xlsx` or `.xlsm` files.

### Multi-Sheet Workflow Support

Work with different worksheets inside the same workbook. ChitraGuPT can inspect sheet headers and append data to the selected worksheet.

### Mobile-First Capture

Capture receipts or physical documents from your phone and send them to your desktop app over the same local Wi-Fi or LAN network.

### Local-First Privacy

Your Excel workbooks remain on your local machine. Only the uploaded document and relevant sheet headers are sent to the AI service for extraction.

---

## Getting Started

### Prerequisites

Before using ChitraGuPT, make sure you have the following:

- Python 3.8 or later installed on your Windows machine
- A valid AI service API key
- Existing Excel workbooks in `.xlsx` or `.xlsm` format
- Old `.xls` files converted to a supported modern Excel format

---

## Installation

### First-Time Setup

1. Clone this repository or download the ZIP release.

2. Run the setup file:

   ```bash
   SETUP.bat
   ```

   This creates a virtual environment and installs the required dependencies.

3. Start the application:

   ```bash
   START_APP.bat
   ```

4. The local web interface will open in your browser.

5. Click **Settings / API** in the top-right corner of the app.

6. Paste your API key.

7. Save the settings.

---

## Usage

### Desktop Workflow

1. Double-click:

   ```bash
   START_APP.bat
   ```

2. Paste the file path of your Excel workbook, or select a workbook from the **Recent** tab.

3. Click **Inspect Sheet** to load the worksheet columns.

4. Upload a PDF, invoice image, receipt, or screenshot into the upload area.

5. Click **Extract and Add Row**.

6. ChitraGuPT will process the document and append a highlighted row to the selected worksheet.

---

### Phone Upload Workflow

Use this workflow when you want to capture a physical receipt or document from your phone.

1. On your main computer, run:

   ```bash
   START_HOST_SHARE.bat
   ```

2. Click the **Host** button in the app header.

3. ChitraGuPT will generate a local sharing link.

4. Open the link on your phone.

5. Make sure your phone and computer are connected to the same Wi-Fi or LAN network.

6. Take a photo using your phone camera.

7. The image will stream to the desktop app and become available for extraction.

---

## Important Notes

### File Locking

Microsoft Excel may lock a workbook while it is open. If the workbook is open in Excel, ChitraGuPT may not be able to append new rows.

Before extracting data, close the workbook in Excel.

### Data Retention

Uploaded documents are not stored or retained by the local application after extraction.

### Extraction History

ChitraGuPT keeps a local extraction history for your reference in:

```text
data/history.jsonl
```

---

## Supported File Types

ChitraGuPT is designed to work with:

- Excel workbooks: `.xlsx`, `.xlsm`
- Document uploads: PDFs
- Image uploads: receipts, invoices, screenshots, scanned documents

---

## Project Structure

```text
ChitraGuPT/
├── data/
│   └── history.jsonl
├── SETUP.bat
├── START_APP.bat
├── START_HOST_SHARE.bat
├── requirements.txt
└── README.md
```

---

## Privacy

ChitraGuPT is built with a local-first approach.

Your Excel files remain on your computer. The application only sends the uploaded document and selected worksheet headers to the AI service for extraction. No workbook files are uploaded.

---

## Built By

**Dharm Thummar**

- [GitHub](https://github.com/dharmthummar)
- [LinkedIn](https://linkedin.com/in/dharmthummar)
- [Twitter](https://twitter.com/dharmthummar)
- [Instagram](https://instagram.com/dharm_1602)

---

## Summary

ChitraGuPT helps reduce repetitive spreadsheet work by turning documents into structured Excel rows with the help of AI. It is ideal for invoice tracking, receipt logging, financial records, business documentation, and other data-entry-heavy workflows.
