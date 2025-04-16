# Universal-File-Converter
Universal File Converter is a robust Python utility that seamlessly converts and merges files across diverse formats—including DOCX, TXT, CSV, XLSX, PPTX, and various image types—to and from PDFs, featuring comprehensive error handling and a user-friendly CLI for streamlined document workflows.

## Overview

MultiFormat Converter Suite simplifies file management by providing a single CLI tool for:
- **File Conversion:** Convert various document, spreadsheet, presentation, and image formats to PDF.
- **Reverse Conversion:** Convert PDFs back into DOCX, images, XLSX, and PPTX.
- **PDF Merging:** Merge multiple PDF files into one consolidated document.

This tool leverages powerful Python libraries to ensure reliable file handling and streamlined workflows, making it an essential utility for automating routine document processes.

## Key Features

- **Multi-Format Conversion:**
  - Convert DOCX, TXT, CSV, XLSX, PPTX, and various image formats (JPEG, PNG, WEBP, etc.) to PDF.
  - Reverse conversion from PDF to DOCX, images, XLSX, and PPTX.
  
- **PDF Merging:**
  - Merge multiple PDF files into a single document using PyPDF2.
  
- **User-Friendly CLI:**
  - Interactive prompts guide users through file selection, format options, and output naming.
  - Robust input validation and file path verification to minimize errors.
  
- **Automation Ready:**
  - Easily integrate into batch processing workflows and automated systems.

## Technologies Used

- **Libraries:** `docx2pdf`, `fpdf`, `pdf2docx`, `pptx`, `pdf2image`, `PIL`, `PyPDF2`, `pandas`, `img2pdf`, `pdfkit`, `tabula`
- **Standard Modules:** `os`

