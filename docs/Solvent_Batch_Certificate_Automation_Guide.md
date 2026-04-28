# Solvent Batch Certificate Automation Guide

## Overview

This VBA macro automates the process of saving and printing solvent or chemical batch certificate sheets from Excel.

It is designed to support quality control, laboratory, compliance, and chemical documentation workflows where certificates need to be saved consistently as PDF files.

## Workflow

1. User completes the batch certificate sheet.
2. User clicks the Save as PDF button.
3. A pre-save checklist appears.
4. The macro validates the sheet.
5. The macro checks for red warning cells.
6. The macro checks for empty green required cells.
7. The user selects or uses a saved folder location.
8. The macro generates a PDF file name using product name, certificate date, and batch number.
9. The certificate is saved as a PDF.
10. The user can print the sheet using a one-click print button.

## Validation Logic

The macro prevents saving or printing if:

- Any red warning cell is found
- Any green required cell is empty
- The user does not confirm the checklist
- No folder is selected

## File Naming Logic

The macro generates PDF names using this structure:

ProductName_YYYYMMDD_BatchNumber.pdf

Example:

SampleSolventA_20260309_BATCH4874.pdf

## Buttons Included

The macro can create three Excel buttons:

- Save as PDF with checklist
- Print Sheet with checklist
- Change Folder and Save PDF

## Suitable Use Cases

- Chemical batch certificates
- Solvent documentation
- Laboratory quality reports
- Certificate of Analysis templates
- Compliance-controlled Excel workflows
- Repetitive PDF saving and printing tasks

## Data Privacy

This project uses generic VBA logic and sample screenshots only. No confidential company data, customer data, supplier data, batch records, or internal documents are included.
