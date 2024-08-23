# Data Parsing

This Python script is designed to extract specific data from  PDFs and save the extracted information into an Excel file. It is particularly useful for extracting and organizing legal case information such as dates, case numbers, attorney names, and other relevant details from court documents.

## Features

- Extracts text from PDF files.
- Parses and organizes data such as case numbers, dates, attorney names, and other court details.
- Saves the extracted data into an Excel file for further analysis or record-keeping.

## Requirements

- Python 3.x
- `pdfplumber` library for extracting text from PDF files.
- `pandas` library for organizing and exporting data to an Excel file.
- `re` (Regular Expressions) for pattern matching and data extraction.

## Installation

To use this script, you'll need to have Python installed on your machine along with the required libraries. You can install the required libraries using pip:

```bash
pip install pdfplumber pandas
```

Usage
Specify the PDF File Path: Update the pdf_path variable in the script to point to the PDF file you want to process.

Specify the Output Path: Update the output_path variable to where you want to save the resulting Excel file.

Run the Script:
- You can run the script using the following command:
```bash
python script_name.py
Make sure to replace script_name.py with the actual name of your Python script.
```
## Example
If you have a court PDF file located at documents/court_case.pdf and you want to save the extracted data to output/court_data.xlsx, you would set:
```bash
pdf_path = 'documents/court_case.pdf'
output_path = 'output/court_data.xlsx'

Then run the script, and the data will be extracted and saved into the specified Excel file.
```
## Functions
- extract_text_from_pdf(pdf_path)
- Extracts all text from the specified PDF file.


## Error Handling
- The script includes basic error handling to manage issues such as:
- PDF file not being accessible or not containing any extractable text.
- Problems saving the extracted data to an Excel file.

## Notes
- Ensure that the PDF file follows a consistent structure so that the regular expressions can correctly identify and extract the necessary data.
- Modify the regular expressions within the parse_data function if the structure of your PDF differs.
