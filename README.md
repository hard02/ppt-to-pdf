# ppt-to-pdf

This Python script allows you to convert PowerPoint (PPT/PPTX) files to PDF format using LibreOffice. The script provides a graphical user interface (GUI) for file selection and outputs the converted PDF files in the specified output folder.

## Prerequisites

- Python 3.x
- LibreOffice installed and accessible from the command line

## Installation

1. Clone the repository or download the script.
2. Install the required Python packages:
   ```
   pip install tkinter
   ```
**Ensure LibreOffice is installed and available in the system's PATH.**

## Usage

1. Open a terminal or command prompt.
2. Navigate to the directory where the script is located.
3. Run the script:
   ```
   python ppt_to_pdf_converter.py
   ```
4. The GUI file selection dialog will appear. Select one or more PowerPoint files.
5. Choose the desired output folder where the converted PDF files will be saved.
6. Click the "Convert" button to start the conversion process.
7. The script will convert each selected PowerPoint file to PDF and save it in the specified output folder.
8. Once the conversion is complete, the script will display a success message.

>Note: The converted PDF files will be saved with the same name as the original PowerPoint files in the **output** folder.
