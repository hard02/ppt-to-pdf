import os
import subprocess
from tkinter import Tk, font
from tkinter.filedialog import askopenfilenames

def ppt_to_pdf(input_files, output_folder):
    for file in input_files:
        ppt_path = file
        pdf_path = os.path.join(output_folder, os.path.splitext(os.path.basename(file))[0] + '.pdf')

        # Use libreoffice to convert PowerPoint file to PDF
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_folder, ppt_path])

        print(f"Converted '{ppt_path}' to '{pdf_path}'")

# Get the current working directory
current_dir = os.getcwd()

# Initialize Tkinter
root = Tk()
root.withdraw()

# Set the desired font size for the file selection dialog
text_size = 12  # Modify this value as needed

# Configure the file selection dialog with the custom font
file_options = {
    'title': "Select PowerPoint Files",
    'filetypes': (("PowerPoint Files", "*.ppt *.pptx"),),
    'initialdir': current_dir  # Set the initial directory to the current working directory
}

# Prompt the user to select input files using the configured options
input_files = askopenfilenames(**file_options)

# Create output folder in the current directory
output_folder = os.path.join(current_dir, 'output')
os.makedirs(output_folder, exist_ok=True)

# Convert the ppt files to pdf
ppt_to_pdf(input_files, output_folder)

