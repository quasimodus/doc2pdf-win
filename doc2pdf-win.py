import os
import subprocess
from pathlib import Path


def find_soffice_path():
    """Find the path to soffice.exe on Windows 10."""
    possible_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    return None


def convert_doc_to_pdf(input_path, output_path):
    """Convert a .doc or .docx file to PDF using LibreOffice on Windows 10."""
    soffice_path = find_soffice_path()
    if not soffice_path:
        print("LibreOffice not found. Ensure LibreOffice is installed.")
        return

    try:
        # Command to run LibreOffice in headless mode
        command = [
            soffice_path,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(Path(output_path).parent),
            str(Path(input_path).resolve())
        ]
        subprocess.run(command, check=True, text=True, capture_output=True, shell=True)
        print(f"Successfully converted {input_path} to {output_path}")
    except subprocess.CalledProcessError as e:
        print(f"Error during conversion of {input_path}: {str(e)}\n{e.stderr}")
    except FileNotFoundError:
        print(f"LibreOffice executable not found at {soffice_path}.")
    except Exception as e:
        print(f"Unexpected error during conversion of {input_path}: {str(e)}")


def main():
    # Define input and output directories
    input_dir = "INPUT"
    output_dir = "OUTPUT"

    # Create output directory if it doesn't exist
    Path(output_dir).mkdir(exist_ok=True)

    # Check if input directory exists
    if not os.path.exists(input_dir):
        print(f"Input directory {input_dir} does not exist.")
        return

    # Get all .doc and .docx files in the input directory
    doc_files = [f for f in os.listdir(input_dir) if f.lower().endswith(('.doc', '.docx'))]

    if not doc_files:
        print(f"No .doc or .docx files found in {input_dir}.")
        return

    # Process each file
    for doc_file in doc_files:
        input_path = os.path.join(input_dir, doc_file)
        # Create output PDF filename by replacing extension
        output_filename = os.path.splitext(doc_file)[0] + ".pdf"
        output_path = os.path.join(output_dir, output_filename)

        print(f"Processing {doc_file}...")
        convert_doc_to_pdf(input_path, output_path)


if __name__ == "__main__":
    # Set console to UTF-8 for proper handling of Cyrillic characters
    os.system("chcp 65001 > nul")
    main()