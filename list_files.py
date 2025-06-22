import os
import openpyxl
import time
from tqdm import tqdm

def list_files_to_excel(root_dir, excel_file):
    """
    Lists all files in a directory and exports the file information to an Excel file,
    displaying a progress bar.

    Args:
        root_dir (str): The root directory to list files from.
        excel_file (str): The output Excel file path.
    """

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["File Path", "File Size (bytes)", "Last Modified"])  # Write headers

    # Calculate the total number of files for the progress bar
    total_files = 0
    for dirpath, dirnames, filenames in os.walk(root_dir):
        total_files += len(filenames)

    with tqdm(total=total_files, desc="Processing Files", unit="file") as pbar:
        for dirpath, dirnames, filenames in os.walk(root_dir):
            for filename in filenames:
                filepath = os.path.join(dirpath, filename)
                try:
                    file_size = os.path.getsize(filepath)
                    modified_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(os.path.getmtime(filepath)))

                    sheet.append([filepath, file_size, modified_time])
                except FileNotFoundError:
                    print(f"Warning: File not found: {filepath}")  # Handle file not found
                except OSError as e:
                    print(f"Warning: Error accessing file {filepath}: {e}")
                pbar.update(1)  # Update the progress bar

    try:
        print(f"Saving to: {excel_file}")
        workbook.save(excel_file)
        print(f"File list saved to: {excel_file}")
    except PermissionError as e:
        print(f"Error: Unable to save file to {excel_file}. Please check your write permissions: {e}")
    except FileNotFoundError as e:
        print(f"Error: File or directory not found: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


if __name__ == "__main__":
    root_directory = input("Enter the root directory to list files from: ")
    output_excel_file = input("Enter the output Excel file path (e.g., files.xlsx): ")

    list_files_to_excel(root_directory, output_excel_file)