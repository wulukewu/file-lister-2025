import os
import openpyxl
import time
import imagehash
from PIL import Image
from tqdm import tqdm

def list_files_and_find_duplicates(root_dir, excel_file, similarity_threshold=5):
    """
    Lists files, identifies duplicates, groups them, sets 'Keep' for unique files,
    and sorts output for easier review.
    """

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["File Path", "File Size (bytes)", "Last Modified", "Duplicate Group", "Keep", "Delete"])

    file_data = []  # Store file data before writing to Excel
    image_hashes = {}
    duplicate_groups = {}

    total_files = 0
    for dirpath, dirnames, filenames in os.walk(root_dir):
        total_files += len(filenames)

    with tqdm(total=total_files, desc="Listing and Analyzing Files", unit="file") as pbar:
        for dirpath, dirnames, filenames in os.walk(root_dir):
            for filename in filenames:
                filepath = os.path.join(dirpath, filename)
                if is_image_or_video(filename):
                    try:
                        file_size = os.path.getsize(filepath)
                        modified_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(os.path.getmtime(filepath)))

                        try:
                            img = Image.open(filepath)
                            hash_value = imagehash.average_hash(img)

                            # Check for duplicates
                            found_duplicate = False
                            for existing_hash, file_list in duplicate_groups.items():
                                if hash_value - existing_hash < similarity_threshold:
                                    duplicate_groups[existing_hash].append(filepath)
                                    found_duplicate = True
                                    break

                            if not found_duplicate:
                                duplicate_groups[hash_value] = [filepath]

                        except Exception as e:
                            print(f"Warning: Error processing image {filepath} for duplicate detection: {e}")
                            hash_value = None # Indicate error during hash calculation


                        file_data.append([filepath, file_size, modified_time, None, "No", "No"])  # Append to file_data

                    except FileNotFoundError:
                        print(f"Warning: File not found: {filepath}")
                        file_data.append([filepath, "N/A", "N/A", None, "No", "No"])
                    except OSError as e:
                        print(f"Warning: Error accessing file {filepath}: {e}")
                        file_data.append([filepath, "N/A", "N/A", None, "No", "No"])
                    except Exception as e:
                        print(f"Warning: An unexpected error occurred processing {filepath}: {e}")
                        file_data.append([filepath, "N/A", "N/A", None, "No", "No"])
                else:
                    file_data.append([filepath, "N/A", "N/A", None, "No", "No"])  # Append to file_data (handles non-image or video)
                pbar.update(1)

    # Assign group numbers and 'Keep' values
    grouped_files = []
    unique_files = []
    group_counter = 1
    for hash_value, file_list in duplicate_groups.items():
        if len(file_list) > 1:  # Duplicate group
            for item in file_data:
                if item[0] in file_list:  # item[0] is the filepath
                    item[3] = group_counter # Group Number
                    grouped_files.append(item)
            group_counter += 1
        else:  # Unique file
            for item in file_data:
                if item[0] == file_list[0]:  # find and set to Yes
                    item[4] = "Yes"  # Set 'Keep' to 'Yes'
                    unique_files.append(item)



    # Write grouped files first, then unique files
    for row in grouped_files:
        sheet.append(row)
    for row in unique_files:
        sheet.append(row)

    try:
        print(f"Saving to: {excel_file}")
        workbook.save(excel_file)
        print(f"File list saved to: {excel_file}")
    except PermissionError as e:
        print(f"Error: Unable to save file to {excel_file}. Check write permissions: {e}")
    except FileNotFoundError as e:
        print(f"Error: File or directory not found: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


def is_image_or_video(filename):
    """Checks if a file is an image or video based on extension."""
    ext = filename.lower().split('.')[-1]
    image_extensions = ["jpg", "jpeg", "png", "gif", "bmp", "webp"]
    video_extensions = ["mp4", "avi", "mov", "mkv", "webm"]
    return ext in image_extensions or ext in video_extensions


if __name__ == "__main__":
    root_directory = input("Enter the root directory to list files from: ")
    output_excel_file = input("Enter the output Excel file path (e.g., files.xlsx): ")
    list_files_and_find_duplicates(root_directory, output_excel_file)