# Web-Based Duplicate File Manager

This web application helps you identify and manage duplicate image and video files within a specified directory. It provides a user-friendly interface to list files, detect potential duplicates, set folder priorities, and delete unwanted files.

## Features

*   **File Listing:** Recursively lists all image and video files in a directory and its subdirectories.
*   **Duplicate Detection:** Identifies potential duplicate files based on content (image hashing).
*   **Directory Priority:** Allows you to prioritize certain directories, automatically marking files in lower-priority directories for deletion if duplicates exist.
*   **Web-Based Interface:** Provides an intuitive web interface for managing files.
*   **Image Preview:** Displays thumbnails of image files in the file list.
*   **Split File Path:** Shows the file name and directory path in separate columns for better readability.
*   **Shortened Directory Paths:** Displays directory paths relative to the root directory.
*   **Drag-and-Drop Folder Priority:** Allows you to visually set folder priority by dragging and dropping folder names in a list.

## Requirements

*   Python 3.6 or higher
*   Flask
*   openpyxl
*   imagehash
*   Pillow
*   jQuery
*   jQuery UI

## Installation

1.  Clone this repository:

    ```bash
    git clone <repository_url>
    ```

2.  Install the required Python packages:

    ```bash
    pip install Flask openpyxl imagehash Pillow
    ```

3.  Install jQuery and jQuery UI
    * This is already included using CDN on the HTML but if you want to install it locally please follow the instruction on Jquery Website (https://jquery.com/)

## Usage

1.  **Run the Flask Application:**

    ```bash
    python app.py
    ```

2.  **Open the Web Interface:** Open your web browser and go to `http://127.0.0.1:5000/`

3.  **Set the Root Directory:** Enter the root directory you want to analyze in the "Root Directory" field and click "Set Directory."

4.  **List Files and Find Duplicates:** Click the "List Files and Find Duplicates" button. The file list will be displayed in the table.

5.  **Set Folder Priority (Optional):**
    *   The application will automatically list directory, drag it based on your desired order.
    *   Click the "Apply Priority" button. The application will automatically mark files in lower-priority directories for deletion if duplicates exist.
    *   Review the file table, and you can update the priority again

6.  **Review and Adjust:** Review the "Keep" and "Delete" values in the table. You can manually change these values if needed. For each group, only one file should be marked as "Keep."

7.  **Delete Marked Files:** Click the "Delete Marked Files" button to delete the files that are marked for deletion.

## Important Notes

*   **Carefully Review:** Before deleting any files, carefully review the "Keep" and "Delete" values in the table. Ensure that you are only deleting files that you no longer need.
*   **Folder Priority:** Setting folder priority is a powerful tool, but it's important to understand how it works. The application will prioritize files in the folders listed *first*.
*   **Backup:** It is highly recommended to back up your data before using this application.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please submit a pull request.