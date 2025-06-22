# Directory to Excel

This Python script lists all files in a specified directory and exports the file information to an Excel file.

## Features

*   Recursively lists all files in a directory and its subdirectories.
*   Exports file path, size (in bytes), and last modified time to an Excel file.
*   Displays a progress bar using the `tqdm` library.
*   Handles `FileNotFoundError` and `OSError` exceptions.

## Requirements

*   Python 3.6 or higher
*   `openpyxl` library
*   `tqdm` library

## Installation

1.  Clone this repository:

    ```bash
    git clone git@github.com:wulukewu/file-lister-2025.git
    ```

2.  Install the required Python packages:

    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  Run the script:

    ```bash
    python list_files.py
    ```

2.  The script will prompt you to enter the root directory to list files from and the desired output Excel file path.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please submit a pull request.