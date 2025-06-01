# PDF Manager

## Overview

PDF Manager is a Python application designed to help users manage and manipulate PDF and PNG files, particularly for compiling factory paperwork. It provides a graphical user interface (GUI) built with PySide6 for easy interaction.

Key features include:
- Listing PDF and PNG files from a specified directory.
- Two view modes for browsing files: a list view and a preview mode.
- Reordering files for compilation.
- Selecting specific files for actions.
- Combining selected or all files into a single PDF.
- Opening files in external applications like Inkscape and GIMP.
- Configuration options for executable paths (Inkscape, GIMP) and working directories.

## Features

- **File Listing and Preview**: Displays files from the "Factory Paperwork Folder" with options to view as a simple list or with image previews (for PNGs and first page of PDFs).
- **File Ordering**: Users can set the order of files using spinboxes. This order is used when combining files.
- **File Selection**: Checkboxes allow users to select multiple files for batch operations.
- **Combine Files**:
    - **Combine ALL**: Merges all files listed in the main view (respecting their set order) into a single PDF.
    - **Combine SELECTED**: Merges only the checked files (respecting their set order) into a single PDF.
- **External Editing**:
    - **Open Last Output PDF in Inkscape**: Opens the most recently created combined PDF in Inkscape.
    - **Open SELECTED in Inkscape**: Opens all currently selected files in Inkscape.
    - **Open SELECTED in GIMP**: Opens all currently selected files in GIMP.
- **Directory Configuration**:
    - **Source Files Directory**: Specifies the folder from which to copy files.
    - **Factory Paperwork Directory**: The primary working directory where files are listed, managed, and from which compilations occur.
- **Copy Files**: Copies PDF and PNG files from the "Source Files Directory" to the "Factory Paperwork Directory".
- **Settings**: A dialog to configure the paths to Inkscape and GIMP executables.
- **Status Bar**: Provides feedback on ongoing operations and errors.

## Requirements

The application relies on the following Python libraries:

- **PySide6**: For the graphical user interface.
- **pypdf**: For PDF manipulation tasks like merging.
- **PyMuPDF (fitz)**: For generating PDF previews and converting PNG images to PDF format for inclusion in combined documents.

These dependencies are listed in the `requirements.txt` file and can be installed using pip:
```bash
pip install -r requirements.txt
```

## Setup and Usage

1.  **Clone the repository or download the source code.**
2.  **Install dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
3.  **Run the application**:
    ```bash
    python pdf_manager.py
    ```
4.  **Initial Configuration**:
    *   Upon first launch, or if not configured, you might need to set the paths to your Inkscape and GIMP executables. This can be done via the "Settings" > "Configure Executable Paths..." menu.
    *   Set the "Source Files Directory" and "Factory Paperwork Directory" in the main window. The application will list files from the "Factory Paperwork Directory".
5.  **Using the Application**:
    *   Use the "Browse..." buttons to select your working directories.
    *   Files in the "Factory Paperwork Directory" will be displayed.
    *   Toggle between "List View" and "Preview View" using the "Toggle View Mode" button.
    *   Use the spinboxes next to each file to define the compilation order.
    *   Check the boxes next to files to select them for actions like "Combine SELECTED", "Open SELECTED in Inkscape", or "Open SELECTED in GIMP".
    *   Use the "Copy Files" button to transfer files from your source to the factory directory.
    *   The "Combine ALL" or "Combine SELECTED" buttons will prompt you to save the resulting PDF.
    *   The status bar at the bottom will show messages about current operations.

## File Structure

-   `pdf_manager.py`: The main Python script containing the application logic and GUI.
-   `requirements.txt`: Lists the Python dependencies.
-   `.gitignore`: Specifies intentionally untracked files that Git should ignore.
-   `lib/`: This directory appears to contain local copies or components of the required libraries (PyMuPDF, PySide6, pypdf). In a typical Python project, these would be managed by `pip` in a virtual environment rather than being included directly in the repository, unless there's a specific reason for bundling them (e.g., for portability or when dealing with non-standard builds).

## Notes

-   The application uses `subprocess.Popen` to launch external applications (Inkscape, GIMP). Ensure these are installed and their paths are correctly configured in the application's settings.
-   PNG file inclusion during PDF combination requires PyMuPDF. If PyMuPDF is not available, PNGs will be skipped.
-   The application maintains a `config.json` file (not explicitly shown in the provided context, but inferred from `save_config` and `self.config` usage) in the user's application data directory to store settings like directory paths and executable paths.

## Potential Future Enhancements

-   Drag-and-drop reordering of files in the GUI.
-   More robust error handling and logging.
-   Support for more file types.
-   Password protection options for combined PDFs.
-   A progress bar for lengthy operations like combining many large files.
