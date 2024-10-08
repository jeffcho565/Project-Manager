# Project Management GUI Application

This Python application provides a simple GUI for managing project information using an Excel file as the database. The application allows users to enter project details, save them, and view the saved entries in a table.

## Features

- **Add Project Information**: Users can input the project name, their name, and the date for the project.
- **Save Data**: The entered data is saved to an Excel file (`project_db.xlsx`), located in the same directory as the script.
- **Duplicate Entry Prevention**: The application checks for duplicate entries before saving.
- **View Saved Data**: The saved project information is displayed in a table format within the GUI.
- **Date Selection**: The user can select the project date using dropdowns for year, month, and day.

## Requirements

- Python 3.x
- `pandas` library
- `openpyxl` library (for handling Excel files)
- `tkinter` library (for the GUI)

## Installation

1. Clone this repository or download the script files.
2. Ensure you have Python 3.x installed on your system.
3. Install the required Python libraries using pip:

    ```bash
    pip install pandas openpyxl
    ```

4. Run the Python script:

    ```bash
    python project_management_gui.py
    ```

## Usage

1. **Name**: Enter your name in the "Name" field.
2. **Project Name**: Enter the project name in the "Project name" field.
3. **Select Date**: Use the dropdowns to select the date for the project.
4. **Save Data**: Click the "Save Data" button to save the project details.
5. **View Data**: The saved data will automatically appear in the table below the input fields.

## File Structure

- `project_management_gui.py`: The main script that runs the application.
- `project_db.xlsx`: The Excel file where the project data is saved. This file is created automatically if it does not exist.

## Important Notes

- The application automatically loads and displays any existing data from `project_db.xlsx` when it starts.
- The application prevents saving duplicate entries based on the combination of project name, date, and user.

## Future Enhancements

- Add the ability to edit or delete existing entries.
- Improve the UI for better user experience.
- Add filtering options to search through saved data.

## License

This project is open-source and available under the MIT License.
