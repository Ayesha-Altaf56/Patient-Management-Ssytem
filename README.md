# Patient Management System

A simple Patient Management System built in Python that allows you to manage patient records. The application uses an Excel file (`Patient_Data.xlsx`) to store and retrieve data.

This project includes two versions:
- A Command-Line Interface (CLI) version (`maincode.py`)
- A Graphical User Interface (GUI) version built with Tkinter (`frontend.py`)

## Features

- **Add Patient**: Register a new patient with details (ID, Name, Gender, Room Number, Disease, Age).
- **View Patients**: Display a list of all registered patients.
- **Search Patient**: Find a specific patient's details using their Patient ID.
- **Update Patient**: Modify existing patient information.
- **Delete Patient**: Remove a patient record from the system.
- **Data Persistence**: All data is saved to and loaded from an Excel workbook (`Patient_Data.xlsx`).

## Prerequisites

To run this project, you need to have Python installed on your system. You also need to install the required `openpyxl` library to handle Excel files.

Install the required dependency using pip:

```bash
pip install openpyxl
```

*(Note: Tkinter is usually included with the standard Python installation. If it's missing on your OS, you may need to install it separately).*

## How to Run

1. Clone or download this repository.
2. Navigate to the project directory in your terminal.
3. Choose which version you want to run:

**To run the GUI version (Recommended):**
```bash
python frontend.py
```

**To run the CLI version:**
```bash
python maincode.py
```

## Project Structure

- `frontend.py`: The Tkinter-based graphical user interface application.
- `maincode.py`: The terminal-based command-line application.
- `Patient_Data.xlsx`: The Excel file where patient records are stored (automatically created if it doesn't exist).
