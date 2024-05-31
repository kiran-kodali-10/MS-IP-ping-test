
# IP Address Ping Checker

This script reads an Excel file containing IP addresses, pings each IP address, and updates the status and timestamp in the Excel file. It handles keyboard interrupts gracefully by saving progress before exiting.

## Prerequisites

- Python 3.x installed on your system
- pip (Python package installer)

## Setup Instructions

### 1. Clone the Repository

If the script is hosted in a repository, clone it using git:

```bash
git clone <repository_url>
cd <repository_directory>
```

### 2. Create a Virtual Environment

#### On Windows

Open Command Prompt or PowerShell and run:

```bash
python -m venv venv
venv\Scripts\activate
```

#### On Linux/MacOS

Open a terminal and run:

```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Required Packages

With the virtual environment activated, install the required packages using pip and the provided `requirements.txt` file:

```bash
pip install -r requirements.txt
```

### 4. Prepare the Excel File

Ensure your Excel file has a column named "expected IP" (case insensitive). Optionally, add "Status" and "Last Updated" columns if they don't exist. The script will add them if missing.

### 5. Run the Script

Run the script with the path to your Excel file as an argument:

```bash
python script.py path_to_your_file.xlsx
```

### 6. Handle Interrupts

If you need to stop the script, press `Ctrl+C`. The script will handle the interrupt, save progress, and exit gracefully.

## Example

```bash
python script.py example.xlsx
```

## Script Description

The script performs the following steps:

1. Loads the Excel file.
2. Reads IP addresses from the "expected IP" column.
3. Pings each IP address.
4. Updates the "Status" and "Last Updated" columns based on the ping results.
5. Saves progress and handles keyboard interrupts gracefully.

## Dependencies

All dependencies are listed in the `requirements.txt` file.

### Additional Notes

- Ensure your IP addresses are correctly formatted in the "expected IP" column.
- The script will skip rows with missing or empty IP addresses.
