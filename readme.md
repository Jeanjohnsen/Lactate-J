# LactateLab

LactateLab is a graphical user interface (GUI) application for managing and analyzing lactate test data for athletes. The application allows users to input data manually, upload data from Excel files, and calculate key performance metrics such as Functional Threshold Power (FTP), Lactate Threshold 1 (LT1), Lactate Threshold 2 (LT2), and FATmax. The application also provides functionality to compare new test data with old test data and visualize the results in graphs.

## Features

- **Data Input**: Manually input lactate, heart rate, and power data.
- **Upload Data**: Upload test data from Excel files.
- **Data Management**: Add, edit, and clear data entries.
- **Export Data**: Export data to CSV and Excel formats.
- **Performance Metrics**: Calculate FTP, LT1, LT2, and FATmax.
- **Graphical Visualization**: Plot and compare test data with visual graphs.
- **Report Generation**: Export results and graphs to PDF.

## Getting Started

### Prerequisites

- Python 3.x
- Required Python packages: `pandas`, `numpy`, `matplotlib`, `tkinter`, `reportlab`

### Installation

1. **Clone the repository**:
   ```sh
   git clone https://github.com/yourusername/LactateLab.git
   cd LactateLab
   pip install pandas numpy matplotlib reportlab
   python main.py
   ```
# Usage

## Data Input

- **Manual Entry**:	
    Enter values for lactate, heart rate, and power in the provided fields.
    Click “Add Data” to add the entry to the table.
- **Upload Data**:
    Click “Upload Excel” to upload data from an Excel file.
    The application supports .xls and .xlsx file formats.
- **Export Data**:
    Click “Export to CSV” or “Export to Excel” to save the table data.
- **Clear Data**:
    Click “Clear Data” to remove all entries from the table.

## Calculations

- **Calculate Metrics**:
    Click “Calculate All” to compute FTP, LT1, LT2, and FATmax based on the entered data.
    The results will be displayed next to their respective labels.

## Visualization

1. **Data Input**:
    Click “Plot Data” to generate graphs for lactate levels, heart rate, and power output.
2.	**Compare Tests**:
    Click “Upload Old Test” to upload an old test data file.
    Click “Compare Tests” to compare new test data with old test data.
    Use the “Show New Test” checkbox to toggle the visibility of the new test data in the comparison graph.

## Reporting

- **Export to PDF**: Click “Export to PDF” to generate a PDF report containing the test results and graphs.

# File Structure
```
LactateLab/
│
├── main.py              # Main application script
├── README.md            # This readme file
├── data/                # Directory for storing data files (if applicable)
├── tests/               # Directory for test scripts
└── requirements.txt     # List of required Python packages
```

# License

This project is licensed under a custom license for use strictly as part of a Software as a Service (SaaS) offering provided by the copyright holder. For any usage outside of the permissions granted, please contact [Your Name or Company] at [Your Contact Information].
