# Lactate Test App

Lactate Test App is a Python application that helps athletes and coaches analyze lactate testing data. The app allows users to input lactate test data manually or upload data from an Excel file. It calculates key metrics such as Functional Threshold Power (FTP), Lactate Threshold 1 (LT1), Lactate Threshold 2 (LT2), and FATmax. Additionally, it provides visualizations of the data.

## Features

- Manually input lactate test data (Stage, Lactate, Heart Rate, Power).
- Upload lactate test data from an Excel file.
- Automatically generate stage numbers if the "Stage" column is missing in the uploaded Excel file.
- Calculate FTP, LT1, LT2, and FATmax.
- Visualize lactate levels, heart rate, and power output across different stages.
- Scrollable graph display.

## Requirements

- Python 3.x
- pandas
- numpy
- matplotlib
- tkinter

## Installation

1. **Clone the repository**:
    ```sh
    git clone https://github.com/yourusername/lactate-test-app.git
    cd lactate-test-app
    ```

2. **Install required Python packages**:
    ```sh
    pip install pandas numpy matplotlib
    ```

3. **Run the application**:
    ```sh
    python lactate_test_app.py
    ```

## Usage

### Manual Data Input

1. Open the application.
2. In the "Data Input" tab, enter the stage, lactate level, heart rate, and power for each test stage.
3. Click "Add Data" to add the data to the table.

### Uploading Excel Data

1. Click the "Upload Excel" button.
2. Select the Excel file containing the lactate test data.
3. If the Excel file does not contain a "Stage" column, the app will generate stage numbers automatically based on the number of rows.

### Calculating Metrics

1. Click the "Calculate FTP", "Calculate LT1", "Calculate LT2", and "Calculate FATmax" buttons to compute the respective values.
2. The calculated values will be displayed next to their corresponding labels.

### Plotting Data

1. Click the "Plot Data" button to generate graphs.
2. Switch to the "Graphs" tab to view the visualizations.

## File Structure

```plaintext
lactate-test-app/
│
├── lactate_test_app.py   # Main application script
├── README.md             # This readme file
└── requirements.txt      # Python package requirements
```

## Example Data Format

The Excel file should contain columns for “Stage”, “Lactate”, “Heart Rate”, and “Power”. If the “Stage” column is missing, the app will generate stage numbers automatically.

| Stage | Lactate | Heart Rate | Power |
|-------|---------|------------|-------|
| 1     | 1.2     | 130        | 200   |
| 2     | 1.5     | 135        | 210   |
| 3     | 1.8     | 140        | 220   |
| ...   | ...     | ...        | ...   |

## License

This project is licensed under a custom license for use strictly as part of a Software as a Service (SaaS) offering provided by the copyright holder. For any usage outside of the permissions granted, please contact [HøjFrekvens] at [jeanjohnsen@pm.me].

## Contact
For any questions or issues, please open an issue on the GitHub repository or contact me at jeanjohnsen@pm.me