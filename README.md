# Battery Data Analysis Tool

A Python tool for processing and visualizing battery cell voltage data from multiple Excel files. This tool combines sequential time-series data files and generates analytical visualizations of battery cell performance.

## Features

- Automatically scans directories for Excel files with timestamp patterns
- Combines sequential files into continuous datasets
- Generates two types of visualization plots:
  - Individual cell voltage trends over time
  - Statistical analysis including mean voltage, voltage spread, and cell variations
- Handles large datasets with efficient data processing
- Supports files with mixed data types (numeric and categorical)
- Automatic handling of missing or invalid data
- Customizable time-based resampling for large datasets

## Prerequisites

### System Requirements
- Python 3.x
- Linux/Ubuntu system
- Sufficient memory to handle your dataset size

### Required Python Packages
```bash
pandas
openpyxl
matplotlib
seaborn
numpy
```

## Installation

1. First, ensure you have the required system packages:
```bash
sudo apt update
sudo apt install python3-venv python3-pip
```

2. Create and activate a virtual environment:
```bash
# Create a virtual environment
python3 -m venv ~/python_env

# Activate the virtual environment
source ~/python_env/bin/activate
```

3. Install required Python packages:
```bash
pip install pandas openpyxl matplotlib seaborn numpy
```

## Usage

1. Activate the virtual environment (if not already activated):
```bash
source ~/python_env/bin/activate
```

2. Prepare your directory structure:
```
your_project_directory/
├── process.py
├── record/
│   ├── 20241028163106-00.xlsx
│   ├── 20241028173106-00.xlsx
│   └── ...
└── record/processed/
    └── (output files will be created here)
```

3. Run the script:
```bash
python3 process.py
```

### Input File Requirements

- Files must be Excel format (.xlsx or .xls)
- Filename format: "YYYYMMDDHHMMSS-00.xlsx"
- Expected data columns:
  - "Date & Time" column in a recognizable datetime format
  - "Cell Voltage X" columns containing numeric voltage values
  - Can handle additional non-numeric columns (they will be ignored for plotting)

### Output Files

For each continuous sequence, the script generates:
1. Combined Excel file:
   - Named: `sequence_X_YYYYMMDDHHMMSS_to_YYYYMMDDHHMMSS.xlsx`
   - Contains all data from the sequence

2. Visualization plots:
   - Cell voltage plot (`*_voltage_plot.png`):
     - Shows individual cell voltage trends
     - Lines for each cell voltage
     - Clear legend identifying each cell
   
   - Statistics plot (`*_voltage_stats.png`):
     - Mean voltage trend
     - Min-Max voltage range
     - Voltage spread analysis
     - Additional statistical measures

## Configuration Options

The script includes several configurable parameters:

- Resampling interval (default: 1 minute)
- Plot dimensions and DPI
- Color schemes and transparency levels
- Statistical calculation methods

To modify these, edit the corresponding variables in the script.

## Troubleshooting

### Common Issues

1. **Memory Errors**
   - Reduce the resampling interval
   - Process smaller sequences of files
   - Ensure sufficient system memory

2. **Data Type Errors**
   - The script automatically handles non-numeric data
   - Check your input files for unexpected data formats
   - Verify column names match expected patterns

3. **Missing or Empty Plots**
   - Verify input files contain non-zero voltage data
   - Check file naming patterns match requirements
   - Ensure correct column names in input files

4. **Performance Issues**
   - Large datasets are automatically resampled
   - Adjust resampling interval if needed
   - Consider processing fewer files at once

### Error Messages

- "No valid sequences found": Check file naming pattern
- "No voltage columns found": Verify column names in Excel files
- "No valid data for statistics": Check for non-zero voltage values

## Data Format Details

### Required Excel Column Format:
```
Date & Time | Cell Voltage 1 | Cell Voltage 2 | ... | Cell Voltage N | Other Columns
---------------------------------------------------------------------------
timestamp   | numeric value  | numeric value  | ... | numeric value  | any format
```

### Important Notes:
- Voltage values should be numeric
- Zero values are automatically filtered
- Non-numeric data in other columns is safely ignored
- Timestamps must be in a recognizable datetime format

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.

### Development Setup
1. Fork the repository
2. Create a development virtual environment
3. Install development dependencies
4. Submit pull requests with tests and documentation

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For bugs, feature requests, or questions:
1. Create an issue in the repository
2. Provide sample data if possible
3. Include error messages and system details

## Acknowledgments

- Built with pandas for efficient data processing
- Visualizations created using matplotlib and seaborn
- Inspired by battery monitoring system needs