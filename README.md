# Excel File Sequence Processor

This Python script processes Excel files containing time-series data, combining sequential files into continuous datasets. It's specifically designed to handle files with names following the pattern "YYYYMMDDHHMMSS-00.xlsx" where files are typically generated hourly with 1Hz data.

## Features

- Automatically scans directories for Excel files with timestamp patterns
- Groups files into continuous sequences based on timestamps
- Concatenates data from sequential files into single output files
- Handles large datasets with proper memory management
- Auto-adjusts column widths in output files
- Provides progress information during processing

## Prerequisites

### System Requirements
- Python 3.x
- Linux/Ubuntu system

### Required Python Packages
- pandas
- openpyxl
- xlsxwriter

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
pip install pandas openpyxl xlsxwriter
```

## Usage

1. Activate the virtual environment (if not already activated):
```bash
source ~/python_env/bin/activate
```

2. Run the script:
```bash
python excel_processor.py
```

### Directory Structure
Place your input files in a directory structure like this:
```
/your/path/
├── input/
│   ├── 20241028163106-00.xlsx
│   ├── 20241028173106-00.xlsx
│   └── ...
└── processed/
    └── (output files will be created here)
```

### Configuration
Edit the following lines in the script to match your directory paths:
```python
input_directory = "path/to/your/input/files"
output_directory = "path/to/your/output/files"
```

## Output

The script will create files named in the following format:
```
sequence_<number>_<start_timestamp>_to_<end_timestamp>.xlsx
```

Example:
```
sequence_1_20241028163106_to_20241028173106.xlsx
```

## Troubleshooting

### Common Issues

1. **Excel File Corruption**
   - Ensure all required packages are installed correctly
   - Verify input files are not corrupted
   - Check available disk space

2. **Permission Issues**
   - Ensure you have write permissions in the output directory
   ```bash
   chmod 755 /path/to/output/directory
   ```

3. **Memory Issues**
   - If processing large files, ensure sufficient system memory is available
   - Consider reducing the batch size by modifying the script

### Package Installation Issues

If you encounter package installation issues, you can alternatively use system packages:
```bash
sudo apt install python3-xlsxwriter python3-pandas python3-openpyxl
```

## Limitations

- Files must follow the naming convention "YYYYMMDDHHMMSS-00.xlsx"
- Assumes 1Hz data with approximately 3600 records per file
- Requires sufficient system memory to process large datasets

## Contributing

Feel free to submit issues and enhancement requests!

## License

This project is licensed under the MIT License - see the LICENSE file for details.