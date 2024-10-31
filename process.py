import pandas as pd
import os
from datetime import datetime, timedelta
from pathlib import Path
import re
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

def parse_timestamp_from_filename(filename):
    """Extract timestamp from filename pattern YYYYMMDDHHMMSS-00.xlsx"""
    match = re.match(r'(\d{14})-\d+\.xlsx?', filename)
    if match:
        timestamp_str = match.group(1)
        return datetime.strptime(timestamp_str, '%Y%m%d%H%M%S')
    return None

def find_continuous_sequences(files, max_gap_hours=1.1):  # allowing 6 min tolerance
    """Group files into continuous sequences"""
    if not files:
        return []
    
    # Sort files by timestamp
    timestamped_files = []
    for f in files:
        ts = parse_timestamp_from_filename(f)
        if ts:
            timestamped_files.append((ts, f))
    
    timestamped_files.sort()  # Sort by timestamp
    
    sequences = []
    current_sequence = [timestamped_files[0]]
    
    for i in range(1, len(timestamped_files)):
        current_ts, current_file = timestamped_files[i]
        prev_ts, prev_file = timestamped_files[i-1]
        
        time_diff = current_ts - prev_ts
        
        if time_diff <= timedelta(hours=max_gap_hours):
            current_sequence.append(timestamped_files[i])
        else:
            sequences.append(current_sequence)
            current_sequence = [timestamped_files[i]]
    
    sequences.append(current_sequence)
    return sequences

def safe_read_excel(file_path):
    """Safely read Excel file with error handling"""
    try:
        # Try reading with default engine
        df = pd.read_excel(file_path, engine='openpyxl')
        return df
    except Exception as e:
        print(f"Error reading {file_path}: {str(e)}")
        return None

def safe_write_excel(df, output_path):
    """Safely write Excel file with error handling and proper engine selection"""
    try:
        # Create a new Excel writer object with the xlsxwriter engine
        with pd.ExcelWriter(output_path, engine='xlsxwriter', mode='w') as writer:
            # Write the DataFrame to Excel
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Auto-adjust columns' width
            for idx, col in enumerate(df.columns):
                series = df[col]
                max_len = max(
                    series.astype(str).apply(len).max(),  # len of largest item
                    len(str(series.name))  # len of column name/header
                ) + 1  # adding a little extra space
                worksheet.set_column(idx, idx, max_len)  # set column width
        
        return True
    except Exception as e:
        print(f"Error writing to {output_path}: {str(e)}")
        return False

def process_excel_files(directory_path, output_directory):
    """Main function to process Excel files and concatenate continuous sequences"""
    # Create output directory if it doesn't exist
    output_dir = Path(output_directory)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Get all Excel files in directory
    excel_files = [f for f in os.listdir(directory_path) 
                  if f.endswith(('.xlsx', '.xls')) and 
                  re.match(r'\d{14}-\d+\.xlsx?', f)]
    
    # Find continuous sequences
    sequences = find_continuous_sequences(excel_files)
    
    # Process each sequence
    for seq_idx, sequence in enumerate(sequences, 1):
        # Initialize empty DataFrame for concatenation
        dfs_to_concat = []
        
        # Get start and end timestamps for naming
        start_ts = sequence[0][0]
        end_ts = sequence[-1][0]
        
        # Process each file in sequence
        for timestamp, filename in sequence:
            file_path = Path(directory_path) / filename
            df = safe_read_excel(file_path)
            if df is not None:
                dfs_to_concat.append(df)
        
        if not dfs_to_concat:
            print(f"No valid data found for sequence {seq_idx}")
            continue
            
        # Combine all DataFrames
        combined_df = pd.concat(dfs_to_concat, ignore_index=True)
        
        # Generate output filename
        output_filename = f"sequence_{seq_idx}_{start_ts.strftime('%Y%m%d%H%M%S')}_to_{end_ts.strftime('%Y%m%d%H%M%S')}.xlsx"
        output_path = output_dir / output_filename
        
        # Save combined data
        if safe_write_excel(combined_df, output_path):
            print(f"Created sequence {seq_idx}: {output_filename}")
            print(f"  - Contains {len(sequence)} files")
            print(f"  - Total records: {len(combined_df)}")
        else:
            print(f"Failed to create sequence {seq_idx}: {output_filename}")

# Example usage
if __name__ == "__main__":
    input_directory = "record/00"
    output_directory = "record/processed"
    
    process_excel_files(input_directory, output_directory)