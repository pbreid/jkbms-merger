import pandas as pd
import os
from datetime import datetime, timedelta
from pathlib import Path
import re
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.dates import DateFormatter

def parse_timestamp_from_filename(filename):
    """Extract timestamp from filename pattern YYYYMMDDHHMMSS-00.xlsx"""
    match = re.match(r'(\d{14})-\d+\.xlsx?', filename)
    if match:
        timestamp_str = match.group(1)
        return datetime.strptime(timestamp_str, '%Y%m%d%H%M%S')
    return None



def create_voltage_plots(df, output_path, sequence_name):
    """Create plots for cell voltages over time with enhanced statistics"""
    print("Starting plot creation...")
    
    try:
        # Convert 'Date & Time' to datetime if it's not already
        print("Converting timestamps...")
        df['Date & Time'] = pd.to_datetime(df['Date & Time'])
        df.set_index('Date & Time', inplace=True)
        
        # Find voltage columns
        voltage_columns = [col for col in df.columns if 'Cell Voltage' in col]
        
        if not voltage_columns:
            print("No voltage columns found in data")
            return
            
        # Create a new DataFrame with only numeric columns we want to plot
        plot_df = df[voltage_columns].copy()
        
        # Convert to numeric, forcing errors to NaN
        for col in plot_df.columns:
            plot_df[col] = pd.to_numeric(plot_df[col], errors='coerce')
        
        # Resample only the numeric data
        print("Resampling data...")
        df_resampled = plot_df.resample('0.5min').mean()
        df_resampled.reset_index(inplace=True)
        
        print("Creating main voltage plot...")
        plt.figure(figsize=(15, 10))
        
        # Track if we've plotted any data
        has_valid_data = False
        
        # Plot each cell voltage that has non-zero values
        for col in voltage_columns:
            # Only plot if the column has any non-zero values
            mask = df_resampled[col].notna() & (df_resampled[col] > 0)
            if mask.any():
                plt.plot(df_resampled.loc[mask, 'Date & Time'],
                        df_resampled.loc[mask, col],
                        label=col,
                        alpha=0.7,
                        linewidth=1)
                has_valid_data = True
        
        if not has_valid_data:
            print("No valid voltage data to plot")
            plt.close()
            return
        
        plt.title(f'Cell Voltages Over Time - {sequence_name}')
        plt.xlabel('Time')
        plt.ylabel('Voltage (V)')
        plt.gca().xaxis.set_major_formatter(DateFormatter('%Y-%m-%d %H:%M'))
        plt.xticks(rotation=45)
        
        if has_valid_data:
            plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize='small')
        
        # Save the first plot
        plot_path = output_path.parent / f"{output_path.stem}_voltage_plot.png"
        plt.savefig(plot_path, bbox_inches='tight', dpi=300)
        plt.close()
        
        print("Creating statistics plot...")
        plt.figure(figsize=(15, 10))
        
        # Calculate statistics only for non-zero values
        voltage_data = df_resampled[voltage_columns]
        mask = voltage_data.notna() & (voltage_data > 0)
        
        if not mask.any().any():
            print("No valid data for statistics calculation")
            plt.close()
            return
        
        mean_voltage = voltage_data[mask].mean(axis=1)
        min_voltage = voltage_data[mask].min(axis=1)
        max_voltage = voltage_data[mask].max(axis=1)
        
        if mean_voltage.notna().any():
            plt.plot(df_resampled['Date & Time'], 
                    mean_voltage,
                    label='Mean Voltage',
                    color='blue',
                    linewidth=2)
            
            valid_range = min_voltage.notna() & max_voltage.notna()
            if valid_range.any():
                plt.fill_between(df_resampled['Date & Time'][valid_range],
                               min_voltage[valid_range],
                               max_voltage[valid_range],
                               alpha=0.2,
                               color='blue',
                               label='Min-Max Range')
            
            # Add voltage spread
            ax2 = plt.gca().twinx()
            voltage_spread = max_voltage - min_voltage
            valid_spread = voltage_spread.notna()
            if valid_spread.any():
                ax2.plot(df_resampled['Date & Time'][valid_spread],
                        voltage_spread[valid_spread],
                        label='Voltage Spread',
                        color='red',
                        linestyle='--',
                        alpha=0.7)
            
            # Formatting
            plt.title(f'Voltage Statistics Over Time - {sequence_name}')
            plt.gca().set_xlabel('Time')
            plt.gca().set_ylabel('Voltage (V)')
            ax2.set_ylabel('Voltage Spread (V)')
            
            plt.gca().xaxis.set_major_formatter(DateFormatter('%Y-%m-%d %H:%M'))
            plt.xticks(rotation=45)
            
            # Combine legends
            lines1, labels1 = plt.gca().get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            if lines1 or lines2:
                ax2.legend(lines1 + lines2, labels1 + labels2,
                          bbox_to_anchor=(1.05, 1),
                          loc='upper left')
            
            # Save the statistics plot
            stats_plot_path = output_path.parent / f"{output_path.stem}_voltage_stats.png"
            plt.savefig(stats_plot_path, bbox_inches='tight', dpi=300)
        
        plt.close()
        print("Plots created successfully!")
        
    except Exception as e:
        print(f"Error during plot creation: {str(e)}")
        plt.close()  # Ensure any open figures are closed
        raise

def process_excel_files(directory_path, output_directory):
    """Main function to process Excel files and concatenate continuous sequences"""
    try:
        # Create output directory if it doesn't exist
        output_dir = Path(output_directory)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Get all Excel files in directory
        input_dir = Path(directory_path)
        if not input_dir.exists():
            raise FileNotFoundError(f"Input directory not found: {input_dir}")
            
        print(f"Processing files from: {input_dir}")
        print(f"Saving output to: {output_dir}")
        
        excel_files = [f for f in os.listdir(input_dir) 
                      if f.endswith(('.xlsx', '.xls')) and 
                      re.match(r'\d{14}-\d+\.xlsx?', f)]
        
        if not excel_files:
            print(f"No matching Excel files found in {input_dir}")
            return
            
        print(f"Found {len(excel_files)} Excel files")
        
        # Find continuous sequences
        sequences = find_continuous_sequences(excel_files)
        
        if not sequences:
            print("No valid sequences found to process")
            return
            
        print(f"Found {len(sequences)} sequences to process")
        
        # Process each sequence
        for seq_idx, sequence in enumerate(sequences, 1):
            print(f"\nProcessing sequence {seq_idx}/{len(sequences)}")
            
            # Initialize empty DataFrame for concatenation
            combined_df = pd.DataFrame()
            
            # Get start and end timestamps for naming
            start_ts = sequence[0][0]
            end_ts = sequence[-1][0]
            
            # Process each file in sequence
            for timestamp, filename in sequence:
                file_path = input_dir / filename
                print(f"  Reading file: {filename}")
                df = safe_read_excel(file_path)
                if df is not None and not df.empty:
                    combined_df = pd.concat([combined_df, df], ignore_index=True)
            
            if combined_df.empty:
                print(f"No valid data found for sequence {seq_idx}")
                continue
            
            # Generate output filename
            output_filename = f"sequence_{seq_idx}_{start_ts.strftime('%Y%m%d%H%M%S')}_to_{end_ts.strftime('%Y%m%d%H%M%S')}.xlsx"
            output_path = output_dir / output_filename
            
            # Save combined data
            if safe_write_excel(combined_df, output_path):
                print(f"Created sequence {seq_idx}: {output_filename}")
                print(f"  - Contains {len(sequence)} files")
                print(f"  - Total records: {len(combined_df)}")
                
                # Create plots for this sequence
                sequence_name = f"Sequence {seq_idx} ({start_ts.strftime('%Y-%m-%d %H:%M')} to {end_ts.strftime('%Y-%m-%d %H:%M')})"
                create_voltage_plots(combined_df, output_path, sequence_name)
            else:
                print(f"Failed to create sequence {seq_idx}: {output_filename}")
                
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        raise



def find_continuous_sequences(files, max_gap_hours=1.1):
    """Group files into continuous sequences"""
    if not files:
        print("No files found to process!")
        return []
    
    # Sort files by timestamp
    timestamped_files = []
    for f in files:
        ts = parse_timestamp_from_filename(f)
        if ts:
            timestamped_files.append((ts, f))
    
    if not timestamped_files:
        print("No valid timestamped files found!")
        return []
    
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
        df = pd.read_excel(file_path, engine='openpyxl')
        return df
    except Exception as e:
        print(f"Error reading {file_path}: {str(e)}")
        return None

def safe_write_excel(df, output_path):
    """Safely write Excel file with error handling"""
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, index=False)
        return True
    except Exception as e:
        print(f"Error writing to {output_path}: {str(e)}")
        return False

def process_excel_files(directory_path, output_directory):
    """Main function to process Excel files and concatenate continuous sequences"""
    try:
        # Convert to Path objects and resolve absolute paths
        input_dir = Path(directory_path).resolve()
        output_dir = Path(output_directory).resolve()
        
        # Check if input directory exists
        if not input_dir.exists():
            print(f"Input directory not found: {input_dir}")
            return
        
        # Create output directory if it doesn't exist
        output_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"Processing files from: {input_dir}")
        print(f"Saving output to: {output_dir}")
        
        # Get all Excel files in directory
        excel_files = [f for f in os.listdir(input_dir) 
                      if f.endswith(('.xlsx', '.xls')) and 
                      re.match(r'\d{14}-\d+\.xlsx?', f)]
        
        if not excel_files:
            print(f"No matching Excel files found in {input_dir}")
            return
        
        print(f"Found {len(excel_files)} Excel files")
        
        # Find continuous sequences
        sequences = find_continuous_sequences(excel_files)
        
        if not sequences:
            print("No valid sequences found to process")
            return
        
        print(f"Found {len(sequences)} sequences to process")
        
        # Process each sequence
        for seq_idx, sequence in enumerate(sequences, 1):
            print(f"\nProcessing sequence {seq_idx}/{len(sequences)}")
            dfs_to_concat = []
            
            # Get start and end timestamps for naming
            start_ts = sequence[0][0]
            end_ts = sequence[-1][0]
            
            # Process each file in sequence
            for timestamp, filename in sequence:
                file_path = input_dir / filename
                print(f"  Reading file: {filename}")
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
                
                # Create plots for this sequence
                sequence_name = f"Sequence {seq_idx} ({start_ts.strftime('%Y-%m-%d %H:%M')} to {end_ts.strftime('%Y-%m-%d %H:%M')})"
                create_voltage_plots(combined_df, output_path, sequence_name)
            else:
                print(f"Failed to create sequence {seq_idx}: {output_filename}")
                
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        raise

# Example usage
if __name__ == "__main__":
    # Get the current directory where the script is located
    script_dir = Path(__file__).parent.resolve()
    
    # Define input and output directories relative to the script location
    input_directory = script_dir / "record/00"
    output_directory = script_dir / "record" / "processed"
    
    print("Starting Excel file processing...")
    process_excel_files(input_directory, output_directory)
    print("Processing complete!")