import pandas as pd
import os
import math
from pathlib import Path

def split_excel_csv_to_10k_rows(folder_path, output_format='xlsx'):
    """
    Split Excel/CSV files in a folder into 10,000 row chunks
    
    Args:
        folder_path (str): Path to folder containing Excel/CSV files
        output_format (str): Output format - 'xlsx' or 'csv' (default: 'xlsx')
    """
    
    folder = Path(folder_path)
    
    # Check if folder exists
    if not folder.exists():
        print(f"Error: Folder '{folder_path}' does not exist!")
        return
    
    # Validate output format
    if output_format.lower() not in ['xlsx', 'csv']:
        print("Error: Output format must be 'xlsx' or 'csv'")
        return
    
    output_format = output_format.lower()
    
    # Find all Excel and CSV files
    file_patterns = ['*.xlsx', '*.xls', '*.csv']
    files_to_process = []
    
    for pattern in file_patterns:
        files_to_process.extend(folder.glob(pattern))
    
    if not files_to_process:
        print("No Excel or CSV files found in the folder!")
        return
    
    print(f"Found {len(files_to_process)} file(s) to process")
    print(f"Output format: {output_format.upper()}")
    
    for file_path in files_to_process:
        print(f"\nProcessing: {file_path.name}")
        
        try:
            # Read file based on extension
            if file_path.suffix.lower() == '.csv':
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            
            total_rows = len(df)
            print(f"Total rows: {total_rows}")
            
            # Skip if file has 10k or fewer rows
            if total_rows <= 10000:
                print("File has 10,000 or fewer rows. No splitting needed.")
                continue
            
            # Calculate number of chunks
            num_chunks = math.ceil(total_rows / 10000)
            base_name = file_path.stem
            
            print(f"Creating {num_chunks} chunks...")
            
            # Create chunks
            for i in range(num_chunks):
                start_row = i * 10000
                end_row = min((i + 1) * 10000, total_rows)
                
                chunk = df.iloc[start_row:end_row].copy()
                
                # Create output filename based on format
                if output_format == 'xlsx':
                    output_name = f"{base_name}_chunk_{i+1:02d}.xlsx"
                    output_path = folder / output_name
                    chunk.to_excel(output_path, index=False)
                else: # csv
                    output_name = f"{base_name}_chunk_{i+1:02d}.csv"
                    output_path = folder / output_name
                    chunk.to_csv(output_path, index=False)
                
                print(f" Created: {output_name} ({len(chunk)} rows)")
            
            print(f"âœ“ Successfully split {file_path.name}")
            
        except Exception as e:
            print(f"âœ— Error processing {file_path.name}: {e}")

def split_with_custom_chunk_size(folder_path, chunk_size=10000, output_format='xlsx'):
    """
    Split Excel/CSV files with custom chunk size and output format
    
    Args:
        folder_path (str): Path to folder containing Excel/CSV files
        chunk_size (int): Number of rows per chunk (default: 10000)
        output_format (str): Output format - 'xlsx' or 'csv' (default: 'xlsx')
    """
    
    folder = Path(folder_path)
    
    if not folder.exists():
        print(f"Error: Folder '{folder_path}' does not exist!")
        return
    
    # Validate output format
    if output_format.lower() not in ['xlsx', 'csv']:
        print("Error: Output format must be 'xlsx' or 'csv'")
        return
    
    output_format = output_format.lower()
    
    # Find all Excel and CSV files
    file_patterns = ['*.xlsx', '*.xls', '*.csv']
    files_to_process = []
    
    for pattern in file_patterns:
        files_to_process.extend(folder.glob(pattern))
    
    if not files_to_process:
        print("No Excel or CSV files found in the folder!")
        return
    
    print(f"Found {len(files_to_process)} file(s) to process")
    print(f"Using chunk size: {chunk_size} rows")
    print(f"Output format: {output_format.upper()}")
    
    for file_path in files_to_process:
        print(f"\nProcessing: {file_path.name}")
        
        try:
            # Read file based on extension
            if file_path.suffix.lower() == '.csv':
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            
            total_rows = len(df)
            print(f"Total rows: {total_rows}")
            
            # Skip if file has fewer rows than chunk size
            if total_rows <= chunk_size:
                print(f"File has {total_rows} rows, which is <= chunk size ({chunk_size}). No splitting needed.")
                continue
            
            # Calculate number of chunks
            num_chunks = math.ceil(total_rows / chunk_size)
            base_name = file_path.stem
            
            print(f"Creating {num_chunks} chunks...")
            
            # Create chunks
            for i in range(num_chunks):
                start_row = i * chunk_size
                end_row = min((i + 1) * chunk_size, total_rows)
                
                chunk = df.iloc[start_row:end_row].copy()
                
                # Create output filename based on format
                if output_format == 'xlsx':
                    output_name = f"{base_name}_part_{i+1:02d}.xlsx"
                    output_path = folder / output_name
                    chunk.to_excel(output_path, index=False)
                else: # csv
                    output_name = f"{base_name}_part_{i+1:02d}.csv"
                    output_path = folder / output_name
                    chunk.to_csv(output_path, index=False)
                
                print(f" Created: {output_name} ({len(chunk)} rows)")
            
            print(f"âœ“ Successfully split {file_path.name}")
            
        except Exception as e:
            print(f"âœ— Error processing {file_path.name}: {e}")

def get_output_format_choice():
    """
    Get output format choice from user
    """
    while True:
        choice = input("Choose output format (1 for XLSX, 2 for CSV): ").strip()
        
        if choice == '1':
            return 'xlsx'
        elif choice == '2':
            return 'csv'
        else:
            print("Invalid choice. Please enter 1 for XLSX or 2 for CSV.")

def main():
    """
    Interactive main function with output format selection
    """
    print("Excel/CSV File Splitter")
    print("=" * 50)
    
    # Get folder path from user
    folder_path = input("Enter the folder path containing Excel/CSV files: ").strip()
    
    # Remove quotes if present
    folder_path = folder_path.strip('"').strip("'")
    
    # Get output format choice
    print("\nSelect output format:")
    print("1. XLSX (Excel format)")
    print("2. CSV (Comma-separated values)")
    output_format = get_output_format_choice()
    
    # Ask for chunk size
    chunk_input = input("\nEnter chunk size (default 10000): ").strip()
    
    try:
        chunk_size = int(chunk_input) if chunk_input else 10000
    except ValueError:
        print("Invalid chunk size. Using default value of 10000.")
        chunk_size = 10000
    
    print(f"\nSettings:")
    print(f"- Folder: {folder_path}")
    print(f"- Output format: {output_format.upper()}")
    print(f"- Chunk size: {chunk_size} rows")
    print("-" * 50)
    
    # Process files
    if chunk_size == 10000:
        split_excel_csv_to_10k_rows(folder_path, output_format)
    else:
        split_with_custom_chunk_size(folder_path, chunk_size, output_format)
    
    print("\nProcess completed!")

if __name__ == "__main__":
    # Uncomment one of the following options:
    
    # Option 1: Run interactively (with output format choice)
    main()
    
    # Option 2: Direct usage with XLSX output
    # split_excel_csv_to_10k_rows(r"C:\Users\YourName\Documents\DataFolder", "xlsx")
    
    # Option 3: Direct usage with CSV output
    # split_excel_csv_to_10k_rows(r"C:\Users\YourName\Documents\DataFolder", "csv")
    
    # Option 4: Custom chunk size with CSV output
    # split_with_custom_chunk_size(r"C:\Users\YourName\Documents\DataFolder", 5000, "csv")
