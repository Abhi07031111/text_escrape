import os
import pandas as pd
import re

def extract_text_between_exec(file_path):
    """Extract text between occurrences of the word 'EXEC' in the given text file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            matches = re.findall(r'EXEC(.*?)EXEC', content, re.DOTALL)
            return '\n'.join(matches).strip() if matches else 'No EXEC block found'
    except Exception as e:
        return f'Error reading file: {str(e)}'

def process_excel(excel_path, folder_path):
    """Process the Excel file and update column B with extracted EXEC text."""
    try:
        df = pd.read_excel(excel_path, sheet_name=0, dtype=str)
        
        if 'A' not in df.columns:
            return "Column A not found in the Excel sheet"
        
        df['B'] = ''  # Initialize column B
        
        for index, program in df['A'].dropna().items():
            txt_file_path = os.path.join(folder_path, f"{program}.txt")
            
            if os.path.isfile(txt_file_path):
                df.at[index, 'B'] = extract_text_between_exec(txt_file_path)
            else:
                df.at[index, 'B'] = 'File not found'
        
        output_path = excel_path.replace('.xlsx', '_updated.xlsx')
        df.to_excel(output_path, index=False)
        return f"Processing complete. Updated file saved as: {output_path}"
    
    except Exception as e:
        return f"Error processing Excel file: {str(e)}"

# Example usage
excel_file_path = "programs.xlsx"  # Change this to the actual Excel file path
folder_path = "./txt_files"  # Change this to the actual folder path containing text files
result_message = process_excel(excel_file_path, folder_path)
print(result_message)
