"""
Utility functions for Excel processing
"""
import pandas as pd
import os
from datetime import datetime
from colorama import Fore, Style, init

init(autoreset=True)

class ExcelUtils:
    
    @staticmethod
    def read_excel_files(input_folder):
        """
        Read all Excel files from the input folder
        Returns: Dictionary of {filename: dataframe}
        """
        excel_files = {}
        supported_extensions = ('.xlsx', '.xls', '.xlsm')
        
        print(f"\n{Fore.CYAN}📂 Scanning folder: {input_folder}")
        
        if not os.path.exists(input_folder):
            raise FileNotFoundError(f"Input folder not found: {input_folder}")
        
        files_found = 0
        
        for file in os.listdir(input_folder):
            if file.endswith(supported_extensions):
                file_path = os.path.join(input_folder, file)
                try:
                    # Read all sheets
                    xl_file = pd.ExcelFile(file_path)
                    
                    # Combine all sheets into one dataframe
                    dataframes = []
                    for sheet_name in xl_file.sheet_names:
                        df = pd.read_excel(file_path, sheet_name=sheet_name)
                        df['Source_File'] = file
                        df['Sheet_Name'] = sheet_name
                        dataframes.append(df)
                    
                    combined_df = pd.concat(dataframes, ignore_index=True)
                    excel_files[file] = combined_df
                    files_found += 1
                    
                    print(f"{Fore.GREEN}✅ Loaded: {file} ({len(combined_df)} rows)")
                    
                except Exception as e:
                    print(f"{Fore.RED}❌ Error reading {file}: {str(e)}")
        
        if files_found == 0:
            print(f"{Fore.YELLOW}⚠️  No Excel files found!")
        
        return excel_files
    
    @staticmethod
    def generate_summary(dataframes_dict):
        """
        Generate summary statistics from all dataframes
        """
        print(f"\n{Fore.CYAN}📊 Generating summary statistics...")
        
        summary_data = {
            'Total Files Processed': len(dataframes_dict),
            'Total Rows': sum(len(df) for df in dataframes_dict.values()),
            'Total Columns': sum(len(df.columns) for df in dataframes_dict.values()) // len(dataframes_dict) if dataframes_dict else 0,
            'Report Generated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        # Combine all data for analysis
        if dataframes_dict:
            combined_data = pd.concat(dataframes_dict.values(), ignore_index=True)
            
            # Add more insights
            summary_data['Unique Files'] = combined_data['Source_File'].nunique() if 'Source_File' in combined_data.columns else 0
            
            return summary_data, combined_data
        
        return summary_data, pd.DataFrame()
    
    @staticmethod
    def style_excel_report(writer, sheet_name, df, summary):
        """
        Apply professional styling to Excel report
        """
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        cell_format = workbook.add_format({
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        date_format = workbook.add_format({
            'border': 1,
            'align': 'left',
            'num_format': 'yyyy-mm-dd'
        })
        
        # Apply header format
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Auto-adjust column widths
        for i, col in enumerate(df.columns):
            max_len = max(
                df[col].astype(str).map(len).max(),
                len(str(col))
            ) + 2
            worksheet.set_column(i, i, min(max_len, 50))
        
        # Add summary sheet
        summary_df = pd.DataFrame(list(summary.items()), columns=['Metric', 'Value'])
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        summary_sheet = writer.sheets['Summary']
        for col_num, value in enumerate(summary_df.columns.values):
            summary_sheet.write(0, col_num, value, header_format)
        
        return writer
