"""
Excel Report Generator - Pro Edition
Fixed Version - No 'float' has no len() Error
"""
import os
import sys
import pandas as pd
import xlsxwriter
from datetime import datetime
from colorama import init, Fore, Style
from config import Config

# Initialize colorama
init(autoreset=True)

class ExcelReportGenerator:
    def __init__(self):
        self.dataframes = {}
        self.summary_data = {}
        self.input_folder = Config.INPUT_FOLDER
        self.output_folder = Config.OUTPUT_FOLDER
        
    def load_all_files(self):
        """
        Load all Excel files safely
        """
        print(f"\n📂 Scanning folder: {self.input_folder}")
        
        # Check if folder exists
        if not os.path.exists(self.input_folder):
            raise Exception(f"❌ Input folder '{self.input_folder}' not found!")
        
        files = [f for f in os.listdir(self.input_folder) 
                if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]
        
        if not files:
            raise Exception(f"❌ No Excel files found in '{self.input_folder}'!")
        
        for file in files:
            try:
                file_path = os.path.join(self.input_folder, file)
                
                # Skip temp files
                if '~$' in file:
                    continue
                
                # Read Excel with error handling
                df = pd.read_excel(file_path, engine='openpyxl')
                
                # Only accept valid DataFrames
                if isinstance(df, pd.DataFrame) and not df.empty:
                    self.dataframes[file] = df
                    print(f"  ✅ Loaded: {file} ({len(df)} rows)")
                else:
                    print(f"  ⚠️ Skipped: {file} (Empty or invalid)")
                    
            except Exception as e:
                print(f"  ❌ Failed: {file} - {str(e)}")
        
        if not self.dataframes:
            raise Exception("❌ No valid Excel files could be loaded!")
        
        print(f"\n📊 Total files loaded: {len(self.dataframes)}")

    def generate_summary(self):
        """
        Generate summary statistics - FIXED for float error
        """
        try:
            valid_dfs = {}
            
            # Filter only valid DataFrames
            for name, df in self.dataframes.items():
                if isinstance(df, pd.DataFrame):
                    valid_dfs[name] = df
                else:
                    print(f"  ⚠️ Skipping {name}: Not a DataFrame")
            
            if not valid_dfs:
                return {'Total Files': 0, 'Total Rows': 0, 'Files': []}
            
            # Calculate totals safely
            total_rows = 0
            for df in valid_dfs.values():
                try:
                    # Ensure we can get length
                    if hasattr(df, '__len__'):
                        total_rows += len(df)
                except:
                    pass
            
            self.summary_data = {
                'Total Files': len(valid_dfs),
                'Total Rows': total_rows,
                'Files': list(valid_dfs.keys())
            }
            
            print(f"\n📈 Summary:")
            print(f"   Files: {self.summary_data['Total Files']}")
            print(f"   Rows: {self.summary_data['Total Rows']}")
            
            return self.summary_data
            
        except Exception as e:
            print(f"⚠️ Summary warning: {e}")
            return {'Total Files': 0, 'Total Rows': 0, 'Files': []}

    def analyze_data(self, df, sheet_name):
        """
        Analyze individual dataframe
        """
        try:
            analysis = {
                'Sheet Name': sheet_name,
                'Rows': len(df) if isinstance(df, pd.DataFrame) else 0,
                'Columns': len(df.columns) if isinstance(df, pd.DataFrame) else 0,
                'Column Names': list(df.columns) if isinstance(df, pd.DataFrame) else [],
                'Numeric Columns': [],
                'Text Columns': []
            }
            
            if isinstance(df, pd.DataFrame):
                for col in df.columns:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        analysis['Numeric Columns'].append(col)
                    else:
                        analysis['Text Columns'].append(col)
            
            return analysis
            
        except Exception as e:
            print(f"⚠️ Analysis error for {sheet_name}: {e}")
            return {
                'Sheet Name': sheet_name,
                'Rows': 0, 'Columns': 0,
                'Column Names': [],
                'Numeric Columns': [],
                'Text Columns': []
            }

    def create_excel_report(self, recipient_email=None):
        """
        Create consolidated Excel report with formatting
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(self.output_folder, f"Consolidated_Report_{timestamp}.xlsx")
        
        try:
            # Create Excel writer
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # Define formats
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#4472C4',
                    'font_color': 'white',
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter'
                })
                
                title_format = workbook.add_format({
                    'bold': True,
                    'font_size': 14,
                    'bg_color': '#70AD47',
                    'font_color': 'white'
                })
                
                summary_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#FFC000',
                    'border': 1
                })
                
                # 1. Create Summary Sheet
                summary_data = [
                    ['Excel Report Generator - Summary'],
                    [''],
                    ['Generated:', datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                    ['Company:', Config.COMPANY_NAME],
                    [''],
                    ['Statistics'],
                    ['Total Files Processed:', self.summary_data.get('Total Files', 0)],
                    ['Total Rows:', self.summary_data.get('Total Rows', 0)],
                    [''],
                    ['Files Included:']
                ]
                
                # Add file list
                for file in self.summary_data.get('Files', []):
                    summary_data.append([file])
                
                df_summary = pd.DataFrame(summary_data)
                df_summary.to_excel(writer, sheet_name='Summary', index=False, header=False)
                
                # Format summary sheet
                worksheet_summary = writer.sheets['Summary']
                worksheet_summary.set_column('A:A', 30)
                worksheet_summary.write('A1', summary_data[0][0], title_format)
                worksheet_summary.write('A6', 'Statistics', header_format)
                
                # 2. Add individual sheets for each file
                for file_name, df in self.dataframes.items():
                    try:
                        # Clean sheet name (Excel limit: 31 chars)
                        sheet_name = file_name.replace('.xlsx', '').replace('.xls', '')[:31]
                        
                        # Write data
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # Format header
                        worksheet = writer.sheets[sheet_name]
                        for col_num, value in enumerate(df.columns.values):
                            worksheet.write(0, col_num, value, header_format)
                        
                        # Auto-adjust columns
                        for i, col in enumerate(df.columns):
                            max_len = max(
                                df[col].astype(str).map(len).max() if not df[col].empty else 0,
                                len(str(col))
                            )
                            worksheet.set_column(i, i, min(max_len + 2, 50))
                            
                    except Exception as e:
                        print(f"  ⚠️ Error writing sheet for {file_name}: {e}")
                
                # 3. Add Analysis Sheet
                analysis_rows = []
                for file_name, df in self.dataframes.items():
                    analysis = self.analyze_data(df, file_name)
                    analysis_rows.append([
                        analysis['Sheet Name'],
                        analysis['Rows'],
                        analysis['Columns'],
                        len(analysis['Numeric Columns']),
                        len(analysis['Text Columns'])
                    ])
                
                if analysis_rows:
                    df_analysis = pd.DataFrame(
                        analysis_rows,
                        columns=['File Name', 'Rows', 'Columns', 'Numeric Cols', 'Text Cols']
                    )
                    df_analysis.to_excel(writer, sheet_name='Analysis', index=False)
                    
                    worksheet_analysis = writer.sheets['Analysis']
                    for col_num, value in enumerate(df_analysis.columns.values):
                        worksheet_analysis.write(0, col_num, value, header_format)
            
            print(f"\n{'='*60}")
            print(f"✅ SUCCESS! Report generated:")
            print(f"📁 Location: {output_file}")
            print(f"{'='*60}")
            
            return output_file
            
        except Exception as e:
            raise Exception(f"Excel creation failed: {e}")

    def send_email(self, file_path, recipient_email):
        """
        Send email with attachment (Optional)
        """
        if not recipient_email:
            print("📧 No email sent (no recipient provided)")
            return
        
        if not Config.SENDER_EMAIL or not Config.SENDER_PASSWORD:
            print("⚠️ Email not configured in .env file")
            print("💡 Report saved locally instead")
            return
        
        try:
            import smtplib
            from email.mime.multipart import MIMEMultipart
            from email.mime.text import MIMEText
            from email.mime.base import MIMEBase
            from email import encoders
            
            msg = MIMEMultipart()
            msg['From'] = Config.SENDER_EMAIL
            msg['To'] = recipient_email
            msg['Subject'] = f"Consolidated Report - {datetime.now().strftime('%Y-%m-%d')}"
            
            body = f"""
            Hello,
            
            Please find attached the consolidated Excel report.
            
            Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            Company: {Config.COMPANY_NAME}
            
            Best regards,
            Excel Report Generator
            """
            
            msg.attach(MIMEText(body, 'plain'))
            
            # Attach file
            with open(file_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {os.path.basename(file_path)}'
                )
                msg.attach(part)
            
            # Send email
            server = smtplib.SMTP(Config.SMTP_SERVER, Config.SMTP_PORT)
            server.starttls()
            server.login(Config.SENDER_EMAIL, Config.SENDER_PASSWORD)
            server.send_message(msg)
            server.quit()
            
            print(f"✅ Email sent to: {recipient_email}")
            
        except Exception as e:
            print(f"⚠️ Email failed: {e}")
            print(f"💡 Report saved locally: {file_path}")

def main():
    print("\n🚀 Excel Report Generator - Business Automation Tool")
    print("-" * 50)
    
    try:
        # Initialize generator
        generator = ExcelReportGenerator()
        
        # Print header
        print(f"\n{'='*60}")
        print(f"   📊 EXCEL REPORT GENERATOR - PRO EDITION")
        print(f"{'='*60}")
        
        # Ask for email
        recipient = input("\n📧 Enter recipient email (or press Enter to skip): ").strip()
        
        # Load files
        generator.load_all_files()
        
        # Generate summary
        print(f"\n📊 Generating summary statistics...")
        generator.generate_summary()
        
        # Create report
        print(f"\n📁 Creating consolidated Excel report...")
        output_file = generator.create_excel_report(recipient)
        
        # Send email if provided
        if recipient and output_file:
            generator.send_email(output_file, recipient)
        
        # Final message
        print(f"\n🎉 Process completed successfully!")
        print(f"💾 Find your report in: {Config.OUTPUT_FOLDER}\\")
        
        # Open output folder
        try:
            os.startfile(Config.OUTPUT_FOLDER)
        except:
            pass
        
    except KeyboardInterrupt:
        print(f"\n\n⚠️ Process interrupted by user")
        sys.exit(0)
    except Exception as e:
        print(f"\n{'='*60}")
        print(f"❌ ERROR: {e}")
        print(f"{'='*60}")
        sys.exit(1)
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
