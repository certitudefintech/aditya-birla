"""
KYC Processing System
A GUI application for processing and verifying KYC (Know Your Customer) data from multiple sources.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd
import threading
import time
from datetime import datetime
import calendar
import customtkinter as ctk

# Constants
REQUIRED_INVESTOR_COLS = ['SCHEME', 'PURCHASEUNITS', 'TRDATE', 'DOB']
REQUIRED_RTA_COLS = ['SCHEME', 'ISIN', 'OPTDESC']
REQUIRED_AMFI_COLS = ['ISIN DIV PAYOUT/ISIN GROWTH', 'ISIN DIV REINVESTMENT', 'NET ASSET VALUE']

ADITYA_FUNDS = [
      'MIDCAP FUND',
    'CONTRA FUND',
    'MANUFACTURING FUND',
    'FLEXI CAP FUND',
    'LARGE & MID CAP FUND',
    'MULTICAP FUND',
    'SMALL CAP FUND',
    'LARGECAP FUND',
    'INFRASTRUCTURE FUND',
    'FINANCIAL SERVICES FUND',
    'ELSS TAX SAVER FUND',
    'LARGE CAP FUND',
    'MID CAP FUND',
    'FRONTLINE EQUITY FUND',
    'FOCUSED FUND',
    'EQUITY ADVANTAGE FUND',
    'MNC FUND',
    'MULTI-CAP FUND',
    'PURE VALUE FUND',
    'MANUFACTURING EQUITY FUND',
    'BANKING AND FINANCIAL SERVICES FUND',
    'DIVIDEND YIELD FUND',
    'DIGITAL INDIA FUND',
    'INDIA GENNEXT FUND',
    'INTERNATIONAL EQUITY FUND',
    'PHARMA & HEALTHCARE FUND',
    'BAL BHAVISHYA YOJNA',
    'RETIREMENT FUND - THE 30S PLAN',
    'RETIREMENT FUND - THE 40S PLAN',
    'RETIREMENT FUND - THE 50S PLAN',
    'PSU EQUITY FUND',
    'SPECIAL OPPORTUNITIES FUND',
    'ESG INTEGRATION STRATEGY FUND',
    'BUSINESS CYCLE FUND',
    'TRANSPORTATION AND LOGISTICS FUND',
    'QUANT FUND',
    'ADITYA CONGLOMERATE FUND',
    'ELSS TAX SAVER FUND'
]

class LoadingWindow:
    """A modal window that displays processing progress."""
    
    def __init__(self, parent):
        self.window = ctk.CTkToplevel(parent)
        self.window.title("Processing")
        self.window.geometry("400x200")
        self.window.transient(parent)
        self.window.grab_set()
        
        # Center the window
        x = parent.winfo_x() + (parent.winfo_width() - 400) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 200) // 2
        self.window.geometry(f"+{x}+{y}")
        
        self._create_widgets()
        
    def _create_widgets(self):
        """Create and configure the window widgets."""
        # Main frame
        self.frame = ctk.CTkFrame(self.window, fg_color="white")
        self.frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Loading text
        self.loading_label = ctk.CTkLabel(
            self.frame,
            text="Processing files...",
            font=("Segoe UI", 16, "bold"),
            text_color="#2c3e50"
        )
        self.loading_label.pack(pady=(20, 10))
        
        # Progress bar
        self.progress = ctk.CTkProgressBar(self.frame)
        self.progress.pack(fill="x", padx=20, pady=10)
        self.progress.set(0)
        
        # Status text
        self.status_label = ctk.CTkLabel(
            self.frame,
            text="Initializing...",
            font=("Segoe UI", 12),
            text_color="#7f8c8d"
        )
        self.status_label.pack(pady=10)
        
        # Make window modal
        self.window.protocol("WM_DELETE_WINDOW", lambda: None)
        
    def update_progress(self, value, status_text):
        """Update the progress bar and status text."""
        self.progress.set(value)
        self.status_label.configure(text=status_text)
        self.window.update()
        
    def close(self):
        """Close the loading window."""
        self.window.destroy()

class KYCProcessor:
    """Main application class for KYC processing."""
    
    def __init__(self):
        self.window = ctk.CTk()
        self.window.title("Miss_Selling Processing System")
        self.window.geometry("900x800")
        
        # Set theme
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        # Initialize file paths
        self.investor_file_path = None
        self.rta_file_path = None
        self.amfi_file_path = None
        
        self._create_gui()
        
    def _create_gui(self):
        """Create and configure the main GUI elements."""
        # Main frame
        self.main_frame = ctk.CTkFrame(self.window, fg_color="white")
        self.main_frame.pack(fill="both", expand=True, padx=30, pady=30)
        
        # Title
        self.title_label = ctk.CTkLabel(
            self.main_frame,
            text="Miss_Selling Processing System",
            font=("Segoe UI", 32, "bold"),
            text_color="#2c3e50"
        )
        self.title_label.pack(pady=(0, 30))
        
        # Create sections
        self._create_file_upload_section()
        self._create_process_button()
        self._create_status_section()
        
    def _create_file_upload_section(self):
        """Create the file upload section of the GUI."""
        self.files_frame = ctk.CTkFrame(self.main_frame, fg_color="#f8f9fa")
        self.files_frame.pack(fill="x", padx=30, pady=10)
        
        # Section title
        section_label = ctk.CTkLabel(
            self.files_frame,
            text="File Upload",
            font=("Segoe UI", 18, "bold"),
            text_color="#2c3e50"
        )
        section_label.pack(pady=(15, 20))
        
        # Create file upload rows
        self._create_file_upload_row("Investor Master KYC", "investor")
        self._create_file_upload_row("RTA Master", "rta")
        self._create_file_upload_row("AMFI Data", "amfi")
        
        # Add scheme name input field
        scheme_frame = ctk.CTkFrame(self.files_frame, fg_color="transparent")
        scheme_frame.pack(fill="x", pady=8)
        
        scheme_label = ctk.CTkLabel(
            scheme_frame,
            text="Underperforming Scheme Names (comma separated):",
            font=("Segoe UI", 14, "bold"),
            text_color="#2c3e50"
        )
        scheme_label.pack(side="left", padx=20, pady=10)
        
        self.scheme_entry = ctk.CTkEntry(
            scheme_frame,
            placeholder_text="Enter scheme names separated by commas",
            width=400,
            height=35,
            font=("Segoe UI", 12)
        )
        self.scheme_entry.pack(side="right", padx=20, pady=10)

        # Add Credit Risk Fund input field
        credit_risk_frame = ctk.CTkFrame(self.files_frame, fg_color="transparent")
        credit_risk_frame.pack(fill="x", pady=8)
        
        credit_risk_label = ctk.CTkLabel(
            credit_risk_frame,
            text="Credit Risk Fund Names (comma separated):",
            font=("Segoe UI", 14, "bold"),
            text_color="#2c3e50"
        )
        credit_risk_label.pack(side="left", padx=20, pady=10)
        
        self.credit_risk_entry = ctk.CTkEntry(
            credit_risk_frame,
            placeholder_text="Enter credit risk fund names separated by commas",
            width=400,
            height=35,
            font=("Segoe UI", 12)
        )
        self.credit_risk_entry.pack(side="right", padx=20, pady=10)
        
    def _create_file_upload_row(self, label_text, file_type):
        """Create a row for file upload with label and button."""
        frame = ctk.CTkFrame(self.files_frame, fg_color="transparent")
        frame.pack(fill="x", pady=8)
        
        label = ctk.CTkLabel(
            frame,
            text=label_text,
            font=("Segoe UI", 14, "bold"),
            text_color="#2c3e50"
        )
        label.pack(side="left", padx=20, pady=10)
        
        button = ctk.CTkButton(
            frame,
            text="Upload File",
            command=lambda: self._upload_file(file_type),
            fg_color="#3498db",
            hover_color="#2980b9",
            font=("Segoe UI", 12),
            width=120,
            height=35
        )
        button.pack(side="right", padx=20, pady=10)
        
        path_label = ctk.CTkLabel(
            frame,
            text="No file selected",
            text_color="#7f8c8d",
            font=("Segoe UI", 12)
        )
        path_label.pack(side="right", padx=20, pady=10)
        
        setattr(self, f"{file_type}_path_label", path_label)
        
    def _create_process_button(self):
        """Create the process button."""
        self.process_button = ctk.CTkButton(
            self.main_frame,
            text="Process Files",
            command=self._start_processing,
            fg_color="#27ae60",
            hover_color="#219a52",
            font=("Segoe UI", 14, "bold"),
            height=45,
            width=200
        )
        self.process_button.pack(pady=30)
        
    def _create_status_section(self):
        """Create the status section of the GUI."""
        self.status_frame = ctk.CTkFrame(self.main_frame, fg_color="#f8f9fa")
        self.status_frame.pack(fill="x", padx=30, pady=20)
        
        section_label = ctk.CTkLabel(
            self.status_frame,
            text="Status",
            font=("Segoe UI", 18, "bold"),
            text_color="#2c3e50"
        )
        section_label.pack(pady=(15, 10))
        
        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="Status: Ready to upload files",
            font=("Segoe UI", 12),
            text_color="#2c3e50"
        )
        self.status_label.pack(pady=10)
        
    def _upload_file(self, file_type):
        """Handle file upload for a specific file type."""
        file_path = filedialog.askopenfilename(
            title=f"Select {file_type.title()} File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            path_label = getattr(self, f"{file_type}_path_label")
            path_label.configure(
                text=os.path.basename(file_path),
                text_color="#27ae60"
            )
            setattr(self, f"{file_type}_file_path", file_path)
            self._update_status(f"{file_type.title()} file uploaded successfully")
    
    def _update_status(self, message):
        """Update the status label with a message."""
        self.status_label.configure(
            text=f"Status: {message}",
            text_color="#27ae60"
        )
    
    def _start_processing(self):
        """Start the file processing in a separate thread."""
        if not all([self.investor_file_path, self.rta_file_path, self.amfi_file_path]):
            messagebox.showerror("Error", "Please upload all three files before processing")
            return
        
        self.loading_window = LoadingWindow(self.window)
        thread = threading.Thread(target=self._process_files)
        thread.start()
    
    def _process_files(self):
        """Process the uploaded files and perform KYC checks."""
        try:
            self._read_files()
            self._perform_kyc_checks()
            self._save_results()
            
        except Exception as e:
            self.loading_window.close()
            self._update_status("Error processing files")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def _read_files(self):
        """Read and validate the uploaded files."""
        self.loading_window.update_progress(0.1, "Reading files...")
            
        # Read files
        self.investor_df = pd.read_excel(self.investor_file_path)
        self.rta_df = pd.read_excel(self.rta_file_path)
        self.amfi_df = pd.read_excel(self.amfi_file_path)
            
        # Convert column names to uppercase
        self.investor_df.columns = self.investor_df.columns.str.upper()
        self.rta_df.columns = self.rta_df.columns.str.upper()
        self.amfi_df.columns = self.amfi_df.columns.str.upper()
            
        # Validate required columns
        self._validate_required_columns()
        
        # Filter out specific schemes
        self.investor_df = self.investor_df[~self.investor_df['SCHEME'].str.contains('LF|ON|AF', case=False, na=False)]
            
    def _validate_required_columns(self):
        """Validate that all required columns are present in the files."""
        for col in REQUIRED_INVESTOR_COLS:
            if col not in self.investor_df.columns:
                raise ValueError(f"Column '{col}' not found in Investor Master KYC file")
            
        for col in REQUIRED_RTA_COLS:
            if col not in self.rta_df.columns:
                raise ValueError(f"Column '{col}' not found in RTA Master file")
            
        for col in REQUIRED_AMFI_COLS:
            if col not in self.amfi_df.columns:
                raise ValueError(f"Column '{col}' not found in AMFI NAV Data file")
            
    def _perform_kyc_checks(self):
        """Perform all KYC verification checks."""
        self.loading_window.update_progress(0.3, "Processing data...")
            
        # Convert dates
        self._convert_dates()
        
        # Calculate age
        self._calculate_age()
        
        # Add ISIN and OPTDESC
        self._add_isin_and_optdesc()
        
        # Add NAV values
        self._add_nav_values()
        
        # Calculate valuation
        self._calculate_valuation()
        
        # Perform specific checks
        self._perform_specific_checks()
        
    def _convert_dates(self):
        """Convert date columns to datetime format."""
        self.investor_df['TRDATE'] = pd.to_datetime(self.investor_df['TRDATE'], errors='coerce')
        self.investor_df['DOB'] = pd.to_datetime(self.investor_df['DOB'], errors='coerce')
            
    def _calculate_age(self):
        """Calculate age at the end of transaction month."""
        def calculate_age(row):
            try:
                if pd.isna(row['TRDATE']) or pd.isna(row['DOB']):
                    return None
                
                last_day = calendar.monthrange(row['TRDATE'].year, row['TRDATE'].month)[1]
                end_of_month = datetime(row['TRDATE'].year, row['TRDATE'].month, last_day)
                
                years = end_of_month.year - row['DOB'].year
                months = end_of_month.month - row['DOB'].month
                
                if (end_of_month.month, end_of_month.day) < (row['DOB'].month, row['DOB'].day):
                    years -= 1
                    months += 12
                
                return years + (months / 12)
            except Exception:
                return None
        
        self.investor_df['AGE AT TRANSACTION'] = self.investor_df.apply(calculate_age, axis=1)
            
    def _add_isin_and_optdesc(self):
        """Add ISIN and OPTDESC columns to investor dataframe."""
        self.loading_window.update_progress(0.7, "Adding ISIN and OPTDESC...")
        
        # Create mappings using only SCHEME
        rta_mapping_isin = self.rta_df.set_index('SCHEME')['ISIN'].to_dict()
        rta_mapping_optdesc = self.rta_df.set_index('SCHEME')['OPTDESC'].to_dict()
        rta_mapping_schemedesc = self.rta_df.set_index('ISIN')['SCHEMEDESC'].to_dict()

        # Add columns using only SCHEME
        self.investor_df['ISIN'] = self.investor_df['SCHEME'].map(rta_mapping_isin).fillna('Not Found')
        self.investor_df['OPTDESC'] = self.investor_df['SCHEME'].map(rta_mapping_optdesc).fillna('Not Found')
        self.investor_df['SCHEMEDESC'] = self.investor_df['ISIN'].map(rta_mapping_schemedesc).fillna('Not Found')
            
    def _add_nav_values(self):
        """Add NAV values based on OPTDESC."""
        self.loading_window.update_progress(0.8, "Adding NAV values...")
        
        # Create NAV mappings
        nav_mapping_growth = self.amfi_df.set_index('ISIN DIV PAYOUT/ISIN GROWTH')['NET ASSET VALUE'].to_dict()
        nav_mapping_reinvestment = self.amfi_df.set_index('ISIN DIV REINVESTMENT')['NET ASSET VALUE'].to_dict()
        
        def get_nav(row):
            if row['ISIN'] == 'Not Found':
                return 'Not Found'
            if 'REINVESTMENT' in str(row['OPTDESC']).upper():
                return nav_mapping_reinvestment.get(row['ISIN'], 'Not Found')
            return nav_mapping_growth.get(row['ISIN'], 'Not Found')
        
        self.investor_df['NAV'] = self.investor_df.apply(get_nav, axis=1)
            
    def _calculate_valuation(self):
        """Calculate valuation of investor."""
        self.loading_window.update_progress(0.85, "Calculating valuation...")
        
        def calculate_valuation(row):
            try:
                if row['NAV'] != 'Not Found' and pd.notna(row['PURCHASEUNITS']):
                    return float(row['PURCHASEUNITS']) * float(row['NAV'])
                return None
            except (ValueError, TypeError):
                return None
        
        self.investor_df['VALUATION OF INVESTOR'] = self.investor_df.apply(calculate_valuation, axis=1)
            
    def _perform_specific_checks(self):
        """Perform specific KYC verification checks."""
        # Check for small cap investments after age 80
        self._check_smallcap_after_80()
        # Check for ELSS investments after age 80
        self._check_elss_after_80()
        # Check investment against income slab (10x only)
        self._check_investment_10x_income()
        # Check high value investments for specific occupations
        self._check_high_value_occupation_investment()
        # Check mid cap investments after age 80
        self._check_midcap_after_80()
        # Check AOP/society investments
        self._check_aop_society_investment()
        # Check underperforming scheme investments
        self._check_underperforming_scheme()
        # Check credit risk fund investments
        self._check_credit_risk_fund()
        # Format dates for display
        self.investor_df['TRDATE'] = self.investor_df['TRDATE'].dt.strftime('%d-%m-%Y')
        self.investor_df['DOB'] = self.investor_df['DOB'].dt.strftime('%d-%m-%Y')
    
    def _check_smallcap_after_80(self):
        """Check for small cap investments by investors aged 80 or above."""
        def check_smallcap_after_80(row):
            try:
                if pd.notna(row.get('ARNNAME')) and 'DIRECT' not in str(row['ARNNAME']).upper():
                    if pd.notna(row.get('SCHEMEDESC')) and 'small cap' in str(row['SCHEMEDESC']).lower():
                        if pd.notna(row.get('AGE AT TRANSACTION')) and float(row['AGE AT TRANSACTION']) >= 80:
                            return 'Check'
                return 'OK'
            except Exception:
                return 'OK'
        
        self.investor_df['Investment in Small Cap Equity Schemes after Age 80'] = self.investor_df.apply(check_smallcap_after_80, axis=1)
            
    def _check_elss_after_80(self):
        """Check for ELSS investments by investors aged 80 or above."""
        def check_elss_after_80(row):
            try:
                if pd.notna(row.get('ARNNAME')) and 'DIRECT' not in str(row['ARNNAME']).upper():
                    if pd.notna(row.get('SCHEMEDESC')) and 'ELSS' in str(row['SCHEMEDESC']).upper():
                        if pd.notna(row.get('AGE AT TRANSACTION')) and float(row['AGE AT TRANSACTION']) >= 80:
                            return 'Check'
                return 'OK'
            except Exception:
                return 'OK'
        
        self.investor_df['Investment in ELSS Tax Saver Fund Equity Schemes after Age 80'] = self.investor_df.apply(check_elss_after_80, axis=1)
            
    def _check_investment_10x_income(self):
        """Check if investment amount is 10 times or more of the upper value of the income slab, ignoring 'DIRECT' in ARNNAME."""
        def convert_income_slab_to_number(income_slab):
            if pd.isna(income_slab):
                return None
            slab = str(income_slab)
            slab = slab.replace('–', '-').replace('—', '-').replace('−', '-')
            slab = '-'.join([s.strip() for s in slab.split('-')]).lower()
            if '1 lakh' in slab and '5 lakh' in slab:
                return 500000
            elif '5 lakh' in slab and '10 lakh' in slab:
                return 1000000
            elif '10 lakh' in slab and '25 lakh' in slab:
                return 2500000
            elif '25 lakh' in slab and '1 crore' in slab:
                return 10000000
            return None

        def check_investment_10x(row):
            arn = str(row.get('ARNNAME', '')).strip().upper()
            if arn == 'DIRECT':
                return 'OK'
            income_value = convert_income_slab_to_number(row.get('INCOMESLAB'))
            investment = row.get('VALUATION OF INVESTOR')
            if pd.isna(income_value) or pd.isna(investment):
                return 'OK'
            try:
                investment = float(investment)
                if investment >= 10 * income_value:
                    return 'Check'
                return 'OK'
            except Exception:
                return 'OK'
        
        self.investor_df['Investment 10x the given INCOMESLAB'] = self.investor_df.apply(check_investment_10x, axis=1)
            
    def _check_high_value_occupation_investment(self):
        """Check for investments higher than 50 lakhs by HOUSEHOLD, FARMER, LABOUR."""
        def check_high_value_occupation_investment(row):
            try:
                if pd.notna(row.get('ARNNAME')) and 'DIRECT' in str(row['ARNNAME']).upper():
                    return 'OK'
                
                occupation = str(row.get('OCCUPATION_DESCRIPTION', '')).upper()
                if not any(occ in occupation for occ in ['HOUSEHOLD', 'FARMER', 'LABOUR']):
                    return 'OK'
                
                investment_value = row.get('VALUATION OF INVESTOR')
                if pd.isna(investment_value) or investment_value == 'Not Found':
                    return 'OK'
                
                if float(investment_value) >= 5000000:
                    return 'Check'
                
                return 'OK'
            except Exception:
                return 'OK'
        
        self.investor_df['Investment of 50 lakhs or higher HOUSEHOLD,FARMER,LABOUR'] = self.investor_df.apply(check_high_value_occupation_investment, axis=1)
            
    def _check_midcap_after_80(self):
        """Check for Mid Cap investments by investors aged 80 or above."""
        def check_midcap_after_80(row):
            try:
                if pd.notna(row.get('ARNNAME')) and 'DIRECT' not in str(row['ARNNAME']).upper():
                    if pd.notna(row.get('SCHEMEDESC')):
                        scheme_desc = str(row['SCHEMEDESC']).lower()
                        if 'mid cap' in scheme_desc and 'large & mid cap' not in scheme_desc:
                            if pd.notna(row.get('AGE AT TRANSACTION')) and float(row['AGE AT TRANSACTION']) >= 80:
                                return 'Check'
                return 'OK'
            except Exception:
                return 'OK'
        
        self.investor_df['Investment in Mid Cap Fund Equity Schemes after Age 80'] = self.investor_df.apply(check_midcap_after_80, axis=1)
            
    def _check_aop_society_investment(self):
        """Check for AOP/society investments in specific Invesco funds."""
        def check_aop_society_investment(row):
            try:
                if pd.notna(row.get('ARNNAME')) and 'DIRECT' not in str(row['ARNNAME']).upper():
                    if pd.notna(row.get('STATDESC')):
                        stat_desc = str(row['STATDESC']).upper()
                        if 'TRUST' in stat_desc or 'SOCIETY' in stat_desc or 'CLUB' in stat_desc:
                            if pd.notna(row.get('SCHEMEDESC')):
                                scheme_desc = str(row['SCHEMEDESC']).upper()
                                if any(fund in scheme_desc for fund in ADITYA_FUNDS):
                                    return 'Check'
                return 'OK'
            except Exception:
                return 'OK'
        
        self.investor_df['AOP/society making investments in equity'] = self.investor_df.apply(check_aop_society_investment, axis=1)
            
    def _check_underperforming_scheme(self):
        """Check for investments in underperforming schemes by individual investors with valuation >= 10 lakhs."""
        scheme_names = [name.strip().upper() for name in self.scheme_entry.get().split(',') if name.strip()]
        if not scheme_names:
            self.investor_df['Allocation to Underperfoming scheme / execption scheme - Greater than 10 lakhs - Individual Investor'] = 'OK'
            return
            
        def check_underperforming_scheme(row):
            try:
                if pd.notna(row.get('SCHEMEDESC')):
                    scheme_desc = str(row['SCHEMEDESC']).upper()
                    # Check if any of the provided scheme names match
                    if any(scheme_name in scheme_desc for scheme_name in scheme_names):
                        if pd.notna(row.get('STATDESC')) and 'INDIVIDUAL' in str(row['STATDESC']).upper():
                            if pd.notna(row.get('VALUATION OF INVESTOR')):
                                try:
                                    valuation = float(row['VALUATION OF INVESTOR'])
                                    if valuation >= 1000000:  # 10 lakhs
                                        return 'Check'
                                except (ValueError, TypeError):
                                    pass
                return 'OK'
            except Exception:
                return 'OK'
        
        self.investor_df['Allocation to Underperfoming scheme / execption scheme - Greater than 10 lakhs - Individual Investor'] = self.investor_df.apply(check_underperforming_scheme, axis=1)
            
    def _check_credit_risk_fund(self):
        """Check for investments in credit risk funds by individual investors aged 80 or above."""
        fund_names = [name.strip().upper() for name in self.credit_risk_entry.get().split(',') if name.strip()]
        if not fund_names:
            self.investor_df['Credit Risk Fund Above 80 - INDIVIDUAL INVESTOR'] = 'OK'
            return
            
        def check_credit_risk_fund(row):
            try:
                if pd.notna(row.get('SCHEMEDESC')):
                    scheme_desc = str(row['SCHEMEDESC']).upper()
                    # Check if any of the provided fund names match
                    if any(fund_name in scheme_desc for fund_name in fund_names):
                        if pd.notna(row.get('STATDESC')) and 'INDIVIDUAL' in str(row['STATDESC']).upper():
                            if pd.notna(row.get('AGE AT TRANSACTION')):
                                try:
                                    age = float(row['AGE AT TRANSACTION'])
                                    if age >= 80:
                                        return 'Check'
                                except (ValueError, TypeError):
                                    pass
                return 'OK'
            except Exception:
                return 'OK'
        
        self.investor_df['Credit Risk Fund Above 80 - INDIVIDUAL INVESTOR'] = self.investor_df.apply(check_credit_risk_fund, axis=1)
            
    def _save_results(self):
        """Save a 'Main Data' sheet with all information, and a 'Checks Info' sheet with grouped checks."""
        self.loading_window.update_progress(0.9, "Preparing to save...")
        
        default_filename = "Processed_MISS_Selling_KYC.xlsx"
        output_path = filedialog.asksaveasfilename(
            title="Save Processed File",
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        
        if output_path:
            self.loading_window.update_progress(0.95, "Saving file...")
            import openpyxl
            from openpyxl import Workbook
            from openpyxl.styles import PatternFill, Font, Alignment
            from openpyxl.utils import get_column_letter

            # Find the columns from ACNO to VALUATION OF INVESTOR (inclusive)
            all_columns = list(self.investor_df.columns)
            try:
                acno_idx = all_columns.index('ACNO')
                val_idx = all_columns.index('VALUATION OF INVESTOR')
                acno_to_val_cols = all_columns[acno_idx:val_idx+1]
            except ValueError:
                # Fallback: if not found, use all columns
                acno_to_val_cols = all_columns

            check_blocks = [
                {
                    'title': 'Investment in Small Cap Equity Schemes after Age 80',
                    'check_col': 'Investment in Small Cap Equity Schemes after Age 80',
                    'color': 'FFC7CE'  # Light Red
                },
                {
                    'title': 'Investment in ELSS Tax Saver Fund Equity Schemes after Age 80',
                    'check_col': 'Investment in ELSS Tax Saver Fund Equity Schemes after Age 80',
                    'color': 'FFEB9C'  # Light Yellow
                },
                {
                    'title': 'Investment 10x the given INCOMESLAB',
                    'check_col': 'Investment 10x the given INCOMESLAB',
                    'color': 'C6EFCE'  # Light Green
                },
                {
                    'title': 'Investment of 50 lakhs or higher HOUSEHOLD,FARMER,LABOUR',
                    'check_col': 'Investment of 50 lakhs or higher HOUSEHOLD,FARMER,LABOUR',
                    'color': 'B4C6E7'  # Light Blue
                },
                {
                    'title': 'Investment in Mid Cap Fund Equity Schemes after Age 80',
                    'check_col': 'Investment in Mid Cap Fund Equity Schemes after Age 80',
                    'color': 'F4B084'  # Light Orange
                },
                {
                    'title': 'AOP/society making investments in equity',
                    'check_col': 'AOP/society making investments in equity',
                    'color': 'D9E1F2'  # Light Purple
                },
                {
                    'title': 'Allocation to Underperfoming scheme / execption scheme - Greater than 10 lakhs - Individual Investor',
                    'check_col': 'Allocation to Underperfoming scheme / execption scheme - Greater than 10 lakhs - Individual Investor',
                    'color': 'E2EFDA'  # Light Mint
                },
                {
                    'title': 'Credit Risk Fund Above 80 - INDIVIDUAL INVESTOR',
                    'check_col': 'Credit Risk Fund Above 80 - INDIVIDUAL INVESTOR',
                    'color': 'FFD9D9'  # Light Pink
                }
            ]

            wb = Workbook()
            ws_main = wb.active
            ws_main.title = 'Main Data'
            ws_checks = wb.create_sheet('Checks Info')

            # Write Main Data sheet (all rows, all columns)
            for col_idx, col_name in enumerate(self.investor_df.columns, 1):
                cell = ws_main.cell(row=1, column=col_idx, value=col_name)
                cell.fill = PatternFill(start_color='305496', end_color='305496', fill_type='solid')
                cell.font = Font(color='FFFFFF', bold=True, size=11)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            for row_idx, row in enumerate(self.investor_df.itertuples(index=False), 2):
                for col_idx, value in enumerate(row, 1):
                    cell = ws_main.cell(row=row_idx, column=col_idx, value=value)
            for col_idx in range(1, ws_main.max_column + 1):
                col_letter = get_column_letter(col_idx)
                ws_main.column_dimensions[col_letter].width = 20

            # Write Checks Info sheet (grouped by check, only issues, all columns from ACNO to VALUATION OF INVESTOR + check col)
            row_cursor = 1
            for block in check_blocks:
                check_col = block['check_col']
                # Columns for this block: all from ACNO to VALUATION OF INVESTOR, plus the check column (if not already included)
                block_columns = acno_to_val_cols.copy()
                if check_col not in block_columns:
                    block_columns.append(check_col)
                check_df = self.investor_df[self.investor_df[check_col] == 'Check']
                if check_df.empty:
                    continue
                ws_checks.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=len(block_columns))
                title_cell = ws_checks.cell(row=row_cursor, column=1, value=block['title'])
                title_cell.font = Font(bold=True, size=13, color='305496')
                title_cell.alignment = Alignment(horizontal='center', vertical='center')
                row_cursor += 1
                # Header row
                for col_idx, col_name in enumerate(block_columns, 1):
                    cell = ws_checks.cell(row=row_cursor, column=col_idx, value=col_name)
                    cell.fill = PatternFill(start_color='305496', end_color='305496', fill_type='solid')
                    cell.font = Font(color='FFFFFF', bold=True, size=11)
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                row_cursor += 1
                # Data rows
                for _, row in check_df[block_columns].iterrows():
                    for col_idx, value in enumerate(row, 1):
                        cell = ws_checks.cell(row=row_cursor, column=col_idx, value=value)
                        if col_idx == len(block_columns) and value == 'Check':
                            cell.fill = PatternFill(start_color=block['color'], end_color=block['color'], fill_type='solid')
                            cell.font = Font(bold=True)
                    row_cursor += 1
                row_cursor += 1
            for col_idx in range(1, ws_checks.max_column + 1):
                col_letter = get_column_letter(col_idx)
                ws_checks.column_dimensions[col_letter].width = 20

            wb.save(output_path)
            time.sleep(0.5)
            self.loading_window.close()
            self._update_status("Files processed and formatted successfully!")
            messagebox.showinfo(
                "Success", 
                f"Files have been processed and formatted successfully!\nOutput saved to: {output_path}"
            )
        else:
            self.loading_window.close()
            self._update_status("File save cancelled")
            messagebox.showinfo("Cancelled", "File save was cancelled")
    
    def run(self):
        """Start the application."""
        self.window.mainloop()

if __name__ == "__main__":
    app = KYCProcessor()
    app.run()
