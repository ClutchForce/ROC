import openpyxl

# Function to open and print the content of an Excel file
def print_excel_content(file_path, sheet_name):
    try:
        # Load the workbook
        workbook = openpyxl.load_workbook(file_path)
        
        # Select the specific sheet by name
        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
        else:
            print(f"Error: The sheet '{sheet_name}' does not exist in the workbook.")
            return
        
        # Iterate through rows and print each cell's value
        for row in sheet.iter_rows(values_only=True):
            print(row)
    
    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

# Specify the path to your Excel file
excel_file_path = "../Collections and Payments 2026.xlsx"  # Replace with the actual file path

# Call the function
print_excel_content(excel_file_path, sheet_name="Jan 2026")  # Replace "Sheet1" with the actual sheet name if needed


# 1010	Cash - BMO
# 1020	Cash - DOINA account
# 1030	Cash - Ladies Account
# 1040	Cash - Repairs account
# 1050	Cash - Other Account
# 1060	Cash - Paypal
# 1200	Accounts Receivable - Other
# 1300	Prepaids
# 1800	Investments
# 1900	Land and Building 
# 1910	Church capital renovations
# 1920	Casa Romana capital renovations
# 1930	Small Hall capital renovation
# 1940	Parish House capital renovations
# 1950	Farm capital renovations
# 2000	Accounts Payable
# 2100	Accrued Liabilities
# 2200	Payroll source deductions
# 2300	Loans
# 2400	Deposits payable
# 2500	Loan "DOINA"
# 3000	Retained Earnings
# 4010	Collection plate 
# 4020	Candles
# 4030	Donations
# 4040	Donations - Mostenire
# 4050	Donations - Repairs fund
# 4060	Religious services 
# 4070	Membership
# 4080	News Paper, Calendars, Books
# 4090	Rent (Hall, House, Farm)
# 4100	Dinners 
# 4200	Heritage Programs ("Doina")
# 4300	Other income
# 5100	Salaries
# 5110	Payroll source deductions
# 5120	Employee benefits
# 5130	Religious services
# 5140	Cantor, Cashier, Choir Director
# 5200	Candles
# 5300	Dinners 
# 5400	Diocese contribution
# 5410	Diocese - Mission Fund
# 5420	Diocese - Travel expenses
# 5500	Utilities - Hydro, Gas, Water
# 5610	School expenses
# 5620	"Doina" Dance Ensamble
# 5710	Telephone, Internet
# 5720	Insurance
# 5730	Property taxes
# 5740	Church and other supplies
# 5750	Repairs and maintenance
# 5760	Repairs - Repairs Fund
# 5800	Professional legal services
# 5910	Bank charges
# 5920	Office expenses
# 5930	Travel 
# 5940	Donations
# 5950	Miscellaneous
# 5999	Interest expense
