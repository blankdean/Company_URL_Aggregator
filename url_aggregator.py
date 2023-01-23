import sys
import openpyxl
from googlesearch import search

def fill_urls(excel_file):
    # Open the Excel spreadsheet
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb["Sheet1"]
    print(f"Opening Spreadsheet {excel_file}")
    modified = False
    # Loop through the rows in the spreadsheet
    for row in range(2, sheet.max_row + 1):
        # Get the company name and URL from the spreadsheet
        company_name = sheet.cell(row, 1).value
        company_url = sheet.cell(row, 2).value
        
        # If the company URL is not already filled in
        if not company_url:
            
            # Do a Google search for the company name
            # Docs: https://python-googlesearch.readthedocs.io/en/latest/
            for url in search(query=company_name, num=1, stop=1, pause=2):
                company_url = url
            
            print(company_name)
            print(f"Populating {excel_file} at row {row}: {company_url}")
            print('----------')

            # Update the company URL in the spreadsheet
            sheet.cell(row, 2).value = company_url

            modified = True

    if modified:
        # Save the changes to the spreadsheet
        print(f"Saving changes to {excel_file}")
        wb.save(excel_file)
    else:
        print(f"No changes to spreadsheet. All urls filled in {excel_file}.")

    wb.close()

if __name__=="__main__":
    excel_file = "companies.xlsx"

    try:
        fill_urls(excel_file)
    except Exception as e:
        print(f"An error occurred: {e}")

    input("Press Enter to exit...")
    sys.exit()

