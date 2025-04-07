import logging

from .utils import SalarySlipsManager


class Automation():
    def __init__(self):
        print("\n============ Salary Slips Manager Started ==========\n")

        self.helper = SalarySlipsManager()
        self.helper.load_config()
        self.helper.connect_to_gmail_imap()
        self.email_ids = self.helper.search_emails("Payslip for")

    def run(self):
        print("\n1. Extract Salary Details\n2. Extract Salary Slips\n\n")
        option = int(input("Enter your choice: "))

        match option:
            case 1:
                print("Extracting the details from the slips...")
                self.extract_salary_details()
            case 2:
                print("Extracting the salary slips...")
                self.extract_salary_slips()
            case default:
                print("Invalid option selected.")
        print("\n============ Salary Slips Manager Ended ==========\n\n")
        exit(0)
        
    def extract_salary_details(self):
        # Search for emails with subject containing "Payslip for the month of"
        if self.email_ids:
            for email_id in self.email_ids:
                email_msg = self.helper.fetch_email(email_id)
                if email_msg:
                    # Extract PDF attachment
                    pdf_data = self.helper.extract_pdf_from_attachment(email_msg)
                    if pdf_data:
                        # Read PDF content
                        pdf_text = self.helper.read_pdf_content(pdf_data)
                        if pdf_text:
                            # Extract all values from the PDF
                            extracted_values = self.helper.extract_all_values(pdf_text)
                            if any(extracted_values):  # Only add to the list if there's any extracted value
                                self.helper.salary_data.append(extracted_values)
                                print(f"Extracted Values: {extracted_values}")
                            else:
                                logging.warning("No relevant data found in the PDF.")
                        else:
                            logging.warning("Failed to read PDF content.")
        else:
            print("No emails found with the specified subject.")

        # Save the extracted data to an Excel file
        self.helper.save_to_excel()

    def extract_salary_slips(self):
        if self.email_ids:
            l = len(self.email_ids)
            for idx, email_id in enumerate(self.email_ids):
                email_msg = self.helper.fetch_email(email_id)
                if email_msg:
                    # Extract PDF attachment
                    month_year = "Payslip for the month of "
                    month_year += f"{(10 + idx) % 12 or 12:02d}-{2023 + (10 + idx - 1) // 12}"
                    print(f"Extracting {month_year}")
                    self.helper.save_slip(email_msg, month_year)
