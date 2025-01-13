import imaplib
import json
import email
import logging
from email.header import decode_header
import PyPDF2
from io import BytesIO
import pandas as pd
import re

logging.basicConfig(
    level=logging.INFO,
    format="ExecTime %(asctime)s - LevelName %(levelname)s - FileName %(filename)s - LineNo %(lineno)d - FunctionName %(funcName)s \n %(message)s",
)


class SalarySlipsManager:
    def __init__(self):
        self.email = None
        self.password = None
        self.mail = None
        self.salary_data = []  # List to hold the earnings and tax data

    def load_config(self):
        try:
            with open("config.json", "r") as f:
                config = json.load(f)
                self.email = config["user"]
                self.password = config["pass"]
        except Exception as e:
            logging.error("Failed to load config: {}".format(e))

    def connect_to_gmail_imap(self):
        imap_url = "imap.gmail.com"
        try:
            self.mail = imaplib.IMAP4_SSL(imap_url)
            self.mail.login(self.email, self.password)
            self.mail.select("inbox")  # Connect to the inbox.
            logging.info("Connected to Gmail IMAP successfully.")
        except Exception as e:
            logging.error(f"Connection failed: {e}")

    def search_emails(self, search_text):
        try:
            # Search for emails with specific subject text
            status, messages = self.mail.search(None, f'(SUBJECT "{search_text}")')
            if status != "OK":
                logging.error("No messages found with the specified subject.")
                return []
            message_ids = messages[0].split()
            return message_ids
        except Exception as e:
            logging.error(f"Failed to search emails: {e}")
            return []

    def fetch_email(self, email_id):
        try:
            status, msg_data = self.mail.fetch(email_id, "(RFC822)")
            if status != "OK":
                logging.error(f"Failed to fetch email ID {email_id}")
                return None
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    return msg
        except Exception as e:
            logging.error(f"Failed to fetch email: {e}")
            return None

    def extract_pdf_from_attachment(self, msg):
        try:
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                # Only process PDF attachments
                if content_type == "application/pdf" and "attachment" in content_disposition:
                    file_data = part.get_payload(decode=True)
                    return file_data
            return None
        except Exception as e:
            logging.error(f"Error extracting PDF: {e}")
            return None

    def read_pdf_content(self, pdf_data):
        try:
            # Create a PDF reader object
            pdf_reader = PyPDF2.PdfReader(BytesIO(pdf_data))
            text = ""
            # Read all the pages of the PDF
            for page in pdf_reader.pages:
                text += page.extract_text()
            return text
        except Exception as e:
            logging.error(f"Error reading PDF: {e}")
            return None

    def extract_all_values(self, pdf_text):
        try:
            # Regex patterns for all values
            total_earnings_match = re.search(r"Total Earnings\s+(\d{1,3}(?:,\d{3})*\.\d{2})", pdf_text)
            overtime_match = re.search(r"Overtime\s+(\d{1,3}(?:,\d{3})*\.\d{2})", pdf_text)
            commision_bonus_match = re.search(r"Commission\/Bonus\s+(\d{1,3}(?:,\d{3})*\.\d{2})", pdf_text)

            total_deductions_match = re.search(r"Total Deductions\s+(\d{1,3}(?:,\d{3})*\.\d{2})", pdf_text)
            pf_match = re.search(r"ProvidentFundContributionEmployee\s+(\d{1,3}(?:,\d{3})*\.\d{2})", pdf_text)
            eobi_match = re.search(r"EOBIContributionEmployee\s+(\d{1,3}(?:,\d{3})*\.\d{2})", pdf_text)
            tax_match = re.search(r"PayrollTax\s+(\d{1,3}(?:,\d{3})*\.\d{2})", pdf_text)

            take_home_match = re.search(r"Take Home Pay\s+(\d{1,3}(?:,\d{3})*\.\d{2})", pdf_text)

            # Extract the values or None if not found
            total_earnings = total_earnings_match.group(1) if total_earnings_match else 0 
            overtime = overtime_match.group(1) if overtime_match else 0
            commision_with_bonus = commision_bonus_match.group(1) if commision_bonus_match else 0
            # basic_with_medical = float(total_earnings) - float(commision_with_bonus) - float(overtime)  

            total_deductions = total_deductions_match.group(1) if total_deductions_match else 0
            provident_fund = pf_match.group(1) if pf_match else 0
            payroll_tax = tax_match.group(1) if tax_match else 0
            eobi = eobi_match.group(1) if eobi_match else 0

            take_home = take_home_match.group(1) if take_home_match else 0

            return [
                total_earnings,
                # basic_with_medical,
                overtime,
                commision_with_bonus,
                total_deductions,
                eobi,
                provident_fund,
                payroll_tax,
                take_home,
            ]
        except Exception as e:
            logging.error(f"Error extracting all values: {e}")
            return [None] * 8

    def save_to_excel(self):
        # Create a DataFrame from the salary data list
        df = pd.DataFrame(self.salary_data, columns=[
            "Total Earnings",
            "Overtime",
            "Commission/Bonus",
            "Total Deductions",
            "EOBI",
            "PF",
            "Payroll Tax",
            "Take Home Pay"
        ])

        # Convert numeric fields to numbers
        numeric_columns = [
            "Total Earnings",
            "Overtime",
            "Commission/Bonus",
            "Total Deductions",
            "EOBI",
            "PF",
            "Payroll Tax",
            "Take Home Pay"
        ]
        for column in numeric_columns:
            if column in df.columns:
                # Remove commas and convert to float
                df[column] = df[column].str.replace(',', '').astype(float)

                # Place 0 in empty fields
                df.fillna(0, inplace=True)

        df["Basic With Medical Allowance"] = df["Total Earnings"] - df["Commission/Bonus"] - df["Overtime"]

        # Save to Excel file
        df.to_excel("salary_slips.xlsx", index=False)
        logging.info("Data saved to salary_slips.xlsx")



helper = SalarySlipsManager()
helper.load_config()
helper.connect_to_gmail_imap()

# Search for emails with subject containing "Payslip for the month of"
search_text = "Payslip for the month of"
email_ids = helper.search_emails(search_text)

if email_ids:
    for email_id in email_ids:
        email_msg = helper.fetch_email(email_id)
        if email_msg:
            # Extract PDF attachment
            pdf_data = helper.extract_pdf_from_attachment(email_msg)
            if pdf_data:
                # Read PDF content
                pdf_text = helper.read_pdf_content(pdf_data)
                if pdf_text:
                    # Extract all values from the PDF
                    extracted_values = helper.extract_all_values(pdf_text)
                    if any(extracted_values):  # Only add to the list if there's any extracted value
                        helper.salary_data.append(extracted_values)
                        logging.info(f"Extracted Values: {extracted_values}")
                    else:
                        logging.warning("No relevant data found in the PDF.")
                else:
                    logging.warning("Failed to read PDF content.")
else:
    logging.info("No emails found with the specified subject.")

# Save the extracted data to an Excel file
helper.save_to_excel()


# Earnings Amount PKR
# Basic+MedicalAllowance (1.00@
# 300000.00) 300,000.00
# Total Earnings 300,000.00
# Deductions Amount PKR
# PayrollTax 37,651.00
# ProvidentFundContributionEmployee 27,273.00
# EOBIContributionEmployee 370.00
# Total Deductions 65,294.00
# Take Home Pay 234,706.00PAYSLIP
# Awaisul HassanPay Day
# 31Dec2024
# Pay Period
# 1Dec2024to31Dec2024
