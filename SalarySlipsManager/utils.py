import imaplib
import json
import email
import logging
import os
import PyPDF2
import pandas as pd

from fuzzywuzzy import fuzz
from io import BytesIO


class SalarySlipsManager:
    def __init__(self):
        self.email = None
        self.password = None
        self.mail = None
        self.salary_data = []  # List to hold the earnings and tax data
        self.pdf_text = ""

    def load_config(self):
        try:
            config_path = os.path.join(os.path.dirname(__file__), "config.json")
            if not os.path.exists(config_path):
                raise FileNotFoundError(f"Config file not found: {config_path}")
            with open(config_path, "r") as f:
                config = json.load(f)
                self.email = config["user"]
                self.password = config["pass"]
        except Exception as e:
            print("Failed to load config: {}".format(e))

    def connect_to_gmail_imap(self):
        imap_url = "imap.gmail.com"
        try:
            self.mail = imaplib.IMAP4_SSL(imap_url)
            self.mail.login(self.email, self.password)
            self.mail.select("inbox")  # Connect to the inbox.
            print("Connected to Gmail IMAP successfully.")
        except Exception as e:
            print(f"Connection failed: {e}")

    def search_emails(self, search_text):
        try:
            # Search for emails with specific subject text
            status, messages = self.mail.search(None, f'(SUBJECT "{search_text}")')
            if status != "OK":
                print("No messages found with the specified subject.")
                return []
            message_ids = messages[0].split()
            return message_ids
        except Exception as e:
            print(f"Failed to search emails: {e}")
            return []

    def fetch_email(self, email_id):
        try:
            status, msg_data = self.mail.fetch(email_id, "(RFC822)")
            if status != "OK":
                print(f"Failed to fetch email ID {email_id}")
                return None
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    return msg
        except Exception as e:
            print(f"Failed to fetch email: {e}")
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
            print(f"Error extracting PDF: {e}")
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
            print(f"Error reading PDF: {e}")
            return None

    def extract_value(self, keyword, default_value=0):
        lines = self.pdf_text.split("\n")
        for line in lines:
            # Perform fuzzy matching
            if fuzz.partial_ratio(keyword.lower(), line.lower()) > 80:
                # Look for numbers after the keyword
                words = line.split()
                for word in words:
                    # Check if word is a number
                    if word.replace(',', '').replace('.', '').isdigit():
                        return word
        return default_value
    
    def extract_all_values(self, text):
        try:
            self.pdf_text = text

            # Extract values using keywords
            total_earnings = self.extract_value("Total Earnings")
            overtime = self.extract_value("Overtime")
            commission_bonus = self.extract_value("Commission/Bonus")
            bonus_winner = self.extract_value("Bonus / Winners")
            total_deductions = self.extract_value("Total Deductions")
            provident_fund = self.extract_value("Provident Fund Contribution Employee")
            eobi = self.extract_value("EOBI Contribution")
            payroll_tax = self.extract_value("Payroll Tax")
            medical = self.extract_value("Medical / OPD Reimbursement")

            return [
                total_earnings,
                overtime,
                commission_bonus,
                bonus_winner,
                total_deductions,
                eobi,
                provident_fund,
                payroll_tax,
                medical
            ]
        except Exception as e:
            print(f"Error extracting all values: {e}")
            return [None] * 8

    def save_to_excel(self):
        # Create a DataFrame from the salary data list
        df = pd.DataFrame(self.salary_data, columns=[
            "Earnings",
            "Overtime",
            "CommissionBonus",
            "BonusWinners",
            "Deductions",
            "EOBI",
            "PF",
            "Tax",
            "Medical"
        ])

        # Convert numeric fields to strings and remove commas
        numeric_columns = [
            "Earnings",
            "Overtime",
            "CommissionBonus",
            "BonusWinners",
            "Deductions",
            "EOBI",
            "PF",
            "Tax",
            "Medical"
        ]
        for column in numeric_columns:
            if column in df.columns:
                # Convert values to strings, then remove commas, and finally convert to float
                df[column] = df[column].fillna('0').astype(str).str.replace(',', '')
                # Convert to float
                df[column] = pd.to_numeric(df[column], errors='coerce').fillna(0)

        # Calculate 'Basic With Medical Allowance'
        df["Total Salary"] = df["Earnings"] - df["CommissionBonus"] - df["BonusWinners"] - df["Overtime"] - df["Medical"]
        df["Total Bonus"] = df["CommissionBonus"] + df["BonusWinners"]
        df["Total After Deductions"] = df["Earnings"] - df["Deductions"]

        # Remove the bonus columns 
        df.drop(columns=["CommissionBonus", "BonusWinners"], inplace=True)

        #Iterate through the dataframe and add a column with the incremental month starting from October 2023
        df["Month"] = pd.date_range(start="2023-10-01", periods=len(df), freq='M').strftime("%B %Y")

        # Save to Excel file
        df.to_excel("Joblogic Salary Details.xlsx", index=False)
        print("Data saved to Joblogic Salary Details.xlsx")


    def save_slip(self, msg, file_name):
        """
        Fetch the email by ID and extract the single PDF attachment,
        saving it to the output_dir.
        """
        output_dir=r"C:\Users\hawai\Desktop\ME\Companies\JL\Salary Slips"
        for part in msg.walk():
            # Skip any multipart containers
            if part.get_content_maintype() == 'multipart':
                continue

            # Check for a Content-Disposition header (indicating an attachment)
            if part.get("Content-Disposition") is None:
                continue

            filename = part.get_filename()
            if filename and filename.lower().endswith(".pdf"):
                # Use the filename from the argument if provided, otherwise use the attachment's filename
                filepath = os.path.join(output_dir, file_name + ".pdf")
                try:
                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    print(f"PDF attachment saved to {filepath}")
                    return filepath  # Return the path once the PDF is saved
                except Exception as e:
                    print(f"Failed to save PDF attachment: {e}")
                    return None

        print("No PDF attachment found in the email.")
        return None

