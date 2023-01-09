import openpyxl
from validate_email import validate_email
from tqdm import tqdm

def validate_emails(filename):
    # Create a set to store the email addresses that have already been processed
    processed_emails = set()
  
    # Load the workbook
    wb = openpyxl.load_workbook(filename)
    
    # Select the first worksheet
    ws = wb[wb.sheetnames[0]]
    
    # Open the progress file
    with open("progress.txt", "r") as f:
        # Read the last row number from the file
        last_row = int(f.read())
    
    # Open the files for writing
    with open("valid_emails.txt", "a") as valid_file, open("invalid_emails.txt", "a") as invalid_file:
        # Get the total number of rows in the worksheet
        total_rows = ws.max_row
        
        # Create a progress bar using tqdm
        with tqdm(total=total_rows, desc="Validating emails") as pbar:
            # Iterate through the rows in the worksheet
            for i, row in enumerate(ws.rows):
                # Skip rows that have already been processed
                if i <= last_row:
                    continue
                
                # Get the name and email address from the first and second columns
                name = row[0].value
                email = row[1].value

              # Skip email addresses that have already been processed
                if email in processed_emails:
                    continue
                
                # Validate the email address
                is_valid = validate_email(email_address=email,
                                          check_format=True,
                                          check_blacklist=True,
                                          check_dns=True,dns_timeout=10,
                                          check_smtp=True,
                                          smtp_timeout=10,
                                          smtp_helo_host='my.host.name',
                                          smtp_from_address='admin@emancipation.co.in',
                                          smtp_skip_tls=False,
                                          smtp_tls_context=None,
                                          smtp_debug=False)
                if is_valid:
                    # Write the name and email to the valid file
                    valid_file.write(f"{name},{email}\n")
                else:
                    # Write the email to the invalid file
                    invalid_file.write(f"{email}\n")

                # Add the email address to the set of processed emails
                processed_emails.add(email)
                
                # Update the progress bar
                pbar.update(1)
                
                # Update the progress file
                with open("progress.txt", "w") as f:
                    f.write(str(i))


validate_emails("sampleemail.xlsx")
