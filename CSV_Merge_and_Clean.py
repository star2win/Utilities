import csv
import re

def clean_csv(input_file, output_file):
    with open(input_file, mode='r', newline='') as infile, open(output_file, mode='w', newline='') as outfile:
        reader = csv.DictReader(infile)
        fieldnames = ['Email Address', 'First Name', 'Last Name', 'Company']
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        
        writer.writeheader()
        
        for row in reader:
            if row['Email Address']:
                first_name = ''
                last_name = ''
                if row['Customer Name']:
                    last_name, first_name = split_name(row['Customer Name'])
                writer.writerow({
                    'Email Address': row['Email Address'],
                    'First Name': first_name,
                    'Last Name': last_name,
                    'Company': row['Company'] if row['Company'] else ''
                })

def clean_name_part(name_part):
    """
    Clean a name part by removing words with consecutive XX or more,
    and words containing numbers.
    
    Args:
        name_part: A string containing part of a name
        
    Returns:
        Cleaned name part with problematic words removed
    """
    if not name_part:
        return name_part
    
    # Split the name part into words
    words = name_part.split()
    
    # Filter out words with consecutive XX or more, or containing numbers
    cleaned_words = []
    for word in words:
        # Skip words with consecutive XX or more
        if 'XX' in word.upper():
            continue
        
        # Skip words containing numbers
        if any(char.isdigit() for char in word):
            continue
        
        cleaned_words.append(word)
    
    # Join the remaining words back together
    return ' '.join(cleaned_words)

def split_name(name):
    """
    Split a name string into last_name and first_name.
    If there are multiple commas, split on the last comma.
    If first_name contains a nickname in parentheses, use the nickname but preserve the rest of the name.
    Remove words with consecutive XX or more, and words containing numbers.
    
    Args:
        name: A string containing the full name, potentially with commas and nicknames
        
    Returns:
        A tuple of (last_name, first_name), properly capitalized and cleaned
    """
    if not name:
        return ('', '')
    
    # Find the last comma in the name
    last_comma_index = name.rfind(', ')
    
    if last_comma_index != -1:
        # Split on the last comma
        last_name = name[:last_comma_index].strip()
        first_name = name[last_comma_index + 2:].strip()  # +2 to skip the comma and space
        
        # Check for nickname in parentheses in the first name
        nickname_match = re.search(r'(\w+)\s*\((\w+)\)', first_name)
        if nickname_match:
            # Get the nickname and any text after it
            formal_name = nickname_match.group(1)
            nickname = nickname_match.group(2)
            
            # Find the end position of the parentheses
            end_paren_pos = first_name.find(')', nickname_match.start()) + 1
            
            # Get any text after the parentheses
            remainder = first_name[end_paren_pos:].strip()
            
            # Use the nickname and add the remainder
            if remainder:
                first_name = f"{nickname} {remainder}"
            else:
                first_name = nickname
        
        # Clean the name parts
        last_name = clean_name_part(last_name)
        first_name = clean_name_part(first_name)
        
        return (last_name.title(), first_name.title())
    else:
        # No comma found, assume the entire string is the last name
        # Still check for nickname in case the format is different
        nickname_match = re.search(r'(\w+)\s*\((\w+)\)', name)
        if nickname_match:
            last_name = name[:nickname_match.start()].strip()
            nickname = nickname_match.group(2)
            
            # Find the end position of the parentheses
            end_paren_pos = name.find(')', nickname_match.start()) + 1
            
            # Get any text after the parentheses
            remainder = name[end_paren_pos:].strip()
            
            # Create first name with nickname and remainder
            if remainder:
                first_name = f"{nickname} {remainder}"
            else:
                first_name = nickname
            
            # Clean the name parts
            last_name = clean_name_part(last_name)
            first_name = clean_name_part(first_name)
                
            return (last_name.title(), first_name.title())
        
        # Clean the name part
        cleaned_name = clean_name_part(name.strip())
        
        return (cleaned_name.title(), '')

def is_valid_email(email):
    """
    Check if a string is a valid email address.
    
    Args:
        email: A string to check
        
    Returns:
        True if the string is a valid email address, False otherwise
    """
    # Basic email validation pattern
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))

def split_emails(email_str):
    """
    Split a string of multiple email addresses into a list of individual emails.
    Handles various separators like comma, semicolon, space, slash, etc.
    Only returns valid email addresses.
    
    Args:
        email_str: A string containing one or more email addresses
        
    Returns:
        A tuple of (valid_emails, has_invalid) where valid_emails is a list of valid email addresses
        and has_invalid is a boolean indicating if there were any invalid parts
    """
    if not email_str:
        return [], False
    
    # Replace common separators with a single separator
    normalized = re.sub(r'[;,\s/]+', ',', email_str)
    
    # Split by comma and filter out empty strings
    email_parts = [email.strip() for email in normalized.split(',') if email.strip()]
    
    # Filter for valid emails
    valid_emails = []
    has_invalid = False
    
    for part in email_parts:
        if is_valid_email(part):
            valid_emails.append(part)
        else:
            has_invalid = True
    
    return valid_emails, has_invalid

def merge_and_clean_csv(existing_file, new_file, output_file):
    """
    Merge two customer lists based on email address.
    
    Args:
        existing_file: Path to the existing comprehensive customer list CSV
        new_file: Path to the new customer list CSV with email, Name, and company_name
        output_file: Path to save the merged output CSV
    """
    # Read existing customer list into a dictionary with email as key
    existing_customers = {}
    existing_fieldnames = []
    no_email_rows = []
    
    with open(existing_file, mode='r', newline='', encoding='utf-8') as infile:
        reader = csv.DictReader(infile)
        existing_fieldnames = reader.fieldnames
        for row in reader:
            if 'email_address' in row and row['email_address']:
                email = row['email_address'].lower()
                if is_valid_email(email):
                    existing_customers[email] = row
                else:
                    # Invalid email, add to no_email_rows
                    row_copy = row.copy()
                    row_copy['email_address'] = 'NO EMAIL'
                    no_email_rows.append(row_copy)
    
    # Create output fieldnames with Notes column
    output_fieldnames = existing_fieldnames.copy()
    if 'Notes' not in output_fieldnames:
        output_fieldnames.append('Notes')
    if 'Name' not in output_fieldnames:
        output_fieldnames.append('Name')
    
    # Create a merged dictionary to store all customers
    merged_customers = existing_customers.copy()
    
    # Process existing customers first to extract first and last names from Name field if needed
    for email, customer in merged_customers.items():
        # Check if customer has Name field but missing first_name or last_name
        if ('Name' in customer and customer['Name'] and 
            (not customer.get('first_name') or not customer.get('last_name'))):
            
            last_name, first_name = split_name(customer['Name'])
            customer['last_name'] = last_name
            customer['first_name'] = first_name
    
    # Create a consolidated row for invalid emails
    no_email_consolidated = None
    
    # Process new customer list
    with open(new_file, mode='r', newline='', encoding='utf-8') as infile:
        reader = csv.DictReader(infile)
        
        for row in reader:
            # Check if email_address field exists and has content
            if 'email_address' not in row or not row['email_address']:
                # No email address, add to no_email_consolidated
                if no_email_consolidated is None:
                    no_email_consolidated = row.copy()
                    no_email_consolidated['email_address'] = 'NO EMAIL'
                    if 'Notes' in no_email_consolidated:
                        no_email_consolidated['Notes'] += '; MISSING EMAIL'
                    else:
                        no_email_consolidated['Notes'] = 'MISSING EMAIL'
                continue
            
            # Split and validate email addresses
            valid_emails, has_invalid = split_emails(row['email_address'])
            
            # If no valid emails were found, add to no_email_consolidated
            if not valid_emails:
                if no_email_consolidated is None:
                    no_email_consolidated = row.copy()
                    no_email_consolidated['email_address'] = 'NO EMAIL'
                    if 'Notes' in no_email_consolidated:
                        no_email_consolidated['Notes'] += '; INVALID EMAIL'
                    else:
                        no_email_consolidated['Notes'] = 'INVALID EMAIL'
                continue
            
            # Process each valid email
            for i, email in enumerate(valid_emails):
                email = email.lower()
                
                # Create a copy of the row for this email
                current_row = row.copy()
                current_row['email_address'] = email
                
                # Add DUPLICATE to Notes if this is a split record (more than one email)
                if len(valid_emails) > 1:
                    if 'Notes' in current_row and current_row['Notes']:
                        current_row['Notes'] += '; DUPLICATE'
                    else:
                        current_row['Notes'] = 'DUPLICATE'
                
                # Add INVALID EMAIL PARTS to Notes if there were invalid parts
                if has_invalid:
                    if 'Notes' in current_row and current_row['Notes']:
                        current_row['Notes'] += '; INVALID EMAIL PARTS'
                    else:
                        current_row['Notes'] = 'INVALID EMAIL PARTS'
                
                if email in merged_customers:
                    # Customer already exists, update Notes field
                    notes = f"Customer Name: '{current_row['Name']}' and Company Name: '{current_row['company_name']}'"
                    
                    # Add the notes to existing Notes or create new
                    if 'Notes' in merged_customers[email] and merged_customers[email]['Notes']:
                        merged_customers[email]['Notes'] += f"; {notes}"
                        
                        # Add DUPLICATE if needed
                        if 'DUPLICATE' in current_row.get('Notes', '') and 'DUPLICATE' not in merged_customers[email]['Notes']:
                            merged_customers[email]['Notes'] += "; DUPLICATE"
                        
                        # Add INVALID EMAIL PARTS if needed
                        if 'INVALID EMAIL PARTS' in current_row.get('Notes', '') and 'INVALID EMAIL PARTS' not in merged_customers[email]['Notes']:
                            merged_customers[email]['Notes'] += "; INVALID EMAIL PARTS"
                    else:
                        merged_customers[email]['Notes'] = notes
                        
                        # Add DUPLICATE if needed
                        if 'DUPLICATE' in current_row.get('Notes', ''):
                            merged_customers[email]['Notes'] += "; DUPLICATE"
                        
                        # Add INVALID EMAIL PARTS if needed
                        if 'INVALID EMAIL PARTS' in current_row.get('Notes', ''):
                            merged_customers[email]['Notes'] += "; INVALID EMAIL PARTS"
                    
                    # Only add the 'Name' field if it doesn't already exist or is empty
                    if 'Name' not in merged_customers[email] or not merged_customers[email]['Name']:
                        merged_customers[email]['Name'] = current_row['Name']
                        
                        # If first_name or last_name is missing, extract from Name
                        if not merged_customers[email].get('first_name') or not merged_customers[email].get('last_name'):
                            last_name, first_name = split_name(current_row['Name'])
                            merged_customers[email]['last_name'] = last_name
                            merged_customers[email]['first_name'] = first_name
                else:
                    # New customer, add to merged list
                    new_customer = {field: '' for field in output_fieldnames}
                    new_customer['email_address'] = email
                    new_customer['company_name'] = current_row['company_name']
                    
                    # Preserve the original 'Name' field
                    new_customer['Name'] = current_row['Name']
                    
                    # Split and capitalize first and last name
                    if 'Name' in current_row and current_row['Name']:
                        last_name, first_name = split_name(current_row['Name'])
                        new_customer['last_name'] = last_name
                        new_customer['first_name'] = first_name
                    
                    # Set optin_status
                    new_customer['optin_status'] = 'merged_upload'
                    
                    # Add Notes if needed
                    if 'Notes' in current_row and current_row['Notes']:
                        new_customer['Notes'] = current_row['Notes']
                    elif len(valid_emails) > 1:
                        new_customer['Notes'] = 'DUPLICATE'
                    
                    # Add INVALID EMAIL PARTS if needed
                    if has_invalid:
                        if 'Notes' in new_customer and new_customer['Notes']:
                            new_customer['Notes'] += '; INVALID EMAIL PARTS'
                        else:
                            new_customer['Notes'] = 'INVALID EMAIL PARTS'
                    
                    # Add to merged dictionary
                    merged_customers[email] = new_customer
    
    # Add the consolidated no_email row if it exists
    if no_email_consolidated is not None:
        # If there are already no_email rows, merge with the first one
        if no_email_rows:
            # Update the first no_email row with info from no_email_consolidated
            if 'Notes' in no_email_rows[0] and no_email_rows[0]['Notes']:
                no_email_rows[0]['Notes'] += f"; Customer Name: '{no_email_consolidated.get('Name', '')}' and Company Name: '{no_email_consolidated.get('company_name', '')}'"
            else:
                no_email_rows[0]['Notes'] = f"Customer Name: '{no_email_consolidated.get('Name', '')}' and Company Name: '{no_email_consolidated.get('company_name', '')}'"
            
            # If Name is missing, add it
            if ('Name' not in no_email_rows[0] or not no_email_rows[0]['Name']) and 'Name' in no_email_consolidated:
                no_email_rows[0]['Name'] = no_email_consolidated['Name']
                
                # Extract first_name and last_name
                if 'Name' in no_email_consolidated and no_email_consolidated['Name']:
                    last_name, first_name = split_name(no_email_consolidated['Name'])
                    no_email_rows[0]['last_name'] = last_name
                    no_email_rows[0]['first_name'] = first_name
        else:
            # No existing no_email rows, add the consolidated one
            no_email_rows.append(no_email_consolidated)
    
    # Write all merged customers to output file
    with open(output_file, mode='w', newline='', encoding='utf-8') as outfile:
        writer = csv.DictWriter(outfile, fieldnames=output_fieldnames)
        writer.writeheader()
        
        # Write all valid email customers
        for customer in merged_customers.values():
            writer.writerow(customer)
        
        # Write all no_email rows
        for row in no_email_rows:
            writer.writerow(row)

if __name__ == "__main__":
    # Original clean_csv function usage
    # input_file = '../BME/Client_Email_List_3.csv'
    # output_file = '../BME/Cleaned_Client_Email_List.csv'
    # clean_csv(input_file, output_file)
    
    # New merge_and_clean_csv function usage
    existing_file = '../BME/4485b5fd92_Master_List_download_Mar_21_2025.csv'
    new_file = '../BME/Client_Email_List.csv'
    output_file = '../BME/Merged_Client_Email_List.csv'
    merge_and_clean_csv(existing_file, new_file, output_file)