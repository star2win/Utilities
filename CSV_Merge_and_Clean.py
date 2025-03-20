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
    
    with open(existing_file, mode='r', newline='', encoding='utf-8') as infile:
        reader = csv.DictReader(infile)
        existing_fieldnames = reader.fieldnames
        for row in reader:
            existing_customers[row['email_address'].lower()] = row
    
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
    
    # Process new customer list
    with open(new_file, mode='r', newline='', encoding='utf-8') as infile:
        reader = csv.DictReader(infile)
        
        for row in reader:
            email = row['email_address'].lower() if 'email_address' in row else ''
            
            if not email:
                continue
                
            if email in merged_customers:
                # Customer already exists, update Notes field
                notes = f"Customer Name: '{row['Name']}' and Company Name: '{row['company_name']}'"
                merged_customers[email]['Notes'] = notes
                
                # Only add the 'Name' field if it doesn't already exist or is empty
                if 'Name' not in merged_customers[email] or not merged_customers[email]['Name']:
                    merged_customers[email]['Name'] = row['Name']
                    
                    # If first_name or last_name is missing, extract from Name
                    if not merged_customers[email].get('first_name') or not merged_customers[email].get('last_name'):
                        last_name, first_name = split_name(row['Name'])
                        merged_customers[email]['last_name'] = last_name
                        merged_customers[email]['first_name'] = first_name
            else:
                # New customer, add to merged list
                new_customer = {field: '' for field in output_fieldnames}
                new_customer['email_address'] = row['email_address']
                new_customer['company_name'] = row['company_name']
                
                # Preserve the original 'Name' field
                new_customer['Name'] = row['Name']
                
                # Split and capitalize first and last name
                if 'Name' in row and row['Name']:
                    last_name, first_name = split_name(row['Name'])
                    new_customer['last_name'] = last_name
                    new_customer['first_name'] = first_name
                
                # Set optin_status
                new_customer['optin_status'] = 'merged_upload'
                
                # Add to merged dictionary
                merged_customers[email] = new_customer
    
    # Write all merged customers to output file
    with open(output_file, mode='w', newline='', encoding='utf-8') as outfile:
        writer = csv.DictWriter(outfile, fieldnames=output_fieldnames)
        writer.writeheader()
        
        for customer in merged_customers.values():
            writer.writerow(customer)

if __name__ == "__main__":
    # Original clean_csv function usage
    # input_file = '../BME/Client_Email_List_3.csv'
    # output_file = '../BME/Cleaned_Client_Email_List.csv'
    # clean_csv(input_file, output_file)
    
    # New merge_and_clean_csv function usage
    existing_file = '../BME/73a0283eee_Master_List__1_17_2020_download.csv'
    new_file = '../BME/Client_Email_List.csv'
    output_file = '../BME/Merged_Client_Email_List.csv'
    merge_and_clean_csv(existing_file, new_file, output_file)