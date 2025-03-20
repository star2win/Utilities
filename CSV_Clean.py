import csv

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
                    name_parts = row['Customer Name'].split(', ')
                    if len(name_parts) == 2:
                        last_name, first_name = name_parts
                    else:
                        last_name = row['Customer Name']
                writer.writerow({
                    'Email Address': row['Email Address'],
                    'First Name': first_name,
                    'Last Name': last_name,
                    'Company': row['Company'] if row['Company'] else ''
                })

if __name__ == "__main__":
    input_file = '../BME/Client_Email_List_3.csv'
    output_file = '../BME/Cleaned_Client_Email_List.csv'
    clean_csv(input_file, output_file)