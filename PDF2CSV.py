import pdfplumber
import re
import csv

pdf_path = "../BME/Client_Email.pdf"
csv_path = "../BME/ClientEmailOutput.csv"

email_regex = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
excluded_phrases = [
    "Bavarian Motor Experts",
    "Customer E-mail by Last Name Printed:",
    "E-mail Address Customer Name Company",
    "(C) 2011  Mitchell Repair Information Company, LLC",
    "Page ",  # e.g. "Page 1 of 212"
    "Honolulu, HI.  96813",
    "757 Kawaiahao St."
]

def looks_like_header_or_footer(line):
    # If any known unwanted phrase is present, exclude
    for phrase in excluded_phrases:
        if phrase in line:
            return True
    return False

def parse_line(line):
    line = line.strip()
    match = email_regex.search(line)
    if not match:
        return None, None, None
    
    email = match.group(0)
    before_email = line[: match.start()].strip()
    after_email = line[match.end():].strip()
    
    # In your sample, the pattern you mention looks like:
    #   email + " " + name_or_company
    # or 
    #   name_or_company + " " + email
    #
    # We’ll check which side is bigger, so we can guess what’s “name/company.”

    # If the email appears near the start, then the remainder is presumably name/company:
    # e.g. "jaymiseo@gmail.com LUKACS, JAYMI" => name is "LUKACS, JAYMI"
    # If the email appears near the end, the part before is presumably name/company:
    # e.g. "LUKACS, JAYMI jaymiseo@gmail.com"
    
    # We'll see which side is bigger, or just pick the side that contains a comma 
    # as the name, or otherwise treat it as a company.
    
    # Heuristic approach: 
    # 1) If after_email is non-empty, we treat that as name/company. Otherwise, we treat before_email as name/company.
    # 2) If the portion we treat as name/company has a comma, treat it as name; else treat as company.

    # Which side is the "name/company"?
    if after_email:
        # e.g. "jaymiseo@gmail.com LUKACS, JAYMI"
        name_company = after_email
    else:
        # e.g. "LUKACS, JAYMI  jaymiseo@gmail.com"
        name_company = before_email
    
    # Now decide whether it's a name or a company:
    if "," in name_company:
        return email, name_company, ""
    else:
        return email, "", name_company

with pdfplumber.open(pdf_path) as pdf, open(csv_path, "w", newline="", encoding="utf-8") as fout:
    writer = csv.writer(fout)
    writer.writerow(["e-mail address", "customer name", "company"])
    
    for page in pdf.pages:
        lines = page.extract_text().split("\n")
        
        buffer = []
        
        for line in lines:
            line = line.strip()
            if not line or looks_like_header_or_footer(line):
                continue  # skip empty or unwanted lines

            # If this line has an email:
            if email_regex.search(line):
                # Combine buffered lines + current line
                combined = " ".join(buffer + [line])
                email, name, company = parse_line(combined)
                if email:
                    writer.writerow([email, name, company])
                buffer = []  # reset for next record
            else:
                # Accumulate lines until we encounter an email
                buffer.append(line)

        # If leftover lines remain after the last email on the page,
        # we ignore them because they likely have no email anyway.

print(f"CSV file created at {csv_path}")