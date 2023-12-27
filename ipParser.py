import re
import openpyxl

def extract_ips_from_excel(input_file, output_file):
    # Regular expression pattern for matching IP addresses
    ip_pattern = re.compile(r'\b(?:\d{1,3}\.){3}\d{1,3}\b')

    # Load the Excel workbook and select the active worksheet
    workbook = openpyxl.load_workbook(input_file)
    worksheet = workbook.active

    # Extract and store IP addresses
    ip_addresses = []
    for row in worksheet.iter_rows(values_only=True):
        for cell in row:
            if cell is not None:
                found_ips = ip_pattern.findall(str(cell))
                ip_addresses.extend(found_ips)

    # Write IP addresses to the output file
    with open(output_file, 'w') as file:
        for ip in ip_addresses:
            file.write(ip + '\n')

# Example usage
input_filename = '<add input file name here>'  # Replace with your actual Excel file name
output_filename = '<add output file name here>'
extract_ips_from_excel(input_filename, output_filename)
