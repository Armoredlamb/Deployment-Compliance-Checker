import csv
from collections import Counter
from openpyxl import Workbook

# Purpose: identify hosts that do not have a 1:1 compliance of 2 agents.

# Ask for the filepath and name of the .csv file
csv_filepath = input("Enter the filepath and name of the .csv file: ")

# Read the CSV file and identify hostnames that appear only once
hostnames_counter = Counter()

with open(csv_filepath, 'r') as csv_file:
    csv_reader = csv.reader(csv_file)
    header = next(csv_reader)  # Read the header
    column_index = header.index('Hostname')

    for row in csv_reader:
        hostname = row[column_index]
        hostnames_counter[hostname] += 1

unique_hostnames = {hostname for hostname, count in hostnames_counter.items() if count == 1}

# Export identified rows to a new Excel workbook
new_workbook = Workbook()
worksheet = new_workbook.active

# Write the header row
worksheet.append(header)

# Write the filtered data rows
with open(csv_filepath, 'r') as csv_file:
    csv_reader = csv.reader(csv_file)
    next(csv_reader)  # Skip the header
    for row in csv_reader:
        if row[column_index] in unique_hostnames:
            worksheet.append(row)

# Save the new Excel workbook
new_filepath = csv_filepath.replace('.csv', '_AgentComplianceReport.xlsx')
new_workbook.save(new_filepath)

print(f"Filtered data has been exported to {new_filepath}")
