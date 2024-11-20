# Month_To_Date_Report

This script generates a month-to-date misload report for the UPS Maple Grove, MN Preload operation. It calculates the misload count for each loader for the current month and generates a report with the loader names, their corresponding misload counts, and frequencies.

The code consists of the following functions:

1. get_loader_misloads(doc): Extracts the loader names and misload counts from a Word document's table.

2. calculate_month_to_date_misloads(folder_path): Iterates through the Word documents in the given folder,
   retrieves the loader misloads using the get_loader_misloads function, and calculates the month-to-date misloads
   for each loader.

3. create_month_to_date_report(month_to_date_misloads, output_folder): Generates the month-to-date misloads report
   in a new Word document. It includes the report title, date, a table with the loader names and misload counts,
   and saves the document in the specified output folder.
