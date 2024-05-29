# Month_To_Date_Report

This script generates a month-to-date misloads report for the UPS Maple Grove, MN Preload operation.
It calculates the misload counts for each loader in the current month and creates a report with the loader names
and their corresponding misload counts. The report includes the date of the report and is saved in a specified output folder.

The code consists of the following functions:

1. get_loader_misloads(doc): Extracts the loader names and misload counts from a Word document's table.

2. calculate_month_to_date_misloads(folder_path): Iterates through the Word documents in the given folder,
   retrieves the loader misloads using the get_loader_misloads function, and calculates the month-to-date misloads
   for each loader.

3. create_month_to_date_report(month_to_date_misloads, output_folder): Generates the month-to-date misloads report
   in a new Word document. It includes the report title, date, a table with the loader names and misload counts,
   and saves the document in the specified output folder.
