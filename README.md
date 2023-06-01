# HTML-Table-to-CSV-Transformer

The script starts by importing the necessary modules, including csv, os, BeautifulSoup, time, and datetime. These modules are essential for file manipulation, data extraction, and time-related operations.

Next, the script defines various functions:

convert_table_to_csv(html_file):

This function converts an HTML table into a CSV file. It reads the HTML file, extracts the table headers and data, and saves them as a CSV file.

copy_csv_file_with_suffix(csv_file_path, suffix):

This function creates a copy of a CSV file with an added suffix at the end of its name. It takes the original file path and the desired suffix as input and returns the new file path.

cut_excel_file(source_file_path, destination_folder_path):

This function moves the original HTML file to a destination folder, typically the "History" folder. It renames the file using the current date and time and returns the new file path.

rename_xls_to_html(folder_path):

This function ensures that the "Upload" folder contains only one file in the XLS format. It renames the XLS file to an HTML file by appending the ".html" extension.

rename_html_to_xls(file_path):

This function converts an HTML file back to the XLS format. It renames the file by changing the extension from ".html" to ".xls".

remove_character_from_column(csv_file, column_title):

This function removes the 11th character from a specific column in a CSV file. It reads the CSV file, modifies the data in the specified column, and saves the modified data back to the file.

daily_run():

This function is intended to be executed daily at midnight (00:00:00). It moves and renames two CSV files, specifically "tms_shipments.csv" and "tms_shipments1.csv", to the "History" folder. The new file names include the current date and time.

The script then proceeds with the main execution:

It prompts the user with a starting message.

The HTML files in the "Upload" folder are converted to CSV files using the convert_table_to_csv function.

The specified column ("Harmonized Code") in the CSV file ("tms_shipments.csv") is processed using the remove_character_from_column function. It removes the 11th character from each value in the column.

A copy of the CSV file is made using the copy_csv_file_with_suffix function. The copy is saved with the suffix "1".

The original HTML file is moved to the "History" folder and renamed with the current date and time.

The HTML file in the "History" folder is converted back to the XLS format using the rename_html_to_xls function.

The script outputs a success message and initiates the daily_run function if the current time is midnight (00:00:00).

Finally, the user is prompted to press Enter to exit the script.

The script includes error handling for file-related exceptions and displays appropriate error messages if any issues occur during execution.
