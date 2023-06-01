import csv
import os
from bs4 import BeautifulSoup
import time
import datetime
import shutil

# Getting the current date and time to rename the original excel file to this format and save it to the History folder
current_time = datetime.date.today()
t = time.localtime()
current_time = str(current_time) + str(time.strftime(" %H-%M-%S", t))

# this method converts the whole HTML file to CSV file (without updating the column X) and saves it in the Triumph folder
def convert_table_to_csv(html_file):
    # Read the HTML file
    with open(html_file, 'r') as file:
        html_content = file.read()

    # Create a BeautifulSoup object
    soup = BeautifulSoup(html_content, 'html.parser')

    # Find the table element
    table = soup.find('table')

    # Extract the table headers
    header_row = table.find('tr')
    headers = header_row.find_all('th')
    header_values = [header.get_text().strip() for header in headers]

    # Extract the table data
    data_rows = table.find_all('tr')[1:]  # Exclude header row
    data_values = []
    for row in data_rows:
        cells = row.find_all('td')
        row_values = [cell.get_text().strip() for cell in cells]
        data_values.append(row_values)

    # Save the table as a CSV file
    base_dir = "C://temp//Triumph"
    csv_filename = os.path.join(base_dir, 'tms_shipments.csv')
    with open(csv_filename, 'w', newline='') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(header_values)
        writer.writerows(data_values)


# this method simply makes a copy of the csv file and adds 1 at the ned of its name
def copy_csv_file_with_suffix(csv_file_path, suffix):
    folder_path, file_name = os.path.split(csv_file_path)
    new_file_name = f"{file_name.split('.')[0]}{suffix}.{file_name.split('.')[1]}"
    new_file_path = os.path.join(folder_path, new_file_name)
    with open(csv_file_path, 'rb') as file_in, open(new_file_path, 'wb') as file_out:
        file_out.write(file_in.read())
    return new_file_path


# this method cuts the original html file and puts it in the destination folder (History)
def cut_excel_file(source_file_path, destination_folder_path):
    file_name = os.path.basename(current_time + ".html")
    destination_path = os.path.join(destination_folder_path, file_name)
    os.rename(source_file_path, destination_path)
    return destination_path

# this method makes sure than the Upload folder has only 1 file (xls) and then converts it to html file
def rename_xls_to_html(folder_path):
    files = os.listdir(folder_path)
    if len(files) != 1:
        print("There should be one file in the Upload folder.")
        print("Terminating the script...")
        exit()
    if len(files) == 0:
        print("No file exists in the Upload folder!")
        return
    old_file_path = os.path.join(folder_path, files[0])
    new_file_path = os.path.join(folder_path, str(files[0]) + ".html")
    os.rename(old_file_path, new_file_path)


# this method converts the html file back into xls file (97-2003 excel file format)
def rename_html_to_xls(file_path):
    file_name, _ = os.path.splitext(file_path)
    new_file_path = file_name + ".xls"
    os.rename(file_path, new_file_path)

def remove_character_from_column(csv_file, column_title):
    # Read the CSV file
    with open(csv_file, 'r') as file:
        reader = csv.DictReader(file)
        rows = list(reader)

    # Find the index of the column
    column_index = None
    for index, header in enumerate(reader.fieldnames):
        if header == column_title:
            column_index = index
            break

    if column_index is None:
        print(f"Column with title '{column_title}' not found in the CSV file.")
        return

    # Modify the data in the column
    for row in rows:
        if len(row[column_title]) >= 11:
            row[column_title] = row[column_title][:10] + row[column_title][11:]

    # Save the modified data back to the CSV file
    with open(csv_file, 'w', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=reader.fieldnames)
        writer.writeheader()
        writer.writerows(rows)


# a method which executes every day at midnight (00:00:00) and cut and paste 2 csv files (tms_shipments) and renames it
def daily_run():
    now = datetime.datetime.now() 

    if now.hour == 0 and now.minute == 0:
        source_file = 'C://temp//Triumph//tms_shipments.csv'
        source_file1 = 'C://temp//Triumph//tms_shipments1.csv'

        destination_folder = 'C://temp//Triumph//History'
        destination_name = current_time + '.csv'
        destination_name1 = current_time + '1.csv'

        try:
            # Move and rename 2 files
            shutil.move(source_file, os.path.join(destination_folder, destination_name))
            shutil.move(source_file1, os.path.join(destination_folder, destination_name1))

            print("Files moved and renamed successfully.")
        except FileNotFoundError:
            print("Source file not found.")
        except Exception as e:
            print("An error occurred:", e)

try:

    print("Starting the script...")

    # Specify the directory where the HTML files are located
    html_dir = "C://temp//Triumph//Upload"

    # convert the .xls file to .html file
    rename_xls_to_html(html_dir)

    print("Converting .xls into .csv...")

    # Iterate through all HTML files in the directory and convert each table to a CSV file
    for filename in os.listdir(html_dir):
        if filename.endswith(".html"):
            html_file = os.path.join(html_dir, filename)
            convert_table_to_csv(html_file) # convert the .html file to .csv file

    print("Processing the csv file...")

    # Provide the path to your CSV file
    csv_file = "C://temp//Triumph//tms_shipments.csv"

    # Provide the column title
    column_title = "Harmonized Code"

    # Call the function to remove the 11th character from the specified column
    remove_character_from_column(csv_file, column_title)


    print("Making a copy of the CSV file...")
    my_path = 'C://temp//Triumph//tms_shipments.csv'
    suffix = "1"

    # make another copy from the csv file
    copy_csv_file_with_suffix(my_path, suffix)

    print("Created tms_shipments.csv successfully!")
    print("Created tms_shipments1.csv successfully!")


    print("Moving the Excel file to the History folder...")

    source_file_path = html_file
    destination_folder_path = "C://temp//Triumph//History"
    destination_file_path = cut_excel_file(source_file_path, destination_folder_path) # move the original file to the History folder


    rename_html_to_xls("C://temp//Triumph//History//" + current_time + ".html") # rename the html file to csv file

    print("100% SUCCESSFUL!")

    print("Daily run...")
    daily_run()

    # Prompt the user to press Enter to exit
    input("Press Enter to exit...")

except FileNotFoundError as e:
    print(f"Error: {e}")
    
except Exception as e:
    print(f"An error occurred: {e}")