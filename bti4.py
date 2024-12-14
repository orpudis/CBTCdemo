import zipfile
import xml.etree.ElementTree as ET
import math

def initialize_data():
    """
    Read data from an Excel (.xlsx) file, handling shared strings.
    
    :return: List of rows from the Excel file, or an empty list if there's an error.
    """
    data = []
    shared_strings = []
    
    try:
        with zipfile.ZipFile('C:/Users/User1/Desktop/YouCreation/birthtime3.xlsx') as zip_ref:
            # Read shared strings
            shared_strings_content = zip_ref.read('xl/sharedStrings.xml')
            root = ET.fromstring(shared_strings_content)
            shared_strings = [t.text for t in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')]
            
            # Read sheet data
            sheet_content = zip_ref.read('xl/worksheets/sheet1.xml')
            sheet_root = ET.fromstring(sheet_content)
            
            for row in sheet_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
                row_data = []
                for cell in row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                    if cell.get('t') == 's':  # shared string
                        row_data.append(shared_strings[int(cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v').text)])
                    else:
                        v = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                        row_data.append(v.text if v is not None else '')
                data.append(row_data)
        return data
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return []

def get_user_time():
    """
    Get birth time from user input with basic validation.
    
    :return: Tuple of (time_input, am_pm)
    """
    while True:
        time_input = input("Enter your Birth Time in HH:MM format: ")
        if len(time_input.split(':')) == 2 and all(x.isdigit() for x in time_input.replace(':', '')):
            hours, minutes = map(int, time_input.split(':'))
            if 0 <= hours < 24 and 0 <= minutes < 60:
                break
        print("Invalid time format. Please use HH:MM, where HH is 0-23 and MM is 0-59.")
    
    while True:
        am_pm = input("Is this AM or PM? (Type 'AM' or 'PM'): ").upper()
        if am_pm in ["AM", "PM"]:
            break
        print("Please enter either 'AM' or 'PM'.")
    
    return time_input, am_pm

def time_to_percentage(time_input, am_pm):
    """
    Convert time to percentage of day starting from 6 PM.
    
    :param time_input: Time in 'HH:MM' format
    :param am_pm: 'AM' or 'PM'
    :return: Percentage of the day
    """
    hours, minutes = map(int, time_input.split(':'))
    
    if am_pm == "PM" and hours != 12:
        hours += 12
    elif am_pm == "AM" and hours == 12:
        hours = 0
    
    if hours < 18:  # Before 6 PM
        hours += 24  # Shift to next day for calculation
    time_from_6pm = (hours - 18) * 60 + minutes
    
    total_minutes = 24 * 60
    percentage = (time_from_6pm / total_minutes) * 100
    return percentage

def find_row(percentage, dataset):
    """
    Find the corresponding row in the dataset based on the given percentage.
    
    :param percentage: Percentage of the day
    :param dataset: List of data rows
    :return: Index of the row to use
    """
    if not dataset:  
        return 0
    row_number = math.floor((percentage / 100) * (len(dataset) - 1))
    return max(0, min(row_number, len(dataset) - 1))

if __name__ == "__main__":
    dataset = initialize_data()

    if not dataset:
        print("No data available in the file.")
    else:
        time_input, am_pm = get_user_time()
        
        # Special case for 05:59 PM
        if time_input == "05:59" and am_pm == "PM":
            row_number = min(406, len(dataset) - 1)  # Ensure row_number is within bounds
        else:
            percentage = time_to_percentage(time_input, am_pm)
            row_number = find_row(percentage, dataset)
        
        print("Estimation of your birth time association with the 7-Day Creation week time in Genesis:")
        print("Order of line below: NIV English Translation | Verse | %Day | %Week | Hebrew Word")
        if row_number < len(dataset):  
            print(" ".join(dataset[row_number]))
        else:
            print("Invalid row number calculated.")