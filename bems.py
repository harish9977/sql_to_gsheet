"""
Main module to run three files
"""
import project_data
from employee_utilization_data import get_timesheet_data, transfer_data_to_google_sheets
import employee_project_billing_data


def print_menu():
    """
    function to display the options
    """
    print("Select which file to run:")
    print("1. project_data")
    print("2. employee_utilization_data")
    print("3. employee_project_billing_data")
    print("4. All files")


if __name__ == '__main__':
    while True:
        print_menu()
        option = input("Enter your choice (1/2/3/4): ")

        if option == "1":
            data_timesheet = project_data.get_timesheet_data()
            project_data.transfer_data_to_google_sheets(data_timesheet)
            print("Project data transferred to Google Sheets successfully!")
        elif option == "2":
            START_DATE = input("Enter start date (YYYY-MM-DD): ")
            END_DATE = input("Enter end date (YYYY-MM-DD): ")
            data = get_timesheet_data(START_DATE, END_DATE)
            transfer_data_to_google_sheets(data)
            print("Employee utilization data transferred to Google Sheets successfully!")
        elif option == "3":
            data_timesheet = employee_project_billing_data.get_timesheet_data()
            employee_project_billing_data.transfer_data_to_google_sheets(data_timesheet)
            print("Employee project billing data transferred to Google Sheets successfully!")
        elif option == "4":
            data_timesheet = project_data.get_timesheet_data()
            project_data.transfer_data_to_google_sheets(data_timesheet)
            print("Project data transferred to Google Sheets successfully!")
            START_DATE = input("Enter start date (YYYY-MM-DD): ")
            END_DATE = input("Enter end date (YYYY-MM-DD): ")
            data = get_timesheet_data(START_DATE, END_DATE)
            transfer_data_to_google_sheets(data)
            print("Employee utilization data transferred to Google Sheets successfully!")
            data_timesheet = employee_project_billing_data.get_timesheet_data()
            employee_project_billing_data.transfer_data_to_google_sheets(data_timesheet)
            print("Employee project billing data transferred to Google Sheets successfully!")
        else:
            print("Invalid option. Please enter 1, 2, 3 or 4.")
            continue

        choice = input("Do you want to run another file? (y/n): ")
        if choice.lower() == 'y':
            continue
        else:
            break
