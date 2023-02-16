"""
Main module to run three files
"""
import project_data
from employee_utilization_data import get_timesheet_data, transfer_data_to_google_sheets
import employee_project_billing_data

if __name__ == '__main__':
    data_timesheet = project_data.get_timesheet_data()
    project_data.transfer_data_to_google_sheets(data_timesheet)

    # employee_utilization_data

    START_DATE = '2022-01-01'
    END_DATE = '2022-01-31'

    data = get_timesheet_data(START_DATE, END_DATE)
    transfer_data_to_google_sheets(data)

    # employee_project_billing_data

    data_timesheet = employee_project_billing_data.get_timesheet_data()
    employee_project_billing_data.transfer_data_to_google_sheets(data_timesheet)
