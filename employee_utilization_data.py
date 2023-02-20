"""
Module for calculating employee utilization data.
"""
import sys
import configparser
import datetime
import mysql.connector
import pandas as pd
import pygsheets

# Define the path to the INI file
CONFIG_PATH = 'config.ini'

# Load the configuration from the INI file
config = configparser.ConfigParser()
config.read(CONFIG_PATH)

# Define the section names in the INI file
DB_SECTION = 'DATABASE'
GOOGLE_SECTION = 'GOOGLE'
SPREADSHEET_SECTION = 'SPREADSHEET'


# Function to connect to MySQL database
def connect_to_db():
    """
    Connect to MySQL database using configuration from INI file.
    """
    try:
        db_config = {
            'host': config.get(DB_SECTION, 'host'),
            'user': config.get(DB_SECTION, 'user'),
            'password': config.get(DB_SECTION, 'password'),
            'database': config.get(DB_SECTION, 'database')
        }
        return mysql.connector.connect(**db_config)
    except mysql.connector.Error as err:
        print(f"Error connecting to mysql: {err}")
        return None


# Function connects mysql server and read the data from database
def get_timesheet_data(start_date, end_date):
    """
    Function connects mysql server and read the data from database
    """
    try:
        db_connection = connect_to_db()
        query = f"SELECT concat(UD.first_name,' ' ,UD.last_name) " \
                f"AS Employee_Name,COALESCE((SELECT SUM(TIME_TO_SEC(T.task_hours))/3600 " \
                f"FROM timesheet T WHERE T.user_id = UD.user_id  and T.task_date " \
                f"BETWEEN '{start_date}' AND '{end_date}' and T.project_id IN " \
                f"(select project_id from project_details where billable = 1)), '')  " \
                f"AS Hours_logged_to_Billable_utilization," \
                f"COALESCE((SELECT SUM(TIME_TO_SEC(T.task_hours))/3600 " \
                f"FROM timesheet T WHERE T.user_id = UD.user_id  and T.task_date " \
                f"BETWEEN '{start_date}' AND '{end_date}' and T.project_id IN " \
                f"(select project_id from project_details where " \
                f"utilization = 1 AND billable=0)), '') " \
                f"AS Hours_logged_to_Non_Billable_utilization," \
                f"COALESCE((SELECT MIN(T.task_date) FROM timesheet T " \
                f"WHERE T.user_id = UD.user_id and T.task_date " \
                f"BETWEEN '{start_date}' AND '{end_date}'), '') AS Date_Range_From," \
                f"COALESCE((SELECT MAX(T.task_date) FROM timesheet T " \
                f"WHERE T.user_id = UD.user_id and T.task_date " \
                f"BETWEEN '{start_date}' AND '{end_date}'), '') as Date_Range_To," \
                f"COALESCE((SELECT SUM(TIME_TO_SEC(T.calc_allocated_hours)) " \
                f"FROM timesheet T WHERE T.user_id = UD.user_id and T.status=1 " \
                f"and T.task_date BETWEEN '{start_date}' AND '{end_date}'), '') " \
                f"AS Calc_Allocated_Hours FROM user_details UD  " \
                f"WHERE UD.user_id IN (SELECT user_id FROM user_details) and " \
                f"UD.active = 1".format(start_date, end_date, start_date,
                                        end_date, start_date, end_date, start_date, end_date)

        timesheet_data = pd.read_sql(query, db_connection)
        return timesheet_data
    except mysql.connector.Error as err:
        print(f"Error getting timesheet data: {err}")
        return None


# Function to transfer the data to Google Sheets
def transfer_data_to_google_sheets(timesheet_data):
    """
    Transfer the data from Pandas DataFrame to Google Sheets using configuration from INI file.
    """
    try:
        # Get the path to the service account JSON file from the INI file
        google_json_path = config.get(GOOGLE_SECTION, 'service_account_json_path')

        # Authorize access to the Google Sheets API
        google_sheets_client = pygsheets.authorize(service_file=google_json_path)

        # Open the spreadsheet and worksheet
        spreadsheet_name = config.get(SPREADSHEET_SECTION, 'spreadsheet_name')
        worksheet_name = config.get(SPREADSHEET_SECTION, 'worksheet_name_1')
        sheet = google_sheets_client.open(spreadsheet_name)
        worksheet = sheet.worksheet_by_title(worksheet_name)

        # Clear any existing data and write the new data
        worksheet.clear()
        worksheet.set_dataframe(timesheet_data, (1, 1), extend=True)
    except pygsheets.AuthenticationError as err:
        print(f"Error authenticating with Google Sheets API: {err}")
    except pygsheets.WorksheetNotFound as err:
        print(f"Error finding worksheet in Google Sheets: {err}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Please provide exactly two arguments:start date and end date in (YYYY-MM-DD) format")
        sys.exit()

    START_DATE = sys.argv[1]
    END_DATE = sys.argv[2]

    # Check that the start and end dates are in the correct format (YYYY-MM-DD)
    try:
        datetime.datetime.strptime(START_DATE, '%Y-%m-%d')
    except ValueError:
        print("Please provide start date in the format YYYY-MM-DD.")
        sys.exit()

    try:
        datetime.datetime.strptime(END_DATE, '%Y-%m-%d')
    except ValueError:
        print("Please provide end date in the format YYYY-MM-DD.")
        sys.exit()

    data_timesheet = get_timesheet_data(START_DATE, END_DATE)
    transfer_data_to_google_sheets(data_timesheet)
