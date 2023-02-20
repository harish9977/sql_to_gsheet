"""
Module for calculating employee project billing data.
"""
import configparser
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
        print(f"Error connecting to MySQL: {err}")
        return None


def get_timesheet_data():
    """
    Function connects mysql server and read the data from database
    """
    try:
        db_connection = connect_to_db()
        query = "SELECT project_details.project_name AS " \
                "Project_name,concat(user_details.first_name,' ',user_details.last_name) " \
                "AS Employee_name,sum(TIME_TO_SEC(timesheet.task_hours))/3600 AS " \
                "Hours_logged_for_billable_utilization_for_project," \
                "sum(TIME_TO_SEC(timesheet.task_hours))/3600 * project_user_mapping.hourly_rate " \
                "AS Billing_rate_for_employee," \
                "sum(sum(TIME_TO_SEC(timesheet.task_hours))/3600) OVER " \
                "(PARTITION BY project_details.project_id) " \
                "AS Total_hours_logged_to_project FROM project_details JOIN timesheet ON " \
                "project_details.project_id = timesheet.project_id JOIN project_user_mapping ON " \
                "timesheet.project_id = project_user_mapping.project_id JOIN user_details ON " \
                "timesheet.user_id = user_details.user_id WHERE project_details.billable = 1 " \
                "GROUP BY project_details.project_id, user_details.user_id"
        timesheet_data = pd.read_sql(query, db_connection)
        return timesheet_data
    except mysql.connector.Error as err:
        print(f"Error getting timesheet data: {err}")
        return None


# Function to transfer the data from mysql data to Googlesheets
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
        worksheet_name = config.get(SPREADSHEET_SECTION, 'worksheet_name_2')
        sheet = google_sheets_client.open(spreadsheet_name)
        worksheet = sheet.worksheet_by_title(worksheet_name)

        # Clear any existing data and write the new data
        worksheet.clear()
        worksheet.set_dataframe(timesheet_data, (1, 1), extend=True)
    except pygsheets.AuthenticationError as err:
        print(f"Error authenticating with Google Sheets API: {err}")
    except pygsheets.WorksheetNotFound as err:
        print(f"Error finding worksheet in Google Sheets: {err}")


if __name__ == '__main__':
    # Fetch the timesheet data and transfer it to Google Sheets
    data_timesheet = get_timesheet_data()
    transfer_data_to_google_sheets(data_timesheet)
