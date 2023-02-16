"""
Module for calculating project data.
"""
import configparser
import pandas as pd
import pygsheets
import mysql.connector

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
    db_config = {
        'host': config.get(DB_SECTION, 'host'),
        'user': config.get(DB_SECTION, 'user'),
        'password': config.get(DB_SECTION, 'password'),
        'database': config.get(DB_SECTION, 'database')
    }
    return mysql.connector.connect(**db_config)


# Function to get timesheet data from database
def get_timesheet_data():
    """
    Get timesheet data from MySQL database and return as a Pandas DataFrame.
    """
    db_connection = connect_to_db()
    query = "select pd.Project_name ,pd.Sow_id,pd.Total_hours," \
            "concat(ud.first_name,' ',ud.last_name) AS Manager,pd.Start_date,pd.End_date," \
            "pd.project_classification,pd.budget as Budget,pd.billable " \
            "as Billable,pd.utilization as Utilization," \
            "(SELECT (sum(TIME_TO_SEC(ts.task_hours)))/3600 from timesheet" \
            " ts where ts.project_id=pd.project_id) as utilizedHours " \
            "from project_details pd , user_details ud, project_user_mapping pum where " \
            "pd.project_id=pum.project_id and pum.user_id = ud.user_id and pum.manager = 1  " \
            "ORDER BY pd.project_name"
    timesheet_data = pd.read_sql(query, db_connection)
    return timesheet_data


# Function to transfer the data to Google Sheets
def transfer_data_to_google_sheets(timesheet_data):
    """
    Transfer the data from Pandas DataFrame to Google Sheets using configuration from INI file.
    """
    # Get the path to the service account JSON file from the INI file
    google_json_path = config.get(GOOGLE_SECTION, 'service_account_json_path')

    # Authorize access to the Google Sheets API
    google_sheets_client = pygsheets.authorize(service_file=google_json_path)

    # Open the spreadsheet and worksheet
    spreadsheet_name = config.get(SPREADSHEET_SECTION, 'spreadsheet_name')
    worksheet_name = config.get(SPREADSHEET_SECTION, 'worksheet_name')
    sheet = google_sheets_client.open(spreadsheet_name)
    worksheet = sheet.worksheet_by_title(worksheet_name)

    # Clear any existing data and write the new data
    worksheet.clear()
    worksheet.set_dataframe(timesheet_data, (1, 1), extend=True)


if __name__ == '__main__':
    data_timesheet = get_timesheet_data()
    transfer_data_to_google_sheets(data_timesheet)
