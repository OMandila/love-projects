import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('love_projects')

def define_a_project():
    """
    Get details of a new project from the user.
    Run a while loop to collect a valid strings of data from the user
    via the terminal.
    The loop will repeat until the function captures all valid data.

    """
    print("Defining a new project...")
    data = []

    while True:
        try:
            project_name = input("Enter project name: ")
            description = input("Enter project description: ")
            if not project_name or not description: # This checks for an empty string
                raise ValueError("Project name and Project description are both required.")
            else:
                break
        except ValueError as e:
            print(e)
    
    # Input and validation for start date
    while True:
        start_date_str = input("Enter start date (YYYY-MM-DD): ")
        try:
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            break  # Exit the loop if the date is correctly formatted
        except ValueError:
            print("Invalid date format. Please use YYYY-MM-DD.")
    
    # Input and validation for finish date
    while True:
        finish_date_str = input("Enter finish date (YYYY-MM-DD): ")
        try:
            finish_date = datetime.strptime(finish_date_str, '%Y-%m-%d')
            if finish_date < start_date:
                print("Finish date cannot be earlier than start date. Please enter a valid finish date.")
                continue
            break  # Exit the loop if the date is correctly formatted and valid
        except ValueError:
            print("Invalid date format. Please use YYYY-MM-DD.")

    # Input and validation for budget data
    while True:
        try:
            available_budget = int(input("Enter available budget for the project in EUR: "))
            if not available_budget: # This checks for an empty string
                raise ValueError("The Project budget is required.")
            else:
                break
        except ValueError as e:
            print(e)

    # Calculate the number of weeks between start and finish dates
    number_of_days = (finish_date - start_date).days
    number_of_weeks = number_of_days // 7
    
    # Access the worksheet and add project details
    try:
        define = SHEET.worksheet("define")
        # Clear the entire sheet
        define.clear()
        data.append({"Attribute":"Value",
                     "project_name": project_name, 
                     "project_description": description, 
                     "start_date": start_date_str, 
                     "finish_date": finish_date_str, 
                     "project-budget": available_budget, 
                     "total-project-weeks": str(number_of_weeks)})
        for i in data:
            for key, value in i.items():
                # If value is a list, convert it to a string
                if isinstance(value, list):
                    value = ','.join(map(str, value))
                row = [key, value]  # Create a list from the key-value pair
                define.append_row(row) # Append it as a new row in the sheet
        print("Project defined successfully!")
    except Exception as e:
        print(f"Failed to add project details to Google Sheet: {e}")

define_a_project()