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
    Run a while loop to collect a valid strings of data from the user via the terminal.
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
                print("The Project budget is required.")
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
        define.clear() # Clear the entire sheet
        data.append({"project_name": project_name, 
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
        print(f"Failed to add Project details to the Google Sheet: {e}")

def define_project_tasks():
    """
    Get details of project tasks from the user.
    Run a while loop to collect a valid strings of data from the user via the terminal.
    The loop will repeat until the function captures all valid data.

    """
    print("Defining project tasks...")

    worksheet = SHEET.worksheet("tasks")
    worksheet.clear() # Clear the entire sheet
    worksheet.append_row(["Task","Task_lead","Task_duration_weeks"])

    # Input task details
    while True:
        print("Enter your project tasks one after the other: ")
        try:
            task = input("Enter task name: ")
            if not task:
                print("Task name is required.")
                continue
            else: 
                task_lead = input("Enter first name of task lead: ")
                if not task_lead:
                    print("Task lead first name is required")
                    continue
                else:
                    estimate_weeks_per_task = int(input("Enter estimate weeks for the task: "))
                    if not estimate_weeks_per_task:
                        print("Estimation of task duration is required")
                        continue
            data = [task,task_lead,estimate_weeks_per_task]
            update_sheet(data, "tasks")

            another_task = input("Enter another task? Y/N : ").lower()
            if another_task == "n":
                break

        except ValueError as e:
            print(f"There was an issue: {e}")

def update_sheet(data, worksheet):
    """
    Add all data parsed to the function in the specified worksheet.

    """
    # Access the worksheet and add data
    try:
        worksheet = SHEET.worksheet(worksheet)
        worksheet.append_row(data) # Append it as a new row in the sheet
        print("Data added successfully!")
    except Exception as e:
        print(f"Failed to add task details to the Google Sheet: {e}")

def define_project_stakeholders():
    """
    Get details of project stakeholders from the user.
    Run a while loop to collect a valid strings of data from the user via the terminal.
    The loop will repeat until the function captures all valid data.

    """
    print("Defining project stakeholders...")

    worksheet = SHEET.worksheet("stakeholders")
    worksheet.clear() # Clear the entire sheet
    worksheet.append_row(["Stakeholder_name","Stakeholder_role","Influence_to_project","Interest_in_project"])

    # Input stakeholder details
    while True:
        print("Enter your project stakeholders one after the other: ")
        try:
            name = input("Enter stakeholder's first name: ")
            if not name:
                print("Stakeholder name is required.")
                continue
            else: 
                role = input("Enter title or role of the stakeholder: ")
                if not role:
                    print("Stakeholder's title or role is required")
                    continue
                else:
                    influence = int(input("Enter as a number the influence of the stakeholder to the project (3-high, 2-medium, 1-low): "))
                    if not influence:
                        print("The stakeholder's influence to the project is required")
                        continue
                    else:
                        if influence not in [1, 2, 3]:
                            print("The stakeholder's influence to the project is required as a number (1, 2, or 3).")
                            continue
                        else:
                            interest = int(input("Enter as a number the interest of the stakeholder to the project (3-high, 2-medium, 1-low): "))
                            if not interest:
                                print("The stakeholder's interest to the project is required")
                                continue
                            else:
                                if interest not in [1, 2, 3]:
                                    print("The stakeholder's interest to the project is required as a number (1, 2, or 3).")
                                    continue

            data = [name,role,influence,interest]
            update_sheet(data, "stakeholders")

            another_stakeholder = input("Enter another stakeholder? Y/N : ").lower()
            if another_stakeholder == "n":
                break

        except ValueError as e:
            print(f"There was an issue: {e}")

def define_project_risks():
    """
    Get details of project risks from the user.
    Run a while loop to collect a valid strings of data from the user via the terminal.
    The loop will repeat until the function captures all valid data.

    """
    print("Defining project risks...")
    
    worksheet = SHEET.worksheet("risks")
    worksheet.clear() # Clear the entire sheet
    worksheet.append_row(["Risk_title","Risk_description","Risk_probability","Risk_impact", "Risk_mitigation"])

    # Input risk details
    while True:
        print("Enter your project risks one after the other: ")
        try:
            title = input("Enter risk's title: ")
            if not title:
                print("Risk title is required.")
                continue
            else: 
                description = input("Enter detailed description of the risk: ")
                if not description:
                    print("Risk description is required")
                    continue
                else:
                    probability = int(input("Enter as a number the probability of the risk to the project (3-high, 2-medium, 1-low): "))
                    if not probability:
                        print("The risk's probability to the project is required")
                        continue
                    else:
                        if probability not in [1, 2, 3]:
                            print("The risk's probability to the project is required as a number (1, 2, or 3).")
                            continue
                        else:
                            impact = int(input("Enter as a number the impact of the risk to the project (3-high, 2-medium, 1-low): "))
                            if not impact:
                                print("The risk's impact to the project is required")
                                continue
                            else:
                                if impact not in [1, 2, 3]:
                                    print("The risk's impact to the project is required as a number (1, 2, or 3).")
                                    continue

            data = [title,description,probability,impact]
            update_sheet(data, "risks")

            another_risk = input("Enter another risk? Y/N : ").lower()
            if another_risk == "n":
                break

        except ValueError as e:
            print(f"There was an issue: {e}")

def determine_critical_path():
    """
    Get details of project tasks order from the user.
    Run a while loop to collect a valid strings of data from the user via the terminal.
    The loop will repeat until the function captures all valid data.
    The function will then determine which tasks are on the critical path and display these to the user

    """
    print("Determining project critical path...")

def main():
    # define_a_project()
    # define_project_stakeholders()
    # define_project_tasks()
    define_project_risks()
    # determine_critical_path()

print("\n\n                                          Hi, My name is Critical_Path.")
print("                                          I am your Projects Assistant.")
print("                             1.)  I will help you define and design your project.")
print("                   2.)  I will also help you determine the critical path for your project.")
print("           3.)  I will do this by brainstorming with you a set of questions to collect project data.")
print("  4.)  I will later present you the data in a csv that you can feed into your preferred project management tool.\n\n")

proceed = input("Sounds good? Y/N: ").lower()
while True:
    if proceed != "n":
        main()
    break

