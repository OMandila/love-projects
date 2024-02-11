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
        define = clear_sheet("define")
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
                update_sheet(row, define) # Append it as a new row in the sheet
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

    worksheet = clear_sheet("tasks")
    worksheet = SHEET.worksheet(worksheet)
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
            update_sheet(data, worksheet.title)

            another_task = input("Enter another task? Y/N : ").lower()
            if another_task == "n":
                break

        except ValueError as e:
            print(f"There was an issue: {e}")

def clear_sheet(worksheet):
    """
    Create worksheet with the same name or clear all data in the worksheet whose name is passed to the function.

    """
    try:
        SHEET.worksheet(worksheet).clear()
    except gspread.WorksheetNotFound:
        SHEET.add_worksheet(title = worksheet, rows="100", cols="10")        
    
    return SHEET.worksheet(worksheet).title

def update_sheet(data, worksheet):
    """
    Add all data passed to the function in the specified worksheet.

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

    worksheet = clear_sheet("stakeholders")
    worksheet = SHEET.worksheet(worksheet)
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
            update_sheet(data, worksheet.title)

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
    
    worksheet = clear_sheet("risks")
    worksheet = SHEET.worksheet(worksheet)
    worksheet.append_row(["Risk_title","Risk_description","Risk_probability","Risk_impact", "Risk_mitigation"])

    # Input risk details
    while True:
        print("Enter your project risks one after the other: ")
        try:
            title = input("Enter risk title: ")
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
            update_sheet(data, worksheet.title)

            another_risk = input("Enter another risk? Y/N : ").lower()
            if another_risk == "n":
                break

        except ValueError as e:
            print(f"There was an issue: {e}")

def determine_running_order():
    """
    Get details of project tasks order from the user.
    Run a while loop to collect a valid strings of data from the user via the terminal.
    The loop will repeat until the function captures all valid data.
    The function will then determine which tasks are on the critical path and display these to the user

    """
    print("Determining the running order of project tasks...")

    # Access the worksheet with tasks
    taskssheet = SHEET.worksheet("tasks")
    tasks = taskssheet.get_all_records()  # Assumes first row is header

    # Display tasks to the user
    print("Current project tasks:")
    # for task in tasks:
    task_numbers = []
    for i, task in enumerate(tasks, start=1):
        task_number = chr(64 + i) if i <= 26 else chr(64 + i // 26) + str(i % 26) # Task numbering logic to display unique numbers for up to 52 tasks
        print(f"- {task_number}: {task['Task']} FOR {task['Task_duration_weeks']} WEEKS.")
        task_numbers.append(task_number)

    # Create the TasksOrder worksheet
    worksheet = clear_sheet("taskorder")
    worksheet = SHEET.worksheet(worksheet)
    worksheet.append_row(["task_number", "task", "task_duration_weeks", "predecessors"])

    # Collect and add task order information
    for task_number, task in zip(task_numbers,tasks):
        print(f"Task number {task_number}:- {task['Task']} FOR {task['Task_duration_weeks']} WEEKS.")
        predecessors = input("Enter predecessor task numbers (separated by comma, leave blank if none): ")
        
        # Prepare data
        data = [task_number, task["Task"], task["Task_duration_weeks"], predecessors]

        # Add data to the TasksOrder worksheet
        update_sheet(data, worksheet.title)

    print("Tasks running order planned successfully!")

def determine_critical_path():
    """
    The function will determine which tasks are on the critical path and display these to the user

    """
    print("Determining project critical path...")

    def forward_pass():
        """
        calculate the forward pass data for each task
        
        """
    
    def backward_pass():
        """
        calculate the backward pass data for each task
        
        """

def develop_gantt_chart():
    """
    The function will develop a gantt chart for the project using available data
    """
    print("Developing the project gantt chart...")

def download_project():
    """
    The function will download the csv file containing the project's data

    """
    print("Downloading project data...")

def main():
    print("\nHere are the functionalities available in this app:")
    functionalities = [
        "1. Define a Project",
        "2. Develop a List of Project Stakeholders",
        "3. Develop a List of Project Tasks",
        "4. Develop a List of Project Risks",
        "5. Determine the running order for Tasks",
        "6. Calculate the Critical Path",
        "7. Develop a Project Gantt Chart",
        "8. Download Project Data"
        ]
        
    for func in functionalities:
        print(func)

    choice = int(input("Please enter the number of the function you'd like to start with (1-8): "))
        
    if choice == 1:
        define_a_project()
    elif choice == 2:
        define_project_stakeholders()
    elif choice == 3:
        define_project_tasks()
    elif choice == 4:
        define_project_risks()
    elif choice == 5:
        determine_running_order()
    elif choice == 6:
        determine_critical_path()
    elif choice == 7:
        develop_gantt_chart()
    elif choice == 8:
        # The file_id and destination_folder are predetermined
        file_id = "your_google_sheet_file_id_here"
        destination_folder = "your_destination_folder_path_here"
        download_project(file_id, destination_folder)
    else:
        print("Invalid choice. Please restart the program and select a valid option.")


print("\n\n                                          Hi, My name is Critical_Path.")
print("                                          I am your Projects Assistant.")
print("                             1.)  I will help you define and design your project.")
print("                   2.)  I will also help you determine the critical path for your project.")
print("           3.)  I will do this by brainstorming with you a set of questions to collect project data.")
print("  4.)  I will later present you the data in a csv that you can feed into your preferred project management tool.\n\n")

proceed = input("Should we proceed? Y/N: ").lower()
while True:
    if proceed != "n":
        main()
    break

