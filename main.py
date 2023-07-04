from pathlib import Path
from openpyxl import Workbook, load_workbook
from tkinter import Tk, Canvas, Button

OUTPUT_PATH = Path(__file__).parent

projects_file_path = "projects.xlsx"

def create_excel_file():
    if not Path(projects_file_path).is_file():
        workbook = Workbook()
        workbook.save(projects_file_path)
        print("Projects file created.")

def create_project():
    project_name = input("Enter the project name: ")
    workbook = load_workbook(projects_file_path)
    if project_name in workbook.sheetnames:
        print("Project already exists.")
        return

    workbook.create_sheet(title=project_name)
    workbook.save(projects_file_path)
    print("New project created successfully!")

def assign_employee(project_name, employee_id):
    workbook = load_workbook(projects_file_path)
    if project_name in workbook.sheetnames:
        tasks_sheet = workbook[project_name]
        tasks_sheet.append([f"ID: {employee_id}"])
        workbook.save(projects_file_path)
        print("Employee assigned to project successfully!")
    else:
        print("Project not found.")

def add_task(project_name, employee_id, task, priority):
    workbook = load_workbook(projects_file_path)
    if project_name in workbook.sheetnames:
        tasks_sheet = workbook[project_name]
        tasks_sheet.append([f"ID: {employee_id}", task, priority, "Incomplete"])
        workbook.save(projects_file_path)
        print("Task added successfully!")
    else:
        print("Project not found.")

def mark_complete(project_name, employee_id, task):
    workbook = load_workbook(projects_file_path)
    if project_name in workbook.sheetnames:
        tasks_sheet = workbook[project_name]
        for row in tasks_sheet.iter_rows(min_row=2, values_only=True):
            if row[0] == f"ID: {employee_id}" and row[1] == task:
                tasks_sheet.cell(row=row[0].row, column=4).value = "Complete"
                workbook.save(projects_file_path)
                print("Task marked as complete.")
                return
        print("Task not found.")
    else:
        print("Project not found.")

def display_tasks(project_name, employee_id):
    workbook = load_workbook(projects_file_path)
    if project_name in workbook.sheetnames:
        tasks_sheet = workbook[project_name]
        tasks = []
        for row in tasks_sheet.iter_rows(min_row=2, values_only=True):
            if row[0] == f"ID: {employee_id}":
                task = row[1]
                status = row[3]
                tasks.append(f"Task: {task}, Status: {status}")
        if tasks:
            print(f"Tasks for Employee ID {employee_id}:")
            print("\n".join(tasks))
        else:
            print("No tasks found for the employee.")
    else:
        print("Project not found.")

def create_employee(name, surname, dob):
    workbook = load_workbook(projects_file_path)
    employee_sheet = workbook.create_sheet(title="Employees")
    next_employee_id = get_next_employee_id()
    employee_sheet.append([f"ID: {next_employee_id}", name, surname, dob])
    workbook.save(projects_file_path)
    print("New employee created successfully!")

def get_next_employee_id():
    workbook = load_workbook(projects_file_path)
    employee_sheet = workbook["Employees"]
    max_id = 0
    for row in employee_sheet.iter_rows(min_row=2, values_only=True):
        employee_id = row[0].split(":")[1].strip()
        if int(employee_id) > max_id:
            max_id = int(employee_id)
    return max_id + 1

def get_all_projects():
    workbook = load_workbook(projects_file_path)
    projects = workbook.sheetnames
    projects.remove("Employees")
    return projects

def get_all_employees():
    workbook = load_workbook(projects_file_path)
    employee_sheet = workbook["Employees"]
    employees = []
    for row in employee_sheet.iter_rows(min_row=2, values_only=True):
        employee_id = row[0]
        name = row[1]
        surname = row[2]
        employees.append(f"ID: {employee_id}, Name: {name} {surname}")
    return employees

def overview_projects():
    projects = get_all_projects()
    if len(projects) > 0:
        print("Projects Overview:")
        for project in projects:
            print(f"Project: {project}")
            tasks_sheet = workbook[project]
            if tasks_sheet is not None:
                for row in tasks_sheet.iter_rows(min_row=2, values_only=True):
                    employee_id = row[0]
                    task = row[1]
                    status = row[3]
                    print(f"  - Employee ID: {employee_id}, Task: {task}, Status: {status}")
            else:
                print("  - No tasks found for this project.")
            print()
    else:
        print("No projects found.")
        
def overview_employees():
    workbook = load_workbook(projects_file_path)
    if "Employees" in workbook.sheetnames:
        employee_sheet = workbook["Employees"]
        employees = []
        for row in employee_sheet.iter_rows(min_row=2, values_only=True):
            employee_id = row[0]
            name = row[1]
            surname = row[2]
            employees.append(f"ID: {employee_id}, Name: {name} {surname}")
        if employees:
            print("Employees Overview:")
            print("\n".join(employees))
        else:
            print("No employees found.")
    else:
        print("Employees sheet not found.")

def overview_tasks():
    workbook = load_workbook(projects_file_path)
    projects = workbook.sheetnames
    if "Employees" in projects:
        projects.remove("Employees")

    if len(projects) > 0:
        print("Task Overview:")
        for project in projects:
            tasks_sheet = workbook[project]
            if tasks_sheet is not None:
                for row in tasks_sheet.iter_rows(min_row=2, values_only=True):
                    employee_id = row[0]
                    task = row[1]
                    status = row[3]
                    print(f"Project: {project}, Employee ID: {employee_id}")
                    print(f"  - Task: {task}, Status: {status}")
                    print()
            else:
                print(f"No tasks found for project: {project}")
                print()
    else:
        print("No tasks found for any employee.")

def main():
    create_excel_file()
    while True:
        print("\n=== PROJECT MANAGEMENT ===")
        print("1. Create a new project")
        print("2. Assign an employee to a project")
        print("3. Add a task for an employee")
        print("4. Mark a task as complete")
        print("5. Display tasks for an employee")
        print("6. Create a new employee")
        print("\n=== OVERVIEW SECTION ===")
        print("7. Overview of all projects")
        print("8. Overview of all employees")
        print("9. Overview of tasks for each employee")
        print("\n==========================")
        print("10. Exit")

        choice = input("Enter your choice (1-10): ")

        if choice == "1":
            create_project()
        elif choice == "2":
            project = input("Enter the project name: ")
            employee_id = input("Enter the employee ID: ")
            assign_employee(project, employee_id)
        elif choice == "3":
            project = input("Enter the project name: ")
            employee_id = input("Enter the employee ID: ")
            task = input("Enter the task: ")
            priority = input("Enter the priority: ")
            add_task(project, employee_id, task, priority)
        elif choice == "4":
            project = input("Enter the project name: ")
            employee_id = input("Enter the employee ID: ")
            task = input("Enter the task: ")
            mark_complete(project, employee_id, task)
        elif choice == "5":
            project = input("Enter the project name: ")
            employee_id = input("Enter the employee ID: ")
            display_tasks(project, employee_id)
        elif choice == "6":
            name = input("Enter the employee name: ")
            surname = input("Enter the employee surname: ")
            dob = input("Enter the employee date of birth: ")
            create_employee(name, surname, dob)
        elif choice == "7":
            overview_projects()
        elif choice == "8":
            overview_employees()
        elif choice == "9":
            overview_tasks()
        elif choice == "10":
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main()