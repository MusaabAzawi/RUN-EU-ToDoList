from openpyxl import Workbook, load_workbook
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

# Specify the file paths
projects_file_path = r"C:\Users\DHardi\OneDrive - Technological University of the Shannon Midwest\Organisations\RUN-EU\Python\Work group\RUN-EU-ToDoList\projects.xlsx"
employees_file_path = r"C:\Users\DHardi\OneDrive - Technological University of the Shannon Midwest\Organisations\RUN-EU\Python\Work group\RUN-EU-ToDoList\employees.xlsx"

# Function to create a new project
def create_project():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Projects"
    sheet.append(["Project Name", "Manager"])
    workbook.save(projects_file_path)
    messagebox.showinfo("Success", "New project created successfully!")

# Function to assign an employee to a project
def assign_employee(project, employee_id):
    workbook = load_workbook(projects_file_path)
    sheet = workbook["Projects"]
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] == project:
            if employee_id not in row[2:]:
                row = list(row) + [employee_id]
                sheet.append(row)
                workbook.save(projects_file_path)
                messagebox.showinfo("Success", "Employee assigned to project.")
                return
            else:
                messagebox.showwarning("Duplicate Assignment", "Employee already assigned to project.")
                return
    messagebox.showwarning("Project not found", "Project not found.")

# Function to add a task for an employee
def add_task(project, employee_id, task, priority):
    workbook = load_workbook(projects_file_path)
    sheet = workbook["Projects"]
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] == project and employee_id in row[2:]:
            project_col = row[0].column
            employee_col = row.index(employee_id)
            tasks_sheet = workbook.get_sheet_by_name(project)
            if tasks_sheet is None:
                tasks_sheet = workbook.create_sheet(title=project)
                tasks_sheet.append(["Employee ID", "Task", "Priority", "Status"])
            tasks_sheet.append([employee_id, task, priority, "Incomplete"])
            workbook.save(projects_file_path)
            messagebox.showinfo("Success", "Task added successfully.")
            return
    messagebox.showwarning("Assignment not found", "Assignment not found.")

# Function to mark a task as complete
def mark_complete(project, employee_id, task):
    workbook = load_workbook(projects_file_path)
    sheet = workbook["Projects"]
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] == project and employee_id in row[2:]:
            tasks_sheet = workbook.get_sheet_by_name(project)
            if tasks_sheet is not None:
                for task_row in tasks_sheet.iter_rows(min_row=2, values_only=True):
                    if task_row[0] == employee_id and task_row[1] == task:
                        tasks_sheet.cell(row=task_row[0].row, column=4).value = "Complete"
                        workbook.save(projects_file_path)
                        messagebox.showinfo("Success", "Task marked as complete.")
                        return
                messagebox.showwarning("Task not found", "Task not found.")
                return
    messagebox.showwarning("Assignment not found", "Assignment not found.")

# Function to display the tasks for an employee
def display_tasks(project, employee_id):
    workbook = load_workbook(projects_file_path)
    tasks_sheet = workbook.get_sheet_by_name(project)
    if tasks_sheet is not None:
        window = tk.Tk()
        window.title("Tasks for Employee")

        # Create a listbox to display tasks
        task_listbox = tk.Listbox(window)
        task_listbox.pack(padx=10, pady=10)

        for row in tasks_sheet.iter_rows(min_row=2, values_only=True):
            if row[0] == employee_id:
                task_listbox.insert(tk.END, f"Task: {row[1]}, Priority: {row[2]}, Status: {row[3]}")

        window.mainloop()
    else:
        messagebox.showwarning("No tasks found", "No tasks found for the employee.")

# Function to create a new employee
def create_employee(name, surname, dob):
    workbook = load_workbook(employees_file_path)
    sheet = workbook.active
    employee_id = sheet.max_row + 1
    sheet.append([employee_id, name, surname, dob])
    workbook.save(employees_file_path)
    messagebox.showinfo("Success", f"New employee created successfully with ID: {employee_id}")

# Function to get a list of all projects
def get_all_projects():
    workbook = load_workbook(projects_file_path)
    sheet = workbook["Projects"]
    projects = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        projects.append(row[0])
    return projects

# Function to display an overview of all projects
def overview_projects():
    projects = get_all_projects()
    if len(projects) > 0:
        messagebox.showinfo("Projects Overview", "\n".join(projects))
    else:
        messagebox.showwarning("No projects found", "No projects found.")

# Function to get a list of all employees
def get_all_employees():
    workbook = load_workbook(employees_file_path)
    sheet = workbook.active
    employees = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        employees.append(f"ID: {row[0]}, Name: {row[1]} {row[2]}, DOB: {row[3].strftime('%d.%m.%Y')}")
    return employees

# Function to display an overview of all employees
def overview_employees():
    employees = get_all_employees()
    if len(employees) > 0:
        messagebox.showinfo("Employees Overview", "\n".join(employees))
    else:
        messagebox.showwarning("No employees found", "No employees found.")

# Function to display an overview of tasks for each employee
def overview_tasks():
    workbook = load_workbook(projects_file_path)
    projects = workbook.sheetnames
    if "Projects" in projects:
        projects.remove("Projects")

    if len(projects) > 0:
        window = tk.Tk()
        window.title("Task Overview")

        # Create a listbox to display employees
        employee_listbox = tk.Listbox(window)
        employee_listbox.pack(padx=10, pady=10)

        # Populate employee listbox
        for project in projects:
            tasks_sheet = workbook.get_sheet_by_name(project)
            if tasks_sheet is not None:
                for row in tasks_sheet.iter_rows(min_row=2, values_only=True):
                    employee_id = row[0]
                    employee_listbox.insert(tk.END, f"Project: {project}, Employee ID: {employee_id}")

        # Function to display tasks for the selected employee
        def show_tasks():
            selected_employee = employee_listbox.get(employee_listbox.curselection())
            project, employee_id = selected_employee.split(", ")
            display_tasks(project.split(": ")[1], int(employee_id.split(": ")[1]))

        # Create a button to show tasks
        show_button = tk.Button(window, text="Show Tasks", command=show_tasks)
        show_button.pack(pady=10)

        window.mainloop()
    else:
        messagebox.showwarning("No tasks found", "No tasks found for any employee.")

# Main function to handle user input and menu options
def main():
    while True:
        print("\n=== PROJECT MANAGEMENT ===")
        print("1. Create a new project")
        print("2. Assign an employee to a project")
        print("3. Add a task for an employee")
        print("4. Mark a task as complete")
        print("5. Display tasks for an employee")
        print("6. Create a new employee")
        print("7. Overview of all projects")
        print("8. Overview of all employees")
        print("9. Overview of tasks for each employee")
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
            task = input("Enter the task to mark as complete: ")
            mark_complete(project, employee_id, task)
        elif choice == "5":
            project = input("Enter the project name: ")
            employee_id = input("Enter the employee ID: ")
            display_tasks(project, employee_id)
        elif choice == "6":
            name = input("Enter the employee's name: ")
            surname = input("Enter the employee's surname: ")
            dob = input("Enter the employee's date of birth (dd.mm.yyyy): ")
            create_employee(name, surname, dob)
        elif choice == "7":
            overview_projects()
        elif choice == "8":
            overview_employees()
        elif choice == "9":
            overview_tasks()
        elif choice == "10":
            print("Exiting program...")
            break
        else:
            print("Invalid choice. Please try again.")

        # Prompt to continue or exit
        continue_choice = input("Do you want to continue? (Y/N): ")
        if continue_choice.lower() != "y":
            print("Exiting program...")
            break

if __name__ == "__main__":
    main()
