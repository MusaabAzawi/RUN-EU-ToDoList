from openpyxl import Workbook, load_workbook
import tkinter as tk
from tkinter import messagebox

# Function to create a new To-Do list
def create_todo_list():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "To-Do List"
    sheet.append(["Task", "Priority", "Status"])
    workbook.save("todo_list.xlsx")
    messagebox.showinfo("Success", "New To-Do list created successfully!")

# Function to add a task to the To-Do list
def add_task(task, priority):
    workbook = load_workbook("todo_list.xlsx")
    sheet = workbook["To-Do List"]
    sheet.append([task, priority, "Incomplete"])
    workbook.save("todo_list.xlsx")
    messagebox.showinfo("Success", "Task added successfully!")

# Function to mark a task as complete
def mark_complete(task):
    workbook = load_workbook("todo_list.xlsx")
    sheet = workbook["To-Do List"]
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] == task:
            sheet.cell(row=row[0].row, column=3).value = "Complete"
            workbook.save("todo_list.xlsx")
            messagebox.showinfo("Success", "Task marked as complete.")
            return
    messagebox.showwarning("Task not found", "Task not found.")

# Function to display the To-Do list in a graphical window
def display_todo_list():
    workbook = load_workbook("todo_list.xlsx")
    sheet = workbook["To-Do List"]

    window = tk.Tk()
    window.title("To-Do List")

    # Create a listbox to display tasks
    task_listbox = tk.Listbox(window)
    task_listbox.pack(padx=10, pady=10)

    for row in sheet.iter_rows(min_row=2, values_only=True):
        task_listbox.insert(tk.END, f"Task: {row[0]}, Priority: {row[1]}, Status: {row[2]}")

    window.mainloop()

# Main function to handle user input and menu options
def main():
    while True:
        print("\n=== TO-DO LIST ===")
        print("1. Create a new To-Do list")
        print("2. Add a task")
        print("3. Mark a task as complete")
        print("4. Display the To-Do list")
        print("5. Exit")

        choice = input("Enter your choice (1-5): ")

        if choice == "1":
            create_todo_list()
        elif choice == "2":
            task = input("Enter the task: ")
            priority = input("Enter the priority: ")
            add_task(task, priority)
        elif choice == "3":
            task = input("Enter the task to mark as complete: ")
            mark_complete(task)
        elif choice == "4":
            display_todo_list()
        elif choice == "5":
            print("Exiting program...")
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main()