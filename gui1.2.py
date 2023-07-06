#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jul  4 13:09:09 2023

@author: pietrolunati
"""
import main
import tkinter as tk
from tkinter import font
import tkinter.ttk as ttk
from ttkthemes import ThemedTk

main.create_excel_file()

def opencompletetask(proj_name):
    def subclick(proj_name):
        namestr = name.get()
        aux = main.mark_complete(proj_name, namestr)
        if aux == True:
            tk.messagebox.showinfo(title="Project Manager", message="Task marked as Complete.")
            neww.destroy()
        else:
            tk.messagebox.showinfo(title="Project Manager", message="Task not valid, retry.")
               
    neww = ThemedTk(theme="black")
    neww.configure(background="#424242", height=200, width=200)
    neww.resizable(False, False)
    neww.title("Open Project")
    
    label1 = ttk.Label(neww)
    label1.configure(
        #background="white",
        font="{Arial} 14 {}",
        text="Enter Task Name to set as Complete")
    label1.grid(padx=10, pady=10)
    
    name = ttk.Entry(neww)
    name.configure(font="{Arial} 12 {}")
    name.grid(column=0, row=1, padx=10)
    
    sub = ttk.Button(neww)
    sub.configure(text='Submit', command=lambda: subclick(proj_name))
    sub.grid(column=0, padx=10, pady=10, row=2)     

def exproj(proj_name):
    neww = ThemedTk(theme="black")
    neww.configure(background="#424242", height=200, width=200)
    neww.geometry("800x600")
    neww.resizable(False, False)
    neww.title("Project Manager")

    title = ttk.Label(neww)
    title.configure(
        #background="white",
        font="{Roboto} 48 {bold}",
        text="Project: "+ proj_name)
    title.place(anchor="nw", x=20, y=10)
    
    subtitle = ttk.Label(neww)
    subtitle.configure(
        #background="white",
        font="{Roboto} 24 {}",
        #relief="flat",
        text='Infoview and actions')
    subtitle.place(anchor="nw", x=20, y=70)

    taskadd = ttk.Button(neww)
    taskadd.configure(text="Mark Task as Complete", command=lambda: opencompletetask(proj_name))
    taskadd.place(anchor="nw", height=40, width=300, x=40, y=200)

    newproj = ttk.Button(neww)
    newproj.configure(text="Add Task", command=lambda: opentaskadd(proj_name))
    newproj.place(anchor="nw", height=40, width=300, x=40, y=150)

    close = ttk.Button(neww)
    close.configure(style="Red.TButton", text="Back", command=lambda: neww.destroy())
    close.place(anchor="nw", height=40, width=100, x=40, y=500)

    listshow = tk.Text(neww)
    listshow.configure(background="black", height=10, width=40, foreground="white",
                       font="{Roboto}")
    listshow.place(anchor="nw", height=450, x=400, y=100)
    ov_list = main.overview(proj_name)
    listshow.insert("1.0", ov_list)

    listtitle = ttk.Label(neww)
    listtitle.configure(
        #background="white",
        font="{Roboto} 16 {}",
        text='Project overview:')
    listtitle.place(anchor="nw", height=30, x=640, y=70)
    

def openexproj():
    def subclick():
        namestr = name.get()
        aux = main.list_proj()
        if namestr not in aux:
            tk.messagebox.showinfo(title="Project Manager", message="Project name invalid or project non existing!") 
        else:
            neww.destroy()
            exproj(namestr)
               
    neww = ThemedTk(theme="black")
    neww.configure(background="#424242", height=200, width=200)
    neww.resizable(False, False)
    neww.title("Open Project")
    
    label1 = ttk.Label(neww)
    label1.configure(
        #background="white",
        font="{Roboto} 14 {}",
        text="Enter Project Name")
    label1.grid(padx=10, pady=10)
    
    name = ttk.Entry(neww)
    name.configure(font="{Roboto} 12 {}")
    name.grid(column=0, row=1, padx=10)
    
    sub = ttk.Button(neww)
    sub.configure(text='Submit', command=subclick)
    sub.grid(column=0, padx=10, pady=10, row=2)  

def opennewproj():
    def subclick():
        namestr = name.get()
        flag = main.create_project(namestr)
        if flag is True:
            tk.messagebox.showinfo(title="Project Manager", message="Project created!")
            neww.destroy()
            listshow.delete("1.0", "end")
            list_proj()   
        else:
            tk.messagebox.showinfo(title="Project Manager", message="Project already exists!")
    
    neww = ThemedTk(theme="black")
    neww.configure(background="#424242", height=200, width=200)
    neww.resizable(False, False)
    neww.title("New Project")
    
    label1 = ttk.Label(neww)
    label1.configure(
        #background="white",
        font="{Roboto} 14 {}",
        text="Enter Project Name")
    label1.grid(padx=10, pady=10)
    
    name = ttk.Entry(neww)
    name.configure(font="{Roboto} 12 {}")
    name.grid(column=0, row=1, padx=10)
    
    sub = ttk.Button(neww)
    sub.configure(text='Submit', command=subclick)
    sub.grid(column=0, padx=10, pady=10, row=2)
    
def opennewempl():
    def subclick():
        namestr = name.get()
        surstr = surname.get()
        dobstr = dob.get()
        flag = main.create_employee(namestr, surstr, dobstr)
        if flag is True:
            tk.messagebox.showinfo(title="Project Manager", message="Employee created!")
        else:
            tk.messagebox.showinfo(titlw="Project Manager", message="The date of birth was not recognised.")
    
    neww = ThemedTk(theme="black")
    neww.configure(background="#424242",height=200, width=200)
    neww.resizable(False, False)
    neww.title("New Employee")
    
    label1 = ttk.Label(neww)
    label1.configure(
        #background="white",
        font="{Roboto} 14 {}",
        text="Enter Employee Name")
    label1.grid(padx=10, pady=10)
    
    name = ttk.Entry(neww)
    name.configure(font="{Graphik} 12 {}")
    name.grid(column=0, row=1, padx=10)
    
    label2 = ttk.Label(neww)
    label2.configure(
        background="#424242",
        font="{Roboto} 14 {}",
        text="Enter Employee Surname")
    label2.grid(row=2, padx=10, pady=10)
    
    surname = ttk.Entry(neww)
    surname.configure(font="{Graphik} 12 {}")
    surname.grid(column=0, row=3, padx=10)
    
    label3 = ttk.Label(neww)
    label3.configure(
        background="#424242",
        font="{Roboto} 14 {}",
        text="Enter Employee Date of Birth\n[dd.mm.yyyy]")
    label3.grid(row=4, padx=10, pady=10)
    
    dob = ttk.Entry(neww)
    dob.configure(font="{Graphik} 12 {}")
    dob.grid(column=0, row=5, padx=10)
    
    sub = ttk.Button(neww)
    sub.configure(text='Submit', command=subclick)
    sub.grid(column=0, padx=10, pady=10, row=6)
    
def opentaskadd(proj_name):
    def subclick(proj_name):
        namestr = name.get()
        text = txt.get()
        flag = main.add_task(proj_name, namestr, text)
        if flag is True:
            tk.messagebox.showinfo(title="Project Manager", message="Task created!")
        else:
            tk.messagebox.showinfo(title="Project Manager", message="Employee not found!")
    
    neww = ThemedTk(theme="black")
    neww.configure(background="#424242", height=200, width=200)
    neww.resizable(False, False)
    neww.title("New Employee")
    
    label1 = ttk.Label(neww)
    label1.configure(
        background="#424242",
        font="{Roboto} 14 {}",
        text="Enter Employee Name")
    label1.grid(padx=10, pady=10)
    
    name = ttk.Entry(neww)
    name.configure(font="{Roboto} 12 {}")
    name.grid(column=0, row=1, padx=10)
    
    label2 = ttk.Label(neww)
    label2.configure(
        #background="white",
        font="{Roboto} 14 {}",
        text="Enter Task")
    label2.grid(row=2, padx=10, pady=10)
    
    txt = ttk.Entry(neww)
    txt.configure(font="{Roboto} 12 {}")
    txt.grid(column=0, row=3, padx=10)
    
    sub = ttk.Button(neww)
    sub.configure(text='Submit', command=lambda: subclick(proj_name))
    sub.grid(column=0, padx=10, pady=10, row=6)

root = ThemedTk(theme="black")
root.configure(background="#424242", height=200, width=200)
root.geometry("800x600")
root.resizable(False, False)
root.title("Project Manager")

title = ttk.Label(root)
title.configure(
    #background="white",
    font="{Arial} 48 {bold}",
    text='Project Manager')
title.place(anchor="nw", x=20, y=10)

subtitle = ttk.Label(root)
subtitle.configure(
    #background="white",
    font="{Arial} 24 {}",
    #relief="flat",
    text="Turning Visions into Realities")
subtitle.place(anchor="nw", x=20, y=70)

newempl = ttk.Button(root)
newempl.configure(text="New Employee", command=opennewempl)
newempl.place(anchor="nw", height=40, width=300, x=40, y=200)

newproj = ttk.Button(root)
newproj.configure(text="New Project", command=opennewproj)
newproj.place(anchor="nw", height=40, width=300, x=40, y=150)

taskcheck = ttk.Button(root)
taskcheck.configure(text="Open Project", command=openexproj)
taskcheck.place(anchor="nw", height=40, width=300, x=40, y=250)

button_font = font.Font(weight="bold")
style = ttk.Style()
style.configure("Red.TButton", foreground="red", font=button_font)
close = ttk.Button(root)
close.configure(style="Red.TButton", text="Exit", command=lambda: root.quit())
close.place(anchor="nw", height=40, width=100, x=40, y=500)

listshow = tk.Text(root)
listshow.configure(background="black", height=10, width=40, foreground="white",
                   font="{Roboto}")
listshow.place(anchor="nw", height=450, x=400, y=100)
def list_proj(): 
    proj_list = main.list_proj()
    for i in proj_list:
        listshow.insert("end", " " + i + "\n")
    listshow.insert("end", "\nYou have a total of " + str(len(proj_list)) + " projects active.\n")
list_proj()

listtitle = ttk.Label(root)
listtitle.configure(
    #background="white",
    font="{Roboto} 16 {}",
    text='List of current projects:')
listtitle.place(anchor="nw", height=30, x=590, y=70)

root.mainloop()
