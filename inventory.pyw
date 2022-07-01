# Imports
from distutils import command
from email import header
from fileinput import filename
from getpass import getpass
from multiprocessing import connection
from re import L
from sys import stderr, stdout
import tkinter
from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
from tkinter import messagebox
import sqlite3
import os
from subprocess import Popen
import threading
import smtplib
from email import message
from email.message import EmailMessage
from email import message
import datetime
from datetime import date, time, datetime
import getpass

# Connect to database and create cursor
conn = sqlite3.connect('treebase.db')
c = conn.cursor()
# Create a table
# c.execute("""CREATE TABLE if not exists items_1 (
#         device_type TEXT NOT NULL,
#        quantity INTEGER,
#        id INTEGER)""")


root = Tk()
root.title("Inventory Management")
root.iconbitmap('example.ico')
root.geometry('1000x500')

# Add some style
style = ttk.Style()
# Pick a theme
style.theme_use('default')
# Configure tree-view colors
style.configure("Treeview",
                background="#D3D3D3",
                foreground="black",
                rowheight=25,
                fieldbackground="#D3D3D3")
# Change selected color
style.map('Treeview',
          background=[('selected', "#347083")])
# Create a treeview frame
tree_frame = Frame(root)
tree_frame.pack(fill ="x", padx=100, pady=10)

# Create a treeview scrollbar
tree_scroll = Scrollbar(tree_frame)
tree_scroll.pack(side=RIGHT, fill=Y)

# Create the treeview
my_tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set, selectmode="extended")
my_tree.pack(fill="x")

# Configure the scrollbar
tree_scroll.config(command=my_tree.yview())
yscrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=my_tree.yview)
yscrollbar.pack(side=RIGHT, fill="y")


# Define the columns
my_tree['columns'] = ('Device Type', 'Quantity', 'ID')

# Format the columns
my_tree.column("#0", width=0, stretch=NO)
my_tree.column("Device Type", anchor=W, width=140)
my_tree.column("Quantity", anchor=CENTER, width=140)
my_tree.column("ID", anchor=CENTER, width=100)

# Create Headings
my_tree.heading("#0", text="", anchor=W)
my_tree.heading("Device Type", text="Device Type", anchor=W)
my_tree.heading("Quantity", text="Quantity", anchor=CENTER)
my_tree.heading("ID", text="ID", anchor=CENTER)

# Create striped rows
my_tree.tag_configure('oddrow', background="white")
my_tree.tag_configure('evenrow', background="lightblue")

# Add record entry boxes
data_frame = LabelFrame(root, text="Record:")
data_frame.pack(fill="x", expand="yes", padx=20, pady=25)

d_type = Label(data_frame, text="Device Type:")
d_type.grid(row=0, column=0, padx=10, pady=10)
d_type = Entry(data_frame)
d_type.grid(row=0, column=1, padx=10, pady=10)

quan=Label(data_frame, text="Quantity:")
quan.grid(row=0, column=2, padx=10, pady=10)
quan = Entry(data_frame)
quan.grid(row=0, column=3, padx=10, pady=10)

id=Label(data_frame, text="ID:")
id.grid(row=0, column=4, padx=10, pady=10)
id = Entry(data_frame)
id.grid(row=0, column=5, padx=10, pady=10)

#  Functions

def update():
    # Grab the record number
    selected = my_tree.focus()
    my_tree.item(selected, text="", values=(d_type.get(), quan.get(), id.get() ))

    # update database
    sqlite3.connect('treebase.db')
    c = conn.cursor()

    c.execute("""UPDATE items_1 SET
    device_type = :device,
    quantity = :quan,
    id = :iden

    WHERE oid = :oid""",
              {
                'device': d_type.get(),
                'quan': quan.get(),
                'iden': id.get(),
                'oid': id.get(),
              })

    #messagebox.showinfo("Updated", "Updated Record.")
    refresh()
    conn.commit()
    conn.close()



# Query Function
def query_db():
    c.execute("SELECT rowid, * FROM items_1")
    records = c.fetchall()
    conn.commit()
    

    # Add data to screen
    global count
    count = 0

    for record in records:
        if count % 2 == 0:
            my_tree.insert(parent='', index='end', iid=count, text='', values=(record[1], record[2], record[3]),
                           tags=('evenrow'))
        else:
            my_tree.insert(parent='', index='end', iid=count, text='', values=(record[1], record[2], record[3]),
                           tags=('oddrow'))
            # increment counter
        count += 1



# Remove Record

# Add a new record
def add_new():
    c.execute("INSERT INTO items_1 VALUES (:device, :quan, :iden)",
              {
                  'device': d_type.get(),
                  'quan': quan.get(),
                  'iden': id.get()
              }
              )
    conn.commit()
    conn.close()

    # Clear the text boxes
    d_type.delete(0, END)
    quan.delete(0, END)
    id.delete(0, END)

    # Clear the Treeview table
    my_tree.delete(*my_tree.get_children())
    messagebox.showinfo("Added", "Record has been added.")
    query_db()
    refresh()
    

def confirm():
    # confirm action

    answer = messagebox.askyesno(title='Confrim', message='Are you sure you want to delete this record?')

    if answer:
        remove()


# Delete a record
def remove():

    # Remove from treeview
    x = my_tree.selection()[0]
    my_tree.delete(x)

    # Delete from database
    c.execute("""DELETE FROM items_1 WHERE oid = :oid""",
              {
                'device': d_type.get(),
                'quan': quan.get(),
                'iden': id.get(),
                'oid': id.get(),
              })
    conn.commit()
    conn.commit()
    messagebox.showinfo("Delete", "Record has been deleted.")
    refresh()
    cler_entries1()
    query_db()
    conn.commit()
    conn.close()


# Move Row Up
def up():
    rows = my_tree.selection()
    for row in rows:
        my_tree.move(row, my_tree.parent(row), my_tree.index(row)-1)

# Move Row Down
def down():
    rows = my_tree.selection()
    for row in reversed(rows):
        my_tree.move(row, my_tree.parent(row), my_tree.index(row)+1)



# Clear Entry Boxes
def cler_entries1():
    d_type.delete(0, END)
    quan.delete(0, END)
    id.delete(0, END)



# Select Record
def select_record(e):
    d_type.delete(0, END)
    quan.delete(0, END)
    id.delete(0, END)

    selected = my_tree.focus()
    values_1 = my_tree.item(selected, 'values')

    d_type.insert(0, values_1[0])
    quan.insert(0, values_1[1])
    id.insert(0, values_1[2])

# Quit Program
def quit():
    conf_quit = messagebox.askyesno(title='Exit', message='Are you sure you want to exit?')
    conf_quit
    if conf_quit:
        quit_1()

def quit_1():
    root.destroy()


# Refreshes the program
def refresh():
    root.withdraw()
    root.update()
    #os.popen("Inventory_2.pyw")


# Prints current items/quantity in db to the report.csv excel file.
def print_report():
    f = open("report.csv", "w+")
    f.truncate(0)
    f.close()
    refresh()
    p = Popen("cmnd.bat", cwd="CWD(change)")
    stdout, stderr = p.communicate() 
    messagebox.showinfo("Print Report", "Report has been printed to .csv file.")


# Sends report.csv to Destination Email
def send_message():

    # Variables
    CWD = (CWD Change)
    today = date.today()
    EMAIL_ADDRESS = "example@example.com"
    EMAIL_PASSWORD = "examplepass"
    BODY = ("Attached is the current Inventory Report")
    
    msg = EmailMessage()
    msg['Subject'] = 'Inventory Report for {}'.format(today)
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = 'destinationexampleemail@example.com'
    msg.set_content(BODY)
    
    with open(CWD, 'rb') as f:
        file_data = f.read()
        msg.add_attachment(file_data, maintype='application', subtype='csv', filename="report.csv")



    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)
        smtp.close()
    
    messagebox.showinfo("Email Report", "Report has been Emailed.")





    # Add buttons
button_frame = LabelFrame(root, text="Commands:")
button_frame.pack(fill="x", expand="yes", padx=20)

button_update = Button(button_frame, text="Update Record", command=update)
button_update.grid(row=0, column=0, padx=10, pady=10)

button_add = Button(button_frame, text="Add New Record", command=add_new)
button_add.grid(row=0, column=1, padx=10, pady=10)

button_remove = Button(button_frame, text="Remove Record", command=confirm)
button_remove.grid(row=0, column=2, padx=10, pady=10)

button_moveup = Button(button_frame, text="Move Up", command=up)
button_moveup.grid(row=0, column=3, padx=10, pady=10)

button_movedown = Button(button_frame, text="Move Down", command=down)
button_movedown.grid(row=0, column=4, padx=10, pady=10)

button_select = Button(button_frame, text="Clear Selected", command=cler_entries1)
button_select.grid(row=0, column=5, padx=10, pady=10)

button_report = Button(button_frame, text="Print Report", command=print_report)
button_report.grid(row=0, column=6, padx=10, pady=10 )

button_email = Button(button_frame, text="Email Report", command=send_message)
button_email.grid(row=0, column=7, padx=10, pady=10)

button_refresh = Button(button_frame, text= "Refresh", command=refresh)
button_refresh.grid(row=0, column=8, padx=10, pady=10)

button_quit = Button(button_frame, text="Exit", command=quit)
button_quit.grid(row=0, column=9, padx=10, pady=10)


# Bind
my_tree.bind("<ButtonRelease-1>", select_record)

# Run to pull database from start
query_db()

root.mainloop()
root.after(5000, update, query_db, add_new, confirm, remove)
