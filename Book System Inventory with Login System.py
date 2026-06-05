import tkinter as tk
import openpyxl as op
from tkinter import ttk, messagebox
import os

windows = tk.Tk()
windows.title("Book Inventory System")
windows.geometry("1400x900")
windows.config(bg="lightblue")
windows.withdraw()

# Login Conditions
def account_login():
    if not main_validation():
        return
    
    account_search = ent1.get().lower()
    password_search = ent2.get()

    wbk = op.load_workbook("Login_Data.xlsx")
    sheet = wbk.active

    global lbl2, lbl3

    for account in sheet.iter_rows(min_row=2, values_only=True):
        if account_search == str(account[1]).lower():
            if password_search == str(account[2]):
                messagebox.showinfo("Success", "Login Successfully!!!")

                lbl2 = tk.Label(main_frame1, text=f"{account[1]}", font=("Arial", 12), fg="black", bg="white")
                lbl2.place(x=980, y=25)

                lbl3 = tk.Label(main_frame1, text=f"{account[3]}", font=("Arial", 12, "bold"), fg="green", bg="white")
                lbl3.place(x=980, y=50)

                
                login_windows.destroy()
                windows.deiconify()
                return True
    
    messagebox.showerror("Error", "Invalid Credentials, Try Again!")
    return False

# Saving Information
def account_save():

    if not create_validation():
        return

    user_name = create_ent1.get()
    pass_word = create_ent2.get()
    admin_staff = adminstaff_var.get()

    wbk = op.load_workbook("Login_Data.xlsx")
    sheet = wbk.active

    user_id = f"User No. {sheet.max_row}"

    if admin_staff == "Administrator":
        position = "Administrator"
    
    elif admin_staff == "Staff":
        position = "Staff"

    sheet.append([user_id, user_name, pass_word, position])
    wbk.save("Login_Data.xlsx")
    create.destroy()
    login_windows.deiconify()

#Login Validation
def create_validation():

    username = create_ent1.get().lower()
    password = create_ent2.get()

    if not username:
        messagebox.showerror("Error", "Enter Username!!!")
        return False
    
    if not password:
        messagebox.showerror("Error", "Enter Password!!!")
        return False
    
    if len(username) <= 5:
        messagebox.showerror("Error", "Username Should Have Atleast 5 Characters")
        return False
    
    if len(password) <= 8:
        messagebox.showerror("Error", "Username Should Have Atleast 8 Characters")
        return False

    messagebox.showinfo("Success", "Account Created Successfully!!!")
    return True

def main_validation():

    username = ent1.get().lower()
    password = ent2.get()

    if not username:
        messagebox.showerror("Error", "Enter Username!!!")
        return False
    
    if not password:
        messagebox.showerror("Error", "Enter Password!!!")
        return False
    
    return True

#Check Button Events
def chkbutton2():
    status2 = create_onoff_var.get()

    if status2 == 1:
        create_ent2['show'] = ""
    
    elif status2 == 0:
        create_ent2['show'] = "*"

def chkbutton():
    status1 = onoff_var.get()

    if status1 == 1:
        ent2['show'] = ""
    
    elif status1 == 0:
        ent2['show'] = "*"

# Create and Save Account
def create_account():
    global create
    create = tk.Toplevel(login_windows)
    create.title("Book Inventory System")
    create.geometry("1400x900")
    create.config(bg="lightblue")

    create.iconphoto(False, logo)

    #For Frame
    frame1 = tk.LabelFrame(create, width=760, height=200, bg="lightblue")
    frame1.grid_propagate(False)
    frame1.place(x=300, y= 280)

    #For Label
    lbl1 = tk.Label(create, text="Create User Account", font=("Arial", 30, "bold"), fg="black", bg="lightblue")
    lbl1.place(x=470, y=180)

    lbl2 = tk.Label(create, text="Login System", font=("Arial", 15, "bold"), fg="black", bg="lightblue")
    lbl2.place(x=600, y=250)

    lbl2 = tk.Label(frame1, text="Username:", font=("Arial", 15, "bold"), fg="black", bg="lightblue")
    lbl2.grid(row=0, column=0, padx=100, pady=50)

    lbl3 = tk.Label(frame1, text="Password:", font=("Arial", 15, "bold"), fg="black", bg="lightblue")
    lbl3.grid(row=1, column=0)

    #For Entries
    global create_ent1, create_ent2
    create_ent1 = tk.Entry(frame1, font=("Arial", 15, "bold"))
    create_ent1.grid(row=0, column=1)

    create_ent2 = tk.Entry(frame1, font=("Arial", 15, "bold"), show="*")
    create_ent2.grid(row=1, column=1)

    #For CheckButtons
    global create_onoff_var, adminstaff_var
    create_onoff_var = tk.IntVar()
    create_chkvar1 = tk.Checkbutton(frame1, text="Show Password", variable=create_onoff_var, onvalue=1, offvalue=0, bg="lightblue", command=chkbutton2)
    create_chkvar1.grid(row=1, column=2, padx = 50)

    adminstaff_var = tk.StringVar(value="Staff")
    chkvar2 = tk.Checkbutton(frame1, text="Account as Admin?", variable=adminstaff_var, onvalue="Administrator", offvalue="Staff", bg="lightblue", command=chkbutton)
    chkvar2.grid(row=0, column=2, padx = 50)

    #For Buttons
    btn1 = tk.Button(create, text="Create", font=("Arial", 15, "bold"), command=account_save)
    btn1.place(x=520, y=500)

    btn2 = tk.Button(create, text="Back to Login", font=("Arial", 15, "bold"), command=lambda: (create.destroy(), login_windows.deiconify()))
    btn2.place(x=720, y=500)

# Login User Interface
def login():
    global login_windows
    login_windows = tk.Toplevel(windows)
    login_windows.title("Book Inventory System")
    login_windows.geometry("1400x900")
    login_windows.config(bg="lightblue")

    logo = tk.PhotoImage(file="download.png")
    login_windows.iconphoto(False, logo)

    if not os.path.exists("Login_Data.xlsx"):

        wbk1 = op.Workbook()
        sheet = wbk1.active

        sheet['A1'] = "User ID"
        sheet['B1'] = "Username"
        sheet['C1'] = "Password"
        sheet['D1'] = "Position"

        wbk1.save("Login_Data.xlsx")

    #For Frame
    global frame1
    frame1 = tk.LabelFrame(login_windows, width=760, height=200, bg="lightblue")
    frame1.grid_propagate(False)
    frame1.place(x=300, y= 280)

    #For Label
    global login_bl1, login_lbl2, login_lbl3
    login_lbl1 = tk.Label(login_windows, text="Book Inventory System", font=("Arial", 30, "bold"), fg="black", bg="lightblue")
    login_lbl1.place(x=450, y=180)

    login_lbl2 = tk.Label(login_windows, text="Login System", font=("Arial", 15, "bold"), fg="black", bg="lightblue")
    login_lbl2.place(x=600, y=250)

    login_lbl2 = tk.Label(frame1, text="Username:", font=("Arial", 15, "bold"), fg="black", bg="lightblue")
    login_lbl2.grid(row=0, column=0, padx=100, pady=50)

    lbl3 = tk.Label(frame1, text="Password:", font=("Arial", 15, "bold"), fg="black", bg="lightblue")
    lbl3.grid(row=1, column=0)

    #For Entries
    global ent1, ent2
    ent1 = tk.Entry(frame1, font=("Arial", 15, "bold"))
    ent1.grid(row=0, column=1)

    ent2 = tk.Entry(frame1, font=("Arial", 15, "bold"), show="*")
    ent2.grid(row=1, column=1)

    #For Buttons
    btn1 = tk.Button(login_windows, text="Login", font=("Arial", 15, "bold"), command=account_login)
    btn1.place(x=520, y=500)

    btn2 = tk.Button(login_windows, text="Create Account", font=("Arial", 15, "bold"), command=lambda: (login_windows.withdraw(), create_account()))
    btn2.place(x=720, y=500)

    #For CheckButton
    global onoff_var
    onoff_var = tk.IntVar()
    chkvar1 = tk.Checkbutton(frame1, text="Show Password", variable=onoff_var, onvalue=1, offvalue=0, bg="lightblue", command=chkbutton)
    chkvar1.grid(row=1, column=2, padx = 50)

def logout():

    logout_msg = messagebox.askyesno("Log-Out", "Do You Wish To Log-Out?")
    if not logout_msg:
        return False

    windows.withdraw()
    login()
    lbl2['text'] = " "
    lbl3['text'] = " "
    
#Backup
def export_save():

    save_confirm = messagebox.askyesno("Confirm", "Do You Want To Back It Up?")
    if not save_confirm:
        return
    
    wbk1 = op.load_workbook("Reyes_Database.xlsx")
    sheet = wbk1.active

    wbk2 = op.Workbook()
    sheet1 = wbk2.active

    for data_row in sheet.iter_rows(values_only=True):
        sheet1.append(data_row)
    
    wbk2.save("Reyes_Backup.xlsx")
    messagebox.showinfo("Success", "Data Backup Successfully")
    return

#Main Windows Dashboard
def total_stocks():
    wbk = op.load_workbook("Reyes_Database.xlsx")
    sheet = wbk.active
    
    total_stock = 0

    for book_stocks in sheet.iter_rows(min_row=2, values_only=True):
        total_stock += int(book_stocks[5])
    
    lbl11['text'] = f"{total_stock}"

def out_stocks():

    wbk = op.load_workbook("Reyes_Database.xlsx")
    sheet = wbk.active

    out_of_stocks = 0

    for book_stocks in sheet.iter_rows(min_row=2, values_only=True):
        if int(book_stocks[5]) == 0:
            out_of_stocks += 1

    lbl12['text'] = f"{out_of_stocks}"

#Clearing Entry from Searches
def view_clear():

    search_ent.delete(0, tk.END)

    display_topview()

def del_clear():

    search_ent.delete(0, tk.END)

    display_topdel()

def upt_clear():

    search_ent.delete(0, tk.END)

    display_topupdate()

#Search Functions for CRUD Top Levels except Creating Data
def view_search():

    search_data = search_ent.get().strip().lower()

    worbook = op.load_workbook("Reyes_Database.xlsx")
    shet = worbook.active

    for row_books1 in tree_view.get_children():
        tree_view.delete(row_books1)
    
    for row_books1 in shet.iter_rows(min_row=2, values_only=True):
        if search_data in str(row_books1[1]).lower() or search_data in str(row_books1[2]).lower() or search_data in str(row_books1[3]).lower() or search_data in str(row_books1[4]).lower():
            tree_view.insert("", tk.END, values=row_books1)

    if not search_data:
        messagebox.showerror("Error", "Searched Book Not Available")
        return

def del_search():

    search_data = search_ent.get().strip().lower()

    worbook = op.load_workbook("Reyes_Database.xlsx")
    shet = worbook.active

    for row_books1 in tree_del.get_children():
        tree_del.delete(row_books1)
    
    for row_books1 in shet.iter_rows(min_row=2, values_only=True):
        if search_data in str(row_books1[1]).lower() or search_data in str(row_books1[2]).lower() or search_data in str(row_books1[3]).lower() or search_data in str(row_books1[4]).lower():
            tree_del.insert("", tk.END, values=row_books1)

    if not search_data:
        messagebox.showerror("Error", "Searched Book Not Available")
        return

def update_search():

    search_data = search_ent.get().strip().lower()

    worbook = op.load_workbook("Reyes_Database.xlsx")
    shet = worbook.active

    for row_books1 in tree_update.get_children():
        tree_update.delete(row_books1)
    
    for row_books1 in shet.iter_rows(min_row=2, values_only=True):
        if search_data in str(row_books1[1]).lower() or search_data in str(row_books1[2]).lower() or search_data in str(row_books1[3]).lower() or search_data in str(row_books1[4]).lower():
            tree_update.insert("", tk.END, values=row_books1)

    if not search_data:
        messagebox.showerror("Error", "Searched Book Not Available")
        return

#Adding/Creating Data
def add_books():
    if not adding_validation():
        return
    
    book_title = add_ent1.get()
    author = add_ent2.get()
    publisher = add_ent3.get()
    category = add_cmbx1.get()
    stocks = int(add_ent4.get())

    workbook1 = op.load_workbook("Reyes_Database.xlsx")
    sheet1 = workbook1.active

    book_id = f"B{sheet1.max_row}"

    if stocks == 0:
        status = "No Stocks"
    
    elif stocks >=1 and stocks <=10:
        status = "Low Stocks"
    
    else:
        status = "Available"

    sheet1.append([book_id, book_title, author, category, publisher, stocks, status])
    count_books = sheet1.max_row - 1
    workbook1.save("Reyes_Database.xlsx")

    add_ent1.delete(0, tk.END)
    add_ent2.delete(0, tk.END)
    add_ent3.delete(0, tk.END)
    add_ent4.delete(0, tk.END)
    add_cmbx1.set("")


    display_main1()
    display_main2()
    total_stocks()
    out_stocks()
    lbl9 ['text'] = f"{count_books}"

# Deleting Data
def del_books():

    book_select = tree_del.focus()

    if not book_select:
        messagebox.showerror("Error", "Select Data First")
        return

    values = tree_del.item(book_select, "values")
    book_id = values[0]

    confirm_del = messagebox.askyesno("Confirm", "Are You Sure You Want To Delete?")
    if not confirm_del:
        return
    
    workbook = op.load_workbook("Reyes_Database.xlsx")
    sheet = workbook.active

    for books, row_books in enumerate(sheet.iter_rows(min_row=2), start=2):
        if row_books[0].value == book_id:
            sheet.delete_rows(books)
            break

    messagebox.showinfo("Sucess", "Book Data Removed")
    workbook.save("Reyes_Database.xlsx")

    del_ent1.delete(0, tk.END)
    del_ent2.delete(0, tk.END)
    del_ent3.delete(0, tk.END)
    del_ent4.delete(0, tk.END)
    del_cmbx1.set("")
    del_lbl9['text'] = ""
    del_lbl10['text'] = ""

    count_books = sheet.max_row - 1
    display_main1()
    display_main2()
    display_topdel()
    lbl9 ['text'] = f"{count_books}"
    total_stocks()
    out_stocks()

#Updating Data
def update_books():

    book_selected = tree_update.focus()

    if not book_selected:
        messagebox.showerror("Error", "Select a Book First")
        return
    
    if not updating_validation():
        return
    
    worbook = op.load_workbook("Reyes_Database.xlsx")
    sheet = worbook.active
    
    values = tree_update.item(book_selected, "values")
    book_id = values[0]

    book_title = upt_ent1.get()
    author = upt_ent2.get()
    publisher = upt_ent3.get()
    stocks = upt_ent4.get()

    count_books = sheet.max_row - 1
    stocks = int(upt_ent4.get())

    if stocks == 0:
        status = "No Stocks"
    elif stocks <= 10:
        status = "Low Stocks"
    else:
        status = "Available"
            
    for row_book in sheet.iter_rows(min_row=2):
        if row_book[0].value == book_id:
            row_book[1].value = book_title
            row_book[2].value = author
            row_book[3].value = update_cmbx1.get()
            row_book[4].value = publisher
            row_book[5].value = int(stocks)
            row_book[6].value = status
            
    update_lbl9['text'] = status
    worbook.save("Reyes_Database.xlsx")

    messagebox.showinfo("Success", "Record Updated Successfully")
    display_topupdate()
    display_main1()
    display_main2()
    total_stocks()
    lbl9 ['text'] = f"{count_books}"
    out_stocks()

# Auto Population in Top Levels of CRUD (Except Creating Data and Reading Data)
def selected_update(event):
    select_book = tree_update.focus()
    value_book = tree_update.item(select_book, "values")

    if value_book:
        upt_ent1.delete(0, tk.END)
        upt_ent2.delete(0, tk.END)
        upt_ent3.delete(0, tk.END)
        upt_ent4.delete(0, tk.END)
        update_cmbx1.set("")
        update_lbl8['text'] = ""
        update_lbl9['text'] = ""

        update_lbl8['text'] = value_book[0]
        upt_ent1.insert(0, value_book[1])
        upt_ent2.insert(0, value_book[2])
        update_cmbx1.set(value_book[3])
        upt_ent3.insert(0, value_book[4])
        upt_ent4.insert(0, value_book[5])
        update_lbl9['text'] = value_book[6]

def selected_delete(event):
    select_book = tree_del.focus()
    value_book = tree_del.item(select_book, "values")

    if value_book:
        del_ent1.delete(0, tk.END)
        del_ent2.delete(0, tk.END)
        del_ent3.delete(0, tk.END)
        del_ent4.delete(0, tk.END)
        del_cmbx1.set("")
        del_lbl9['text'] = ""
        del_lbl10['text'] = ""

        del_lbl9['text'] = value_book[0]
        del_ent1.insert(0, value_book[1])
        del_ent2.insert(0, value_book[2])
        del_cmbx1.set(value_book[3])
        del_ent3.insert(0, value_book[4])
        del_ent4.insert(0, value_book[5])
        del_lbl10['text'] = value_book[6]

#Displaying Datas in Top Levels of CRUD (except creating data)
def display_topdel():

    workbook1 = op.load_workbook("Reyes_Database.xlsx")
    sheet = workbook1.active

    for row_books in tree_del.get_children():
        tree_del.delete(row_books)

    for row_books in sheet.iter_rows(min_row=2, values_only=True):
        tree_del.insert("", tk.END, values=row_books)

def display_topview():

    workbook1 = op.load_workbook("Reyes_Database.xlsx")
    sheet = workbook1.active
    
    for row_books1 in tree_view.get_children():
        tree_view.delete(row_books1)

    for row_books1 in sheet.iter_rows(min_row=2, values_only=True):
        tree_view.insert("", tk.END, values=row_books1)

def display_topupdate():

    workbook1 = op.load_workbook("Reyes_Database.xlsx")
    sheet = workbook1.active
    
    for row_books1 in tree_update.get_children():
        tree_update.delete(row_books1)

    for row_books1 in sheet.iter_rows(min_row=2, values_only=True):
        tree_update.insert("", tk.END, values=row_books1)

#Displaying Datas in Treeview of Main Windows
def display_main1():

    workbook1 = op.load_workbook("Reyes_Database.xlsx")
    sheet = workbook1.active
    
    for row_books in tree2.get_children():
        tree2.delete(row_books)

    for row_books in sheet.iter_rows(min_row=2, values_only=True):
        tree2.insert("", tk.END, values=row_books)

def display_main2():

    workbook1 = op.load_workbook("Reyes_Database.xlsx")
    sheet = workbook1.active
    
    for row_books in tree1.get_children():
        tree1.delete(row_books)

    bookdata = list(sheet.iter_rows(min_row=2, values_only=True))

    for row_books in reversed(bookdata):
        tree1.insert("", tk.END, values=row_books)

#Input Validation for the entries
def current_top():
    messagebox.showerror("Error", "You are currently viewing this section")

def updating_validation():
    book_title = upt_ent1.get()
    author = upt_ent2.get()
    publisher = upt_ent3.get()
    stocks = upt_ent4.get()

    if not book_title or not author or not publisher or not stocks:
        messagebox.showerror("Error", "Fill up all blanks")
        return False
    
    if update_cmbx1.get() == "":
        messagebox.showerror("Error", "Fill up the category")
        return False
    
    if not stocks.isdigit():
        messagebox.showerror("Error", "Stocks should be digit")
        return False

    return True

def adding_validation():
    book_title = add_ent1.get()
    author = add_ent2.get()
    publisher = add_ent3.get()
    stocks = add_ent4.get()

    if not book_title or not author or not publisher or not stocks:
        messagebox.showerror("Error", "Fill up all blanks")
        return False
    
    if add_cmbx1.get() == "":
        messagebox.showerror("Error", "Fill up the category")
        return False
    
    if not stocks.isdigit():
        messagebox.showerror("Error", "Stocks should be digit")
        return False
    
    messagebox.showinfo("Success", "Data Added Successfully")
    return True

#Top Level of the CRUD
def top_add():

    windows.withdraw()

    add_inven = tk.Toplevel(windows)
    add_inven.title("Book Inventory System")
    add_inven.geometry("1400x900")
    add_inven.config(bg="white")

    add_inven.iconphoto(False, logo)

    #For Frames (Top Level)
    top_frame2 = tk.LabelFrame(add_inven, bg="#d8c3a5", relief="sunken", width=1250, height=850)
    top_frame2.grid_propagate(False)
    top_frame2.place(x=200, y=0)

    top_frame1 = tk.LabelFrame(top_frame2, height=500, width=500, relief="sunken")
    top_frame1.grid_propagate(False)
    top_frame1.place(x=550, y=150)

    #For Photos (Top Level)
    phtlbl1 = tk.Label(top_frame2, image=toplvl_img, bg="#d8c3a5")
    phtlbl1.place(x=100, y = 30)

    #For Labels (Top Level)
    lbl1 = tk.Label(top_frame2, text="Add Library:", font=("Arial", 25, "bold"), fg="#1a1a1a", bg="#d8c3a5")
    lbl1.place(x=600, y=100)

    lbl2 = tk.Label(top_frame1, text="Book Title:", font=("Arial", 15, "bold"), fg="#04cf0e")
    lbl2.grid(column=0, row=0, padx=30, pady=30)

    lbl3 = tk.Label(top_frame1, text="Author:", font=("Arial", 15, "bold"), fg="#04cf0e")
    lbl3.grid(column=0, row=1)

    lbl4 = tk.Label(top_frame1, text="Category:", font=("Arial", 15, "bold"), fg="#04cf0e")
    lbl4.grid(column=0, row=2, pady=30)

    lbl5 = tk.Label(top_frame1, text="Publisher:", font=("Arial", 15, "bold"), fg="#04cf0e")
    lbl5.grid(column=0, row=3)

    lbl6 = tk.Label(top_frame1, text="Stocks:", font=("Arial", 15, "bold"), fg="#04cf0e")
    lbl6.grid(column=0, row=4, pady=30)

    lbl7 = tk.Label(add_inven, text="MENU:", font=("Arial", 20, "bold"), fg="green", bg="white")
    lbl7.grid(column=0, row=0, padx=30, pady=20)

    #For Buttons (Top Level)

    btn1 = tk.Button(add_inven, text="Add Books in Inventory", font=("Arial", 10, "bold"), bg="green", command=current_top)
    btn1.grid(column=0, row=1)

    btn2 = tk.Button(add_inven, text="View Inventory", font=("Arial", 10, "bold"), command=lambda:( add_inven.destroy(),top_view() ))
    btn2.grid(column=0, row=2, pady=20)

    btn3 = tk.Button(add_inven, text="Delete Inventory", font=("Arial", 10, "bold"), command=lambda:(add_inven.destroy(), top_del() ))
    btn3.grid(column=0, row=3)

    btn4 = tk.Button(add_inven, text="Update Book Information", font=("Arial", 10, "bold"), command=lambda:(add_inven.destroy(), top_update() ))
    btn4.grid(column=0, row=4, padx=20, pady=20)

    btn5 = tk.Button(add_inven, text="Backup / Export", font=("Arial", 10, "bold"), command=lambda:(add_inven.destroy(), windows.deiconify(), export_save()))
    btn5.grid(column=0, row=5)

    btn7 = tk.Button(add_inven, text="Main Menu", font=("Arial", 10, "bold"), command=lambda:(add_inven.destroy(), windows.deiconify()))
    btn7.grid(column=0, row=6, pady=20)

    btn8 = tk.Button(top_frame1, text="Submit", font=("Arial", 10, "bold"), command=add_books)
    btn8.grid(column=0, row=5, columnspan=2, pady=20)

    #For Entry (Top Level)
    global add_ent1, add_ent2, add_ent3, add_ent4
    add_ent1 = tk.Entry(top_frame1, font=("Arial", 12), relief="sunken")
    add_ent1.grid(column=1, row=0, padx=30, pady=30)

    add_ent2 = tk.Entry(top_frame1, font=("Arial", 12), relief="sunken")
    add_ent2.grid(column=1, row=1)

    add_ent3 = tk.Entry(top_frame1, font=("Arial", 12), relief="sunken")
    add_ent3.grid(column=1, row=3)

    add_ent4 = tk.Entry(top_frame1, font=("Arial", 12), relief="sunken")
    add_ent4.grid(column=1, row=4, pady=30)

    #For Combobox
    global add_cmbx1
    top_cat1=["Academic","Fiction","Non-Fiction","Science","Technology","History","Business","Self Help","Reference","Others", ""]
    add_cmbx1 = ttk.Combobox(top_frame1, values=top_cat1)
    add_cmbx1.grid(column=1, row=2, pady=30)

def top_view():
    windows.withdraw()

    view_inven = tk.Toplevel(windows)
    view_inven.title("Book Inventory System")
    view_inven.geometry("1400x900")
    view_inven.config(bg="white")
    view_inven.iconphoto(False, logo)

    #For Frame (Top Level)
    top_frame1 = tk.LabelFrame(view_inven, bg="#ffcb44", relief="sunken", width=1250, height=850)
    top_frame1.grid_propagate(False)
    top_frame1.place(x=200, y=0)

    #For Label (Top Label)
    lbl1 = tk.Label(view_inven, text="MENU:", font=("Arial", 20, "bold"), fg="green", bg="white")
    lbl1.grid(column=0, row=0, padx=30, pady=20)

    lbl2 = tk.Label(top_frame1, text="Current Inventory", font=("Arial", 20, "bold"), fg="black", bg="#ffcb44")
    lbl2.place(x=460, y=50)

    #For Buttons (Top Level)
    btn1 = tk.Button(view_inven, text="Add Books in Inventor", font=("Arial", 10, "bold"), command=lambda:(view_inven.destroy(), top_add(), ))
    btn1.grid(column=0, row=1)

    btn2 = tk.Button(view_inven, text="View Inventory", font=("Arial", 10, "bold"), bg="green", command=current_top)
    btn2.grid(column=0, row=2, pady=20)

    btn3 = tk.Button(view_inven, text="Delete Inventory", font=("Arial", 10, "bold"), command=lambda:(view_inven.destroy(), top_del() ))
    btn3.grid(column=0, row=3)

    btn4 = tk.Button(view_inven, text="Update Book Information", font=("Arial", 10, "bold"), command=lambda:(view_inven.destroy(), top_update() ))
    btn4.grid(column=0, row=4, padx=20, pady=20)

    btn5 = tk.Button(view_inven, text="Backup / Export", font=("Arial", 10, "bold"), command=lambda:(view_inven.destroy(), windows.deiconify(), export_save()))
    btn5.grid(column=0, row=5)

    btn7 = tk.Button(view_inven, text="Main Menu", font=("Arial", 10, "bold"), command=lambda:(view_inven.destroy(), windows.deiconify()))
    btn7.grid(column=0, row=6, pady=20)

    btn8 = tk.Button(top_frame1, text="Search", font=("Arial", 15, "bold"), command=view_search)
    btn8.place(x=485, y=680)

    btn9 = tk.Button(top_frame1, text="Clear", font=("Arial", 15, "bold"), command=view_clear)
    btn9.place(x=575, y=680)

    global search_ent
    #Search Entry
    search_ent = tk.Entry(top_frame1, relief="sunken", font=("Arial", 15))
    search_ent.place(x=460, y=640)

    #Photo Images
    phtlbl1 = tk.Label(top_frame1, image=toplvl_img1, bg="#ffcb44")
    phtlbl1.grid(column=0, row=0)

    phtlbl2 = tk.Label(top_frame1, image=toplvl_img2, bg="#ffcb44")
    phtlbl2.place(x=940, y=550)

    global tree_view
    #Treeview in Top Level
    tree_view = ttk.Treeview(top_frame1, columns=("Book ID", "Title", "Author", "Category", "Publisher", "Stock", "Status"), show="headings", height=25)
    for item1 in ("Book ID", "Title", "Author", "Category", "Publisher", "Stock", "Status"):
        tree_view.heading(item1, text=item1)
    tree_view.column("Book ID", width=53)
    tree_view.column("Title", width=150)
    tree_view.column("Author", width=118)
    tree_view.column("Category", width=68)
    tree_view.column("Publisher", width=108)
    tree_view.column("Stock", width=142)
    tree_view.column("Status", width=100)
    tree_view.place(x=220, y=100)

    display_topview()

def top_del():
    windows.withdraw()

    del_inven = tk.Toplevel(windows)
    del_inven.title("Book Inventory System")
    del_inven.geometry("1400x900")
    del_inven.config(bg="white")
    del_inven.iconphoto(False, logo)

    #Frame
    top_frame1 = tk.LabelFrame(del_inven, bg="#ec4c4c", relief="sunken", width=1250, height=850)
    top_frame1.grid_propagate(False)
    top_frame1.place(x=200, y=0)

    top_frame2 = tk.LabelFrame(top_frame1, bg="white", relief="sunken", width=300, height=350)
    top_frame2.grid_propagate(False)
    top_frame2.place(x=840, y=80)

    #Label
    del_lbl1 = tk.Label(del_inven, text="MENU:", font=("Arial", 20, "bold"), fg="green", bg="white")
    del_lbl1.grid(column=0, row=0, padx=30, pady=20)

    del_lbl2 = tk.Label(top_frame2, text="Book ID:", font=("Arial", 15, "bold"), fg="black", bg="white")
    del_lbl2.grid(column=0, row=0, padx=30, pady=10)

    del_lbl3 = tk.Label(top_frame2, text="Book Title:", font=("Arial", 15, "bold"), fg="black", bg="white")
    del_lbl3.grid(column=0, row=1)

    del_lbl4 = tk.Label(top_frame2, text="Author:", font=("Arial", 15, "bold"), fg="black", bg="white")
    del_lbl4.grid(column=0, row=2, pady=10)

    del_lbl5 = tk.Label(top_frame2, text="Category:", font=("Arial", 15, "bold"), fg="black", bg="white")
    del_lbl5.grid(column=0, row=3)

    del_lbl6 = tk.Label(top_frame2, text="Publisher:", font=("Arial", 15, "bold"), fg="black", bg="white")
    del_lbl6.grid(column=0, row=4, pady=10)

    del_lbl7 = tk.Label(top_frame2, text="Stocks:", font=("Arial", 15, "bold"), fg="black", bg="white")
    del_lbl7.grid(column=0, row=5)

    del_lbl8 = tk.Label(top_frame2, text="Status:", font=("Arial", 15, "bold"), fg="black", bg="white")
    del_lbl8.grid(column=0, row=6, pady=10)

    global del_lbl9, del_lbl10
    del_lbl9 = tk.Label(top_frame2, text="__", font=("Arial", 15, "bold"), fg="black", bg="white")
    del_lbl9.grid(column=1, row=0, pady=10)

    del_lbl10= tk.Label(top_frame2, text="__", font=("Arial", 15, "bold"), fg="black", bg="white")
    del_lbl10.grid(column=1, row=6, pady=10)

    #Buttons
    btn1 = tk.Button(del_inven, text="Add Books in Inventor", font=("Arial", 10, "bold"), command=lambda:(del_inven.destroy(), top_add(), ))
    btn1.grid(column=0, row=2)

    btn2 = tk.Button(del_inven, text="View Inventory", font=("Arial", 10, "bold"), command=lambda:(del_inven.destroy(), top_view() ))
    btn2.grid(column=0, row=3, pady=20)

    btn3 = tk.Button(del_inven, text="Delete Inventory", font=("Arial", 10, "bold"), bg="green", command=current_top)
    btn3.grid(column=0, row=4)

    btn4 = tk.Button(del_inven, text="Update Book Information", font=("Arial", 10, "bold"), command=lambda:(del_inven.destroy(), top_update() ))
    btn4.grid(column=0, row=5, padx=20, pady=20)

    btn5 = tk.Button(del_inven, text="Backup / Export", font=("Arial", 10, "bold"), command=lambda:(del_inven.destroy(), windows.deiconify(), export_save()))
    btn5.grid(column=0, row=6)

    btn7 = tk.Button(del_inven, text="Main Menu", font=("Arial", 10, "bold"), command=lambda:(del_inven.destroy(), windows.deiconify(), export_save()))
    btn7.grid(column=0, row=7, pady=20)

    global btn8
    btn8 = tk.Button(top_frame2, text="Delete", font=("Arial", 15, "bold"), command=del_books)
    btn8.grid(column=0, row=7, columnspan=2)

    btn9 = tk.Button(top_frame1, text="Clear", font=("Arial", 15, "bold"), command=del_clear)
    btn9.place(x=445, y=650)

    btn10 = tk.Button(top_frame1, text="Search", font=("Arial", 15, "bold"), command=del_search)
    btn10.place(x=355, y=650)

    global search_ent
    #Search Entry
    search_ent = tk.Entry(top_frame1, relief="sunken", font=("Arial", 15))
    search_ent.place(x=330, y=620)

    #For Photo Image (Top Level)
    phtlbl1 = tk.Label(top_frame1, image=toplvl_img3, bg="#ec4c4c")
    phtlbl1.place(x=840, y=450)

    global tree_del
    #Treeview in Top Level
    tree_del = ttk.Treeview(top_frame1, columns=("Book ID", "Title", "Author", "Category", "Publisher", "Stock", "Status"), show="headings", height=20)
    for item1 in ("Book ID", "Title", "Author", "Category", "Publisher", "Stock", "Status"):
        tree_del.heading(item1, text=item1)
    tree_del.column("Book ID", width=53)
    tree_del.column("Title", width=150)
    tree_del.column("Author", width=118)
    tree_del.column("Category", width=68)
    tree_del.column("Publisher", width=108)
    tree_del.column("Stock", width=142)
    tree_del.column("Status", width=100)
    tree_del.place(x=50, y=180)
    tree_del.bind("<<TreeviewSelect>>", selected_delete)

    #Label (Top Level)
    lbl_top1 = tk.Label(top_frame1, text="Select Book To Delete:", font=("Arial", 20, "bold"), fg="black", bg="#ec4c4c")
    lbl_top1.place(x=80, y=130)

    #Entry
    global del_ent1, del_ent2, del_ent3, del_ent4
    del_ent1 = tk.Entry(top_frame2, relief="sunken")
    del_ent1.grid(column=1, row=1)

    del_ent2 = tk.Entry(top_frame2, relief="sunken")
    del_ent2.grid(column=1, row=2)

    del_ent3 = tk.Entry(top_frame2, relief="sunken")
    del_ent3.grid(column=1, row=4)

    del_ent4 = tk.Entry(top_frame2, relief="sunken")
    del_ent4.grid(column=1, row=5)

    #For combobox
    global del_cmbx1
    top_cat1=["Academic","Fiction","Non-Fiction","Science","Technology","History","Business","Self Help","Reference","Others", ""]
    del_cmbx1 = ttk.Combobox(top_frame2, values=top_cat1)
    del_cmbx1.grid(column=1, row=3)

    display_topdel()

def top_update():
    windows.withdraw()

    update_inven = tk.Toplevel(windows)
    update_inven.title("Book Inventory System")
    update_inven.geometry("1400x900")
    update_inven.config(bg="white")
    update_inven.iconphoto(False, logo)

    #Frame
    top_frame1 = tk.LabelFrame(update_inven, bg="#90EE90", relief="sunken", width=1250, height=850)
    top_frame1.grid_propagate(False)
    top_frame1.place(x=200, y=0)

    top_frame2 = tk.LabelFrame(top_frame1, bg="white", relief="sunken", width=300, height=350)
    top_frame2.grid_propagate(False)
    top_frame2.place(x=840, y=80)

    #Label
    lbl1 = tk.Label(update_inven, text="MENU:", font=("Arial", 20, "bold"), fg="green", bg="white")
    lbl1.grid(column=0, row=0, padx=30, pady=20)

    lbl2 = tk.Label(top_frame1, text="Select Book To Update Information:", font=("Arial", 20, "bold"), fg="black", bg="#90EE90")
    lbl2.place(x=80, y=110)

    lbl3 = tk.Label(top_frame2, text="Book ID:", font=("Arial", 15, "bold"), fg="black", bg="white")
    lbl3.grid(column=0, row=0, padx=30, pady=10)

    lbl4 = tk.Label(top_frame2, text="Book Title:", font=("Arial", 15, "bold"), fg="black", bg="white")
    lbl4.grid(column=0, row=1)

    lbl5 = tk.Label(top_frame2, text="Author:", font=("Arial", 15, "bold"), fg="black", bg="white")
    lbl5.grid(column=0, row=2, pady=10)

    lbl6 = tk.Label(top_frame2, text="Category:", font=("Arial", 15, "bold"), fg="black", bg="white")
    lbl6.grid(column=0, row=3)

    lbl7 = tk.Label(top_frame2, text="Publisher:", font=("Arial", 15, "bold"), fg="black", bg="white")
    lbl7.grid(column=0, row=4, pady=10)

    lbl6 = tk.Label(top_frame2, text="Stocks:", font=("Arial", 15, "bold"), fg="black", bg="white")
    lbl6.grid(column=0, row=5)

    lbl7 = tk.Label(top_frame2, text="Status:", font=("Arial", 15, "bold"), fg="black", bg="white")
    lbl7.grid(column=0, row=6, pady=10)

    global update_lbl8, update_lbl9
    update_lbl8 = tk.Label(top_frame2, text="__", font=("Arial", 15, "bold"), fg="black", bg="white")
    update_lbl8.grid(column=1, row=0, pady=10)

    update_lbl9= tk.Label(top_frame2, text="__", font=("Arial", 15, "bold"), fg="black", bg="white")
    update_lbl9.grid(column=1, row=6, pady=10)

    #Buttons
    btn1 = tk.Button(update_inven, text="Add Books in Inventor", font=("Arial", 10, "bold"), command=lambda:(update_inven.destroy(), top_add(), ))
    btn1.grid(column=0, row=2)

    btn2 = tk.Button(update_inven, text="View Inventory", font=("Arial", 10, "bold"), command=lambda:(update_inven.destroy(), top_view() ))
    btn2.grid(column=0, row=3, pady=20)

    btn3 = tk.Button(update_inven, text="Delete Inventory", font=("Arial", 10, "bold"), command=lambda:(update_inven.destroy(), top_del() ))
    btn3.grid(column=0, row=4)

    btn4 = tk.Button(update_inven, text="Update Book Information", font=("Arial", 10, "bold"), bg="green", command=current_top)
    btn4.grid(column=0, row=5, padx=20, pady=20)

    btn5 = tk.Button(update_inven, text="Backup / Export", font=("Arial", 10, "bold"), command=lambda:(update_inven.destroy(), windows.deiconify(), export_save()))
    btn5.grid(column=0, row=6)

    btn7 = tk.Button(update_inven, text="Main Menu", font=("Arial", 10, "bold"), command=lambda:(update_inven.destroy(), windows.deiconify()))
    btn7.grid(column=0, row=7, pady=20)

    btn8 = tk.Button(top_frame2, text="Update", font=("Arial", 15, "bold"), command=update_books)
    btn8.grid(column=0, row=7, columnspan=2)

    global update_btn9, update_btn10
    update_btn9 = tk.Button(top_frame1, text="Clear", font=("Arial", 15, "bold"), command=upt_clear)
    update_btn9.place(x=445, y=670)

    update_btn10 = tk.Button(top_frame1, text="Search", font=("Arial", 15, "bold"), command=update_search)
    update_btn10.place(x=355, y=670)

    global search_ent
    #Search Entry
    search_ent = tk.Entry(top_frame1, relief="sunken", font=("Arial", 15))
    search_ent.place(x=330, y=620)

    global tree_update
    #For TreeViews
    tree_update = ttk.Treeview(top_frame1, columns=("Book ID", "Title", "Author", "Category", "Publisher", "Stock", "Status"), show="headings", height=20)
    for item1 in ("Book ID", "Title", "Author", "Category", "Publisher", "Stock", "Status"):
        tree_update.heading(item1, text=item1)
    tree_update.column("Book ID", width=53)
    tree_update.column("Title", width=150)
    tree_update.column("Author", width=118)
    tree_update.column("Category", width=68)
    tree_update.column("Publisher", width=108)
    tree_update.column("Stock", width=142)
    tree_update.column("Status", width=100)
    tree_update.place(x=50, y=160)
    tree_update.bind("<<TreeviewSelect>>", selected_update)

    #For Images
    phtlbl1 = tk.Label(top_frame1, image=toplvl_img4, bg="#90EE90")
    phtlbl1.place(x=840, y=450)

    #for Entry (Top View)
    global upt_ent1, upt_ent2, upt_ent3, upt_ent4
    upt_ent1 = tk.Entry(top_frame2, relief="sunken")
    upt_ent1.grid(column=1, row=1)

    upt_ent2 = tk.Entry(top_frame2, relief="sunken")
    upt_ent2.grid(column=1, row=2)

    upt_ent3 = tk.Entry(top_frame2, relief="sunken")
    upt_ent3.grid(column=1, row=4)

    upt_ent4 = tk.Entry(top_frame2, relief="sunken")
    upt_ent4.grid(column=1, row=5)

    #For Combobox
    global update_cmbx1
    top_cat1=["Academic","Fiction","Non-Fiction","Science","Technology","History","Business","Self Help","Reference","Others", ""]
    update_cmbx1 = ttk.Combobox(top_frame2, values=top_cat1)
    update_cmbx1.grid(column=1, row=3)

    display_topupdate()

# Condition if the file exist
if not os.path.exists("Reyes_Database.xlsx"):

    wbk1 = op.Workbook()
    sheet = wbk1.active

    sheet['A1'] = "Book ID"
    sheet['B1'] = "Title"
    sheet['C1'] = "Author"
    sheet['D1'] = "Category"
    sheet['E1'] = "Publisher"
    sheet['F1'] = "Stocks"
    sheet['G1'] = "Status"

    wbk1.save("Reyes_Database.xlsx")

toplvl_img4 = tk.PhotoImage(file="bookupdate.png")
toplvl_img3 = tk.PhotoImage(file="removebook.png")
toplvl_img2 = tk.PhotoImage(file="viewbook.png")
toplvl_img1 = tk.PhotoImage(file="viewbooks.png")
toplvl_img = tk.PhotoImage(file="books.png")
global logo
logo = tk.PhotoImage(file="download.png")
windows.iconphoto(False, logo)
icon1 = tk.PhotoImage(file="bookicon.png")

wbk = op.load_workbook("Reyes_Database.xlsx")
sheet = wbk.active

total_stock = 0
out_of_stocks = 0

for book_stocks in sheet.iter_rows(min_row=2, values_only=True):
    total_stock += int(book_stocks[5])

#For Frames
main_frame1 = tk.LabelFrame(windows, bg="white", relief="sunken", width=1180, height=750)
main_frame1.grid_propagate(False)
main_frame1.place(x=200, y=0)

frame2 = tk.LabelFrame(main_frame1, bg="lightblue", width=220, height=120)
frame2.grid_propagate(False)
frame2.place(x=850, y=200)

frame3 = tk.LabelFrame(main_frame1, bg="lightgreen", width=220, height=120)
frame3.grid_propagate(False)
frame3.place(x=850, y=350)

frame4 = tk.LabelFrame(main_frame1, bg="#FFD580", width=220, height=120)
frame4.grid_propagate(False)
frame4.place(x=850, y=500)

#For Images
img1 = tk.Label(main_frame1, image=logo)
img1.place(x=50, y=20)

#For Labels
lbl1 = tk.Label(main_frame1, text="Book Inventory System", font=("Arial", 30, "bold"), fg="blue", bg="white")
lbl1.place(x=250, y=50)

lbl4 = tk.Label(main_frame1, text="By Ryan Angelo C. Reyes (BSIT-2C)", font=("Arial", 12), fg="green", bg="white")
lbl4.place(x=320, y=100)

lbl5 = tk.Label(main_frame1, text="Newly Book Arrivals:", font=("Arial", 12, "bold"), fg="green", bg="white")
lbl5.place(x=75, y=220)

lbl6 = tk.Label(main_frame1, text="Available Books:", font=("Arial", 12, "bold"), fg="green", bg="white")
lbl6.place(x=75, y=420)

lbl7 = tk.Label(frame2, text="Total Books:", font=("Arial", 13, "bold"), fg="black", bg="lightblue")
lbl7.place(x=5, y=15)

lbl8 = tk.Label(windows, text="MENU:", font=("Arial", 20, "bold"), fg="green", bg="lightblue")
lbl8.grid(column=0, row=0, padx=30, pady=20)

lbl9 = tk.Label(frame2, text=f"{sheet.max_row - 1}", font=("Arial", 25, "bold"), fg="black", bg="lightblue")
lbl9.place(x=100, y=50)

lbl10 = tk.Label(frame3, text="Books in Stocks:", font=("Arial", 13, "bold"), fg="black", bg="lightgreen")
lbl10.place(x=5, y=15)

lbl11 = tk.Label(frame3, text=f"{total_stock}", font=("Arial", 25, "bold"), fg="black", bg="lightgreen")
lbl11.place(x=100, y=50)

lbl12 = tk.Label(frame4, text=f"{out_of_stocks}", font=("Arial", 25, "bold"), fg="black", bg="#FFD580")
lbl12.place(x=100, y=35)

lbl13 = tk.Label(frame4, text="Books in Out of Stocks:", font=("Arial", 13, "bold"), fg="black", bg="#FFD580")
lbl13.place(x=5, y=15)

#For TreeView
tree1 = ttk.Treeview(main_frame1, columns=("Book ID", "Title", "Author", "Category", "Publisher", "Stock"), show="headings", height=5)
for item in ("Book ID", "Title", "Author", "Category", "Publisher", "Stock"):
    tree1.heading(item, text=item)
tree1.column("Book ID", width=63)
tree1.column("Title", width=160)
tree1.column("Author", width=120)
tree1.column("Category", width=70)
tree1.column("Publisher", width=110)
tree1.column("Stock", width=160)
tree1.place(x=50, y=250)

tree2 = ttk.Treeview(main_frame1, columns=("Book ID", "Title", "Author", "Category", "Publisher", "Stock", "Status"), show="headings", height=8)
for item1 in ("Book ID", "Title", "Author", "Category", "Publisher", "Stock", "Status"):
    tree2.heading(item1, text=item1)
tree2.column("Book ID", width=53)
tree2.column("Title", width=150)
tree2.column("Author", width=118)
tree2.column("Category", width=68)
tree2.column("Publisher", width=108)
tree2.column("Stock", width=142)
tree2.column("Status", width=100)
tree2.place(x=50, y=450)

#For Buttons
btn1 = tk.Button(main_frame1, text="View All", font=("Arial", 10, "bold"), command=top_view)
btn1.place(x=600, y=420)

btn2 = tk.Button(windows, text="Add Books in Inventory", font=("Arial", 10, "bold"), command=top_add)
btn2.grid(column=0, row=1)

btn3 = tk.Button(windows, text="View Inventory", font=("Arial", 10, "bold"), command=top_view)
btn3.grid(column=0, row=2, pady=20)

btn4 = tk.Button(windows, text="Delete Inventory", font=("Arial", 10, "bold"), command=top_del)
btn4.grid(column=0, row=3)

btn5 = tk.Button(windows, text="Update Book Information", font=("Arial", 10, "bold"), command=top_update)
btn5.grid(column=0, row=4, padx=20, pady=20)

btn7 = tk.Button(windows, text="Backup / Export", font=("Arial", 10, "bold"), command= export_save)
btn7.grid(column=0, row=5)

btn8 = tk.Button(windows, text="Log Out", font=("Arial", 10, "bold"), command=logout)
btn8.grid(column=0, row=6, pady=20)

display_main1()
display_main2()
out_stocks()
login()

windows.mainloop()