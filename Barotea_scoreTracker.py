import re
from tkinter import messagebox
from tkinter import *
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ================= MAIN WINDOW SCORE TRACKER ================= #
window = Tk()
window.title("Score Tracker")
window.geometry("500x500")

# ================= FUNCTIONS ================= #
# Saving The Info To Excel
def save_info_to_excel():
    if not user_data_debugger():
        return
    
    name = username_entry.get().strip().lower()
    score = int(score_entry.get())

    try:
        wb = load_workbook("student_scores.xlsx")
        if "Userdata" in wb.sheetnames:
            ws = wb["Userdata"]
        else:
            ws = wb.create_sheet("Userdata")
    
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.title = "Userdata"
        ws.append(["Name", "Score"])
    
    ws.append([name, score])
    wb.save("student_scores.xlsx")

    formula_excel()
    format_excel()
    show_data()

    messagebox.showinfo(title="Success", message="Data Saved Successfully.")

    username_entry.delete(0, END)
    score_entry.delete(0, END)

    username_entry.insert(0, "Name")
    score_entry.insert(0, "Score")

# Update Function
def update_info_to_excel():
    if not user_data_debugger():
        return
    
    name = username_entry.get().strip().lower()
    score = int(score_entry.get())

    try:
        wb = load_workbook("student_scores.xlsx")
        ws = wb["Userdata"]

        for row in ws.iter_rows(min_row=2):
            cell_name = row[0].value
            if cell_name and cell_name.strip().lower() == name:
                row[1].value = score
                wb.save("student_scores.xlsx")
                formula_excel()
                show_data()
                messagebox.showinfo(title="Success", message="Score Updated Successfully.")
                return True

        messagebox.showerror(title="Error", message="User not found.")
        return False

    except Exception as e:
        print(f"Error updating score: {e}")
        return False



# Formula Fixer Function
def formula_excel():
    if not user_data_debugger():
        return
    
    wb = load_workbook("student_scores.xlsx")
    ws = wb["Userdata"]

    rows = ws.max_row

    formula_cell = f"B{rows + 2}"
    ws[formula_cell] = f"=AVERAGE(B2:B{rows})"

    ws[f"A{rows + 2}"] = "Average"

    wb.save("student_scores.xlsx")
    print("awefwf")

# Format Fixer Function
def format_excel():
    wb = load_workbook("student_scores.xlsx")
    ws = wb["Userdata"]

    # Bold Headers
    for cells in ws[1]:
        cells.font = Font(bold=True)

    # Auto Column Width
    for cols in ws.columns:
        max_length = max(len(str(cells.value)) for cell in cols)
        col_letter = get_column_letter(cols[0].column)
        ws.column_dimensions[col_letter].width = max_length + 4
        
    wb.save("student_scores.xlsx")

# New Window Data Shower Function
def show_data():
    wb = load_workbook("student_scores.xlsx")
    ws = wb["Userdata"]

    data_window = Toplevel(window)
    data_window.geometry("200x200")
    data_window.title("Stored User Data")

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        for j, value in enumerate(row):
            label = Label(data_window, text=value, width=1)
            label.grid(row=i, column=j)

# ================= DEBUGGERS ================= #
# Userdata Debugger
def user_data_debugger():
    if not validate_name_str():
        return
    elif not validate_score_int():
        return

    name = username_entry.get().strip().lower()
    score = score_entry.get()

    if (name.strip().lower() and score and name.strip().lower() != "Name" and score != "Score"):
        print("\n\nUser Info Debugger:\n\n"
              "Name: " + name + "\n"
              "Score: " + score + "\n\n")
        return True
    
    else:
        messagebox.showerror(title="Error", message="Please enter your name and score.")
        return False

# Name Debugger
def validate_name_str():
    name = username_entry.get()
    if re.fullmatch(r"[A-Za-z\s]+", name):
        return True
    else:
        messagebox.showerror("Invalid Input", "Please enter a valid name (letters only).")
        return False
    
# Score Debugger
def validate_score_int():
    try:
        value = int(score_entry.get())
        if value >= 0 and value <= 100:
            return True
        
        else:
            messagebox.showerror(title="Error", message="Please enter a score between 0 and 100.")
            return False
        
    except ValueError:
        messagebox.showerror(title="Error", message="Please enter a valid score.")
        return False

# ================= ENTRY EVENTS ================= #
# Username Entry Events
def username_on_click(event):
    if username_entry.get() == "Name":
        username_entry.delete(0, END)
def username_on_leave(event):
    if username_entry.get() == "":
        username_entry.insert(0, "Name")

# Score Entry Events
def score_on_click(event):
    if score_entry.get() == "Score":
        score_entry.delete(0, END)
def score_on_leave(event):
    if score_entry.get() == "":
        score_entry.insert(0, "Score")

# User Handle Enter Event
def user_handle_enter(event):
    if not save_info_to_excel():
        return

# ================= MAIN FRAMES ================= #
main_frame = Frame(window)
main_frame.pack()

# ================= USER INFORMATION FRAME ================= #
stundet_info_frame = LabelFrame(main_frame, text="User Information", font= ("times new roman", 20, "bold"))
stundet_info_frame.grid(row=0, column=0, sticky= NSEW, padx=10, pady=10)


# ================= WIDGETS ================= #
username_label = Label(stundet_info_frame, text="Student Name", font=("Arial", 12, "bold"))
username_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")

username_entry = Entry(stundet_info_frame, font=("Arial", 12))
username_entry.grid(row=1, column=0, sticky="nwe", padx=10, pady=5)
username_entry.insert(0, "Name")
username_entry.bind("<FocusIn>", username_on_click)
username_entry.bind("<FocusOut>", username_on_leave)
username_entry.bind("<Return>", user_handle_enter)

score_label = Label(stundet_info_frame, text="Student Score", font=("Arial", 12, "bold"))
score_label.grid(row=2, column=0, sticky="w", padx=10, pady=10)

score_entry = Entry(stundet_info_frame, font=("Arial", 12))
score_entry.grid(row=3, column=0, sticky="nwe", padx=10)
score_entry.insert(0, "Score")
score_entry.bind("<FocusIn>", score_on_click)
score_entry.bind("<FocusOut>", score_on_leave)
score_entry.bind("<Return>", user_handle_enter)

save_button = Button(stundet_info_frame, text="Save", font=("Arial", 12), width=7, command=save_info_to_excel)
save_button.grid(row=4, column=0, padx=20, pady=10, sticky="w")

update_button = Button(stundet_info_frame, text="Update", font=("Arial", 12), width=7, command=update_info_to_excel)
update_button.grid(row=4, column=0, padx=20, pady=10, sticky="e")


# ================= WINDOW SCORE TRACKER STARTER ================= #
window.mainloop()