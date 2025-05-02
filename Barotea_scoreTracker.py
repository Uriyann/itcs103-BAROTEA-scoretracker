import re
from PIL import Image
from tkinter import messagebox
from customtkinter import *
import customtkinter as ctk
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


""" ================= MAIN WINDOW SCORE TRACKER ================= """
window = CTk()
window.title("Score Tracker")
window.geometry("450x450")
window.resizable(False, False)
ctk.set_appearance_mode("dark")

# ================= FUNCTIONS ================= #
# Saving The Info To Excel Function
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
        ws.append(["Name", "Score", "Remarks"])
        ws.append(["", "", ""])

    if score >= 75 and score <= 100:
        ws.append([name, score, "Passed"])
        messagebox.showinfo(title="Success", message="You Passed The Exam.")
    elif score >= 50 and score < 75:
        ws.append([name, score, "Failed"])
        messagebox.showinfo(title="Success", message="You Failed The Exam.")
    else:
        ws.append([name, score, "You Need To Study More"])
        messagebox.showinfo(title="Success", message="Get Good On Your Studies.")

    wb.save("student_scores.xlsx")
    format_excel()

    username_entry.delete(0, END)
    score_entry.delete(0, END)

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
                if score >= 75 and score <= 100:
                    row[2].value = "Passed"
                elif score >= 50 and score < 75:
                    row[2].value = "Failed"
                else:
                    row[2].value = "You Need To Study More"
                wb.save("student_scores.xlsx")

                messagebox.showinfo(title="Success", message="Score Updated Successfully.")
                return True

        messagebox.showerror(title="Error", message="User not found.")
        return False

    except Exception as e:
        messagebox.showerror(title="Error", message=str(e))
        return False

# Show Data Function
def show_data_excel():
    try:

        wb = load_workbook("student_scores.xlsx")
        ws = wb["Userdata"]

        data_window = CTkToplevel(window)
        data_window.title("Data")
        data_window.geometry("350x400")
        data_window.resizable(False, False)
        data_window.configure(fg_color="#2b2b2b")

        toplevel_frame = CTkFrame(data_window, fg_color="#2b2b2b")
        toplevel_frame.pack(pady=10, padx=10, fill="both", expand=True)

        for i, row in enumerate(ws.iter_rows(values_only=True)):
            for j, value in enumerate(row):
                label = CTkLabel(toplevel_frame, text=value, padx=6, pady=3, font=("Arial", 16), fg_color="#2b2b2b")
                label.grid(row=i, column=j)

        destroy_button = CTkButton(data_window, text="Close", command=data_window.destroy, font=("Arial", 16))
        destroy_button.pack(side="bottom", pady=10)

        

    except Exception as e:
        messagebox.showerror(title="Error", message=str(e))

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
        username_entry.delete(0, END)
        return False
    
# Score Debugger
def validate_score_int():
    try:
        value = int(score_entry.get())
        if value >= 0 and value <= 100:
            return True
        
        else:
            messagebox.showerror(title="Error", message="Please enter a valid score.")
            score_entry.delete(0, END)
            return False
        
    except ValueError:
        messagebox.showerror(title="Error", message="Please enter a real number.")
        score_entry.delete(0, END)
        return False

# ================= ENTRY EVENTS ================= #
# User Handle Enter Event
def user_handle_enter(event):
    if not save_info_to_excel():
        return

# ================= MAIN FRAMES ================= #
main_frame = CTkFrame(window, border_width= 3, corner_radius= 15)
main_frame.place(rely=0.5, relx=0.5, anchor=CENTER)

background_image = Image.open(r".\588ee4b393c082b8422e7307e0ece441.png")
bg_img = CTkImage(light_image=background_image, dark_image=background_image, size=(90, 90))
bg_label = CTkLabel(main_frame, image=bg_img, text="")
bg_label.grid(row=0, column=0, sticky="nw", pady= 15, padx= 55)

background_image = Image.open(r".\e4140a754afb75b19463aefe149e65ab.png")
bg_img = CTkImage(light_image=background_image, dark_image=background_image, size=(90, 90))
bg_label = CTkLabel(main_frame, image=bg_img, text="")
bg_label.grid(row=0, column=0, sticky="ne", pady= 15, padx= 55)

# ================= WIDGETS ================= #
title_label = CTkLabel(main_frame, text= "User Score Tracker", font= ("Times New Roman bold", 35))
title_label.grid(row=1, column=0, sticky="n", pady= 5, padx= 15)

username_label = CTkLabel(main_frame, text="Student Name:", font=("Arial", 18, "bold"))
username_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")

username_entry = CTkEntry(main_frame, font=("Arial", 18), border_width=0, placeholder_text="Name", height= 10, width= 10, fg_color= "transparent", bg_color= "transparent")
username_entry.grid(row=3, column=0, sticky="nwe", padx=10)
username_entry.bind("<Return>", user_handle_enter)

user_line = CTkFrame(main_frame, width=300, height=3)
user_line.grid(row=4, column=0, sticky= S)

score_label = CTkLabel(main_frame, text="Student Score:", font=("Arial", 18, "bold"))
score_label.grid(row=5, column=0, sticky="w", padx=10, pady=5)

score_entry = CTkEntry(main_frame, font=("Arial", 18), border_width=0, placeholder_text="Score", height= 10, width= 10, fg_color= "transparent", bg_color= "transparent")
score_entry.grid(row=6, column=0, sticky="nwe", padx=10)
score_entry.bind("<Return>", user_handle_enter)

score_line = CTkFrame(main_frame, width=300, height=3)
score_line.grid(row=7, column=0, sticky= S)

save_button = CTkButton(main_frame, text="Save", font=("Arial", 15), width=13, height=30, command=save_info_to_excel)
save_button.grid(row=8, column=0, padx=20, pady=20, sticky="w")

update_button = CTkButton(main_frame, text="Update", font=("Arial", 15), width=13, height=30, command=update_info_to_excel)
update_button.grid(row=8, column=0, padx=20, pady=20, sticky="e")

show_data = CTkButton(main_frame, text="Show Data", font=("Arial", 15), width=13, height=40, command=show_data_excel)
show_data.grid(row=8, column=0, padx=20, pady=20)

# ================= WINDOW SCORE TRACKER STARTER ================= #
window.mainloop()