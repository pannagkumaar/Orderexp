import pandas as pd
import customtkinter as ctk
import pyperclip
from tkinter import filedialog, messagebox
from datetime import datetime, timedelta
import os

# Global variables
file_path = None
df = None

# Core Logic

def process_meal(meal_type, df):
    today_col = f"Today {meal_type}"
    default_col = f"Default {meal_type}"
    skip_col = f"Skip {meal_type} Until"
    result = []
    today = datetime.today().date()

    for index, row in df.iterrows():
        name = str(row.get("Name", "")).strip()
        address = str(row.get("Address", "")).strip()
        today_val = str(row.get(today_col, "")).strip()
        default_val = str(row.get(default_col, "")).strip()
        skip_until = row.get(skip_col, "")

        # Normalize NaN values
        if today_val.lower() in ["nan", "none"]:
            today_val = ""
        if default_val.lower() in ["nan", "none"]:
            default_val = ""

        # Handle existing skip dates (don't modify if set)
        if isinstance(skip_until, (pd.Timestamp, datetime)) and skip_until.date() > today:
            continue  # Skip meal due to existing date

        # Process -N logic
        if today_val.startswith("-") and today_val[1:].isdigit():
            days_to_skip = int(today_val[1:])
            future_date = today + timedelta(days=days_to_skip)
            df.at[index, skip_col] = future_date.strftime("%Y-%m-%d")
            df.at[index, today_col] = ""  # Clear Today <Meal> after processing
            continue  # Don't process this meal today

        # Handle temporary meal orders
        if today_val in ["-", "no"]:
            df.at[index, today_col] = ""  # Clear it after processing
            continue  # Skip this meal

        meal = today_val if today_val else default_val
        if meal:
            result.append(f"{name} ({address}) -> {meal}")
        if today_val:  # One-day order, clear after use
            df.at[index, today_col] = ""
    return "\n".join(result)

def copy_meal(meal_type):
    global df
    if df is None:
        messagebox.showerror("Error", "No file loaded. Please load an Excel file first.")
        return

    text = process_meal(meal_type, df)
    pyperclip.copy(text)
    
    if text:
        copied_label.configure(text=text, text_color="green")
        status_label.configure(text=f"Copied {meal_type} to clipboard", text_color="blue")
    else:
        copied_label.configure(text="Nothing copied.", text_color="gray")
        status_label.configure(text="No meals copied.", text_color="gray")

    save_updated_excel()

def save_updated_excel():
    global file_path
    if file_path and df is not None:
        df.to_excel(file_path, index=False)
        status_label.configure(text="Updated Excel saved.", text_color="blue")

def load_excel():
    global df, file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if file_path:
        df = pd.read_excel(file_path).fillna("")
        filename_label.configure(text=f"Loaded File: {os.path.basename(file_path)}", text_color="black")
        status_label.configure(text="File loaded successfully", text_color="blue")
        copied_label.configure(text="")

# UI Setup
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Meal Copy Tool")
app.geometry("600x400")

frame = ctk.CTkFrame(app)
frame.pack(pady=10, fill="both", expand=True)

load_button = ctk.CTkButton(frame, text="Load Excel", command=load_excel)
load_button.pack(pady=5)

button_frame = ctk.CTkFrame(frame)
button_frame.pack(pady=5, fill="x")

bf_button = ctk.CTkButton(button_frame, text="Copy Breakfast", command=lambda: copy_meal("BF"))
bf_button.pack(side="left", padx=5, expand=True)

lunch_button = ctk.CTkButton(button_frame, text="Copy Lunch", command=lambda: copy_meal("Lunch"))
lunch_button.pack(side="left", padx=5, expand=True)

dinner_button = ctk.CTkButton(button_frame, text="Copy Dinner", command=lambda: copy_meal("Dinner"))
dinner_button.pack(side="left", padx=5, expand=True)

filename_label = ctk.CTkLabel(frame, text="No file loaded", text_color="gray")
filename_label.pack(pady=2)

status_label = ctk.CTkLabel(frame, text="", text_color="red")
status_label.pack(pady=2)

copied_label = ctk.CTkLabel(frame, text="", wraplength=500, justify="left")
copied_label.pack(pady=2)

app.mainloop()
