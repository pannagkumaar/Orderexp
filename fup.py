import pandas as pd
import customtkinter as ctk
import pyperclip
from tkinter import filedialog
from datetime import datetime, timedelta
import os
import tkinter as tk
from tkcalendar import DateEntry

# Global DataFrame & file path
file_path = None
df = None
autocomplete_list = []
entry_widgets = {}

# Color scheme
COLORS = {
    "primary": "#4A6572",
    "secondary": "#5D7B8A",
    "accent": "#FF9E4A",
    "background": "#F0F0F0",
    "text": "#344955",
    "success": "#4CAF50",
    "warning": "#FFC107"
}

def process_meal(meal_type, df):
    """Generate meal orders with skip logic."""
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

        # Clean values
        today_val = "" if today_val.lower() in ["nan", "none"] else today_val
        default_val = "" if default_val.lower() in ["nan", "none"] else default_val

        # Handle skips
        if isinstance(skip_until, pd.Timestamp) and today < skip_until.date():
            continue

        # Process skip commands
        if today_val.startswith("-") and today_val[1:].isdigit():
            days_to_skip = int(today_val[1:])
            df.at[index, skip_col] = (today + timedelta(days=days_to_skip)).strftime("%Y-%m-%d")
            df.at[index, today_col] = ""
            continue

        if today_val in ["-", "no"]: continue

        meal = today_val if today_val else default_val
        if meal:
            result.append(f"{name} ({address}) -> {meal}")

    return "\n".join(result)

def copy_meal(meal_type):
    """Copy meal to clipboard."""
    global df
    if df is None:
        status_label.configure(text="No file loaded!", text_color=COLORS["warning"])
        return

    text = process_meal(meal_type, df)
    pyperclip.copy(text)
    copied_label.configure(text=text if text else "Nothing copied.", text_color=COLORS["success"])
    status_label.configure(text=f"Copied {meal_type} to clipboard", text_color=COLORS["accent"])
    save_updated_excel()

def save_updated_excel():
    """Save DataFrame to Excel."""
    global file_path
    if file_path and df is not None:
        df.to_excel(file_path, index=False)

def load_excel():
    """Load Excel file."""
    global df, file_path
    path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if path:
        file_path = path
        df_temp = pd.read_excel(file_path).fillna("")
        df_temp.reset_index(drop=True, inplace=True)
        df = df_temp
        filename_label.configure(text=f"Loaded: {os.path.basename(file_path)}", text_color=COLORS["text"])
        status_label.configure(text="File loaded successfully", text_color=COLORS["success"])
        copied_label.configure(text="")
        update_autocomplete_list()
        search_combobox.configure(values=autocomplete_list)

def update_autocomplete_list():
    """Update search list."""
    global df
    autocomplete_list.clear()
    if df is not None:
        for _, row in df.iterrows():
            name = str(row.get("Name", "")).strip()
            address = str(row.get("Address", "")).strip()
            if name: autocomplete_list.append(f"{name} ({address})")

def update_suggestions(event=None):
    """Update combobox suggestions."""
    search_text = search_var.get().lower()
    filtered = [s for s in autocomplete_list if all(v in s.lower() for v in search_text.split())]
    search_combobox.configure(values=filtered)
    if len(search_text) >= 2:
        search_combobox.event_generate('<Down>')

def on_combobox_select(event):
    """Handle combobox selection."""
    selected = search_combobox.get()
    search_var.set(selected)
    clear_entries()
    populate_entries(selected)

def clear_entries():
    """Clear input fields."""
    for widget in entry_widgets.values():
        if isinstance(widget, ctk.CTkEntry):
            widget.delete(0, 'end')
        elif isinstance(widget, DateEntry):
            widget.set_date(datetime.today().date())

def populate_entries(selected):
    """Populate fields with data."""
    if df is None: return
    
    for index, row in df.iterrows():
        name = str(row.get("Name", "")).strip()
        address = str(row.get("Address", "")).strip()
        if f"{name} ({address})" == selected:
            for meal in ["BF", "Lunch", "Dinner"]:
                entry_widgets[f"{meal}_today"].insert(0, str(row.get(f"Today {meal}", "")))
                entry_widgets[f"{meal}_default"].insert(0, str(row.get(f"Default {meal}", "")))
                
                skip_date = row.get(f"Skip {meal} Until", datetime.today().date())
                if isinstance(skip_date, pd.Timestamp):
                    skip_date = skip_date.date()
                entry_widgets[f"{meal}_skip"].set_date(skip_date)
            break

def update_meal():
    """Update meal preferences."""
    global df
    if df is None or search_var.get().strip() == "":
        status_label.configure(text="No selection made!", text_color=COLORS["warning"])
        return

    selected = search_var.get()
    for index, row in df.iterrows():
        name = str(row.get("Name", "")).strip()
        address = str(row.get("Address", "")).strip()
        if f"{name} ({address})" == selected:
            for meal in ["BF", "Lunch", "Dinner"]:
                today_val = entry_widgets[f"{meal}_today"].get()
                df.at[index, f"Today {meal}"] = today_val if today_val else ""
                
                default_val = entry_widgets[f"{meal}_default"].get()
                df.at[index, f"Default {meal}"] = default_val if default_val else ""
                
                skip_date = entry_widgets[f"{meal}_skip"].get_date()
                df.at[index, f"Skip {meal} Until"] = skip_date.strftime("%Y-%m-%d")
            
            status_label.configure(text=f"Updated: {name}", text_color=COLORS["success"])
            save_updated_excel()
            clear_entries()
            break

# ========== UI SETUP ==========
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Anjali Meal Management")
app.geometry("900x800")
app.grid_columnconfigure(0, weight=1)
app.grid_rowconfigure(0, weight=1)

# Main container
main_frame = ctk.CTkFrame(app, corner_radius=15, fg_color=COLORS["background"])
main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

# Header
header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
header_frame.pack(pady=15, fill="x")

header_label = ctk.CTkLabel(header_frame,
                           text="🥗 Anjali Meal Management",
                           font=("Arial", 24, "bold"),
                           text_color=COLORS["primary"])
header_label.pack(side="left", padx=20)

# File controls
file_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
file_frame.pack(pady=10, fill="x")

load_button = ctk.CTkButton(file_frame,
                           text="📂 Load Excel File",
                           command=load_excel,
                           fg_color=COLORS["secondary"],
                           hover_color=COLORS["primary"],
                           width=200,
                           height=35)
load_button.pack(side="left", padx=20)

filename_label = ctk.CTkLabel(file_frame,
                             text="No file loaded",
                             text_color=COLORS["text"],
                             font=("Arial", 12))
filename_label.pack(side="left", padx=10)

# Copy buttons
copy_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
copy_frame.pack(pady=15, fill="x")

button_style = {
    "fg_color": COLORS["accent"],
    "hover_color": "#FF8F00",
    "width": 120,
    "height": 35
}

bf_button = ctk.CTkButton(copy_frame, text="BF", **button_style, command=lambda: copy_meal("BF"))
lunch_button = ctk.CTkButton(copy_frame, text="Lunch", **button_style, command=lambda: copy_meal("Lunch"))
dinner_button = ctk.CTkButton(copy_frame, text="Dinner", **button_style, command=lambda: copy_meal("Dinner"))

bf_button.pack(side="left", padx=10)
lunch_button.pack(side="left", padx=10)
dinner_button.pack(side="left", padx=10)

# Status area
status_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
status_frame.pack(pady=10, fill="x")

status_label = ctk.CTkLabel(status_frame,
                           text="Ready",
                           text_color=COLORS["text"],
                           font=("Arial", 12))
status_label.pack(side="left", padx=20)

copied_label = ctk.CTkLabel(status_frame,
                           text="",
                           wraplength=700,
                           justify="left",
                           font=("Arial", 12),
                           text_color=COLORS["success"])
copied_label.pack(pady=5, fill="x")

# Update section
update_frame = ctk.CTkFrame(main_frame,
                           corner_radius=12,
                           border_width=2,
                           border_color=COLORS["secondary"])
update_frame.pack(pady=20, fill="x", padx=20)

# Search combobox
search_frame = ctk.CTkFrame(update_frame, fg_color="transparent")
search_frame.pack(pady=15, fill="x", padx=20)

search_var = ctk.StringVar()
search_combobox = ctk.CTkComboBox(search_frame,
                                 variable=search_var,
                                 values=autocomplete_list,
                                 dropdown_fg_color=COLORS["background"],
                                 dropdown_hover_color=COLORS["secondary"],
                                 button_color=COLORS["accent"],
                                 width=400,
                                 height=35)
search_combobox.pack(pady=5)
search_combobox.bind("<KeyRelease>", update_suggestions)
search_combobox.bind("<<ComboboxSelected>>", on_combobox_select)

# Input grid
input_grid = ctk.CTkFrame(update_frame, fg_color="transparent")
input_grid.pack(pady=15, padx=20)

# Headers
headers = ["Meal", "Today", "Default", "Skip Until"]
for col, header in enumerate(headers):
    ctk.CTkLabel(input_grid,
                text=header,
                font=("Arial", 12, "bold"),
                text_color=COLORS["primary"]).grid(row=0, column=col, padx=10, pady=5)

# Meal rows
meals = ["BF", "Lunch", "Dinner"]
for row, meal in enumerate(meals, 1):
    ctk.CTkLabel(input_grid,
                text=meal,
                font=("Arial", 12),
                text_color=COLORS["text"]).grid(row=row, column=0, padx=10, sticky="e")
    
    today_entry = ctk.CTkEntry(input_grid, width=150, height=30, border_color=COLORS["secondary"])
    today_entry.grid(row=row, column=1, padx=10, pady=5)
    entry_widgets[f"{meal}_today"] = today_entry
    
    default_entry = ctk.CTkEntry(input_grid, width=150, height=30, border_color=COLORS["secondary"])
    default_entry.grid(row=row, column=2, padx=10, pady=5)
    entry_widgets[f"{meal}_default"] = default_entry
    
    skip_picker = DateEntry(input_grid,
                          date_pattern="yyyy-mm-dd",
                          background=COLORS["background"],
                          foreground=COLORS["text"],
                          borderwidth=1)
    skip_picker.grid(row=row, column=3, padx=10, pady=5)
    entry_widgets[f"{meal}_skip"] = skip_picker

# Update button
update_button = ctk.CTkButton(update_frame,
                             text="💾 Save Changes",
                             command=update_meal,
                             fg_color=COLORS["accent"],
                             hover_color="#FF8F00",
                             width=200,
                             height=40)
update_button.pack(pady=20)

app.mainloop()