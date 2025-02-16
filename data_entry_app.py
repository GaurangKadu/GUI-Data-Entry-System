import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import random
from datetime import datetime

# Function to generate a unique 10-digit number
def generate_unique_number():
    fixed_part = "50034"
    random_part = ''.join([str(random.randint(0, 9)) for _ in range(5)])
    return fixed_part + random_part

# Function to save data to Excel
def save_data():
    global entry_count
    data = {
        "Unique No.": generate_unique_number(),
        "PO": po_entry.get(),
        "Invoice No.": invoice_entry.get(),
        "GRN": grn_entry.get(),
        "Gate Pass No.": gate_pass_entry.get(),
        "Check List No.": check_list_entry.get(),
        "Scroll No.": scroll_entry.get(),
    }

    df = pd.DataFrame([data])

    try:
        existing_df = pd.read_excel("data.xlsx")
        updated_df = pd.concat([existing_df, df], ignore_index=True)
    except FileNotFoundError:
        updated_df = df

    updated_df.to_excel("data.xlsx", index=False)

    entry_count += 1
    counter_label.config(text=f"Entries: {entry_count}")

    messagebox.showinfo("Success", "Data saved successfully!")

# Function to view data
def view_data():
    try:
        df = pd.read_excel("data.xlsx")
        latest_entries = df.tail(5)

        view_window = tk.Toplevel(root)
        view_window.title("View Data (Latest 5 Entries)")

        tree = ttk.Treeview(view_window, columns=list(latest_entries.columns), show="headings")
        tree.pack(fill="both", expand=True)

        for col in latest_entries.columns:
            tree.heading(col, text=col)

        for _, row in latest_entries.iterrows():
            tree.insert("", "end", values=list(row))

    except FileNotFoundError:
        messagebox.showinfo("Data", "No data found!")

# Function to update the clock
def update_clock():
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    clock_label.config(text=now)
    root.after(1000, update_clock)

# Function to move focus to the next entry field when Enter is pressed
def focus_next(event, field_list, current_index):
    if current_index < len(field_list) - 1:
        field_list[current_index + 1].focus_set()
    return "break"

# Initialize the main window
root = tk.Tk()
root.title("Data Entry Form")
root.geometry("400x500")

entry_count = 0

clock_label = tk.Label(root, font=("Arial", 12), fg="blue")
clock_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
update_clock()

fields = ["PO", "Invoice No.", "GRN", "Gate Pass No.", "Check List No.", "Scroll No."]
entries = []

for i, field in enumerate(fields):
    label = tk.Label(root, text=field + ":")
    label.grid(row=i+1, column=0, padx=10, pady=5, sticky="w")
    entry = tk.Entry(root, width=30)
    entry.grid(row=i+1, column=1, padx=10, pady=5)
    entries.append(entry)
    entry.bind("<Return>", lambda event, idx=i: focus_next(event, entries, idx))  # Enter key navigation

po_entry, invoice_entry, grn_entry, gate_pass_entry, check_list_entry, scroll_entry = entries

save_button = tk.Button(root, text="Save Data", command=save_data)
save_button.grid(row=len(fields)+1, column=0, columnspan=2, pady=10)

view_button = tk.Button(root, text="View Data", command=view_data)
view_button.grid(row=len(fields)+2, column=0, columnspan=2, pady=10)

counter_label = tk.Label(root, text="Entries: 0", font=("Arial", 10), fg="green")
counter_label.grid(row=len(fields)+3, column=1, padx=10, pady=10, sticky="e")

root.mainloop()
