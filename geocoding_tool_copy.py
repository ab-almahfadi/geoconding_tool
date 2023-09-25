import pandas as pd
from geopy.geocoders import Nominatim
import time
import requests
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading

MAX_RETRIES = 3

# Function to perform geocoding with retries
def geocode_location(latitude, longitude, row_index, df):
    geolocator = Nominatim(user_agent="geocoding_tool")
    retries = 0
    while retries < MAX_RETRIES:
        try:
            location = geolocator.reverse((latitude, longitude), language='en', timeout=10)
            address = location.address if location else None
            log_text.insert(tk.END, f"Row {row_index}: Geocoded successfully - Address: {address}\n")
            df.loc[row_index, 'Address'] = address  # Update DataFrame with the geocoded address
            return address
        except requests.exceptions.ReadTimeout:
            log_text.insert(tk.END, f"Row {row_index}: Read timeout. Retrying ({retries + 1}/{MAX_RETRIES})...\n")
            retries += 1
            time.sleep(1)  # Wait for a short duration before retrying
    log_text.insert(tk.END, f"Row {row_index}: Max retries exceeded. Unable to fetch data.\n")
    return None

# Load the Excel file into a DataFrame
def load_excel_data(file_path):
    return pd.read_excel(file_path)

# Save the geocoded addresses to a new Excel file
def save_geocoded_data(df, output_file):
    df.to_excel(output_file, index=False)

# Function to handle geocoding and update the DataFrame
def geocode_and_save():
    input_file = input_file_entry.get()
    output_file = output_file_entry.get()

    if not input_file or not output_file:
        messagebox.showerror("Error", "Please select both input and output files.")
        return

    try:
        # Load the data from the Excel file
        df = load_excel_data(input_file)
        df['Address'] = ""  # Add an empty 'Address' column

        total_rows = len(df)
        for index, row in df.iterrows():
            geocode_location(row['Latitude'], row['Longitude'], index, df)  # Pass df to update the DataFrame

            progress_label.config(text=f"Progress: {index + 1}/{total_rows}")
            root.update_idletasks()  # Update the GUI

        # Save the geocoded data to a new Excel file
        save_geocoded_data(df, output_file)
        messagebox.showinfo("Success", "Geocoding completed. Geocoded data saved to " + output_file)
    except Exception as e:
        messagebox.showerror("Error", "An error occurred during geocoding: " + str(e))

# Function to run geocoding in a separate thread
def run_geocoding_thread():
    threading.Thread(target=geocode_and_save).start()

# Create the GUI
root = tk.Tk()
root.title("Geocoding Tool")

# Input file selection
input_file_label = tk.Label(root, text="Input Excel File:")
input_file_label.pack()
input_file_entry = tk.Entry(root, width=50)
input_file_entry.pack()

def browse_input_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    input_file_entry.delete(0, tk.END)
    input_file_entry.insert(0, file_path)

browse_input_button = tk.Button(root, text="Browse", command=browse_input_file)
browse_input_button.pack()

# Output file selection
output_file_label = tk.Label(root, text="Output Excel File:")
output_file_label.pack()
output_file_entry = tk.Entry(root, width=50)
output_file_entry.pack()

def browse_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    output_file_entry.delete(0, tk.END)
    output_file_entry.insert(0, file_path)

browse_output_button = tk.Button(root, text="Browse", command=browse_output_file)
browse_output_button.pack()

# Run geocoding button
geocode_button = tk.Button(root, text="Run Geocoding", command=run_geocoding_thread)
geocode_button.pack()

# Progress indicator (label-based progress)
progress_label = tk.Label(root, text="Progress: 0/0")
progress_label.pack()

# Log display
log_label = tk.Label(root, text="Log:")
log_label.pack()
log_text = ScrolledText(root, width=50, height=10)
log_text.pack()

root.mainloop()
