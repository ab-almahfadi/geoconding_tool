import tkinter as tk
from tkinter import filedialog
from geopy.geocoders import Nominatim
import pandas as pd
from tqdm import tqdm
import time
import requests

MAX_RETRIES = 3

# Function to perform geocoding with retries
def geocode_location_with_retry(latitude, longitude):
    geolocator = Nominatim(user_agent="geocoding_tool")
    retries = 0
    while retries < MAX_RETRIES:
        try:
            location = geolocator.reverse((latitude, longitude), language='en', timeout=10)
            return location.address if location else None
        except requests.exceptions.ReadTimeout as e:
            error_message = f"Read timeout. Retrying ({retries + 1}/{MAX_RETRIES})..."
            print(error_message)
            retries += 1
            time.sleep(1)  # Wait for a short duration before retrying
    error_message = "Max retries exceeded. Unable to fetch data."
    print(error_message)
    return None

# Modify the geocode_from_excel function to use geocode_location_with_retry
def geocode_from_excel():
    input_file = file_path_entry.get()
    output_file = output_file_entry.get()

    try:
        df = pd.read_excel(input_file)
        
        # Update progress bar while iterating through the DataFrame
        progress_bar['maximum'] = len(df)
        for i, row in tqdm(df.iterrows(), total=len(df)):
            latitude, longitude = row['Latitude'], row['Longitude']
            tracking_message = f"Geocoding Row {i+1}: Latitude={latitude}, Longitude={longitude}..."
            print(tracking_message)
            result_label.config(text=tracking_message)
            df.at[i, 'Address'] = geocode_location_with_retry(latitude, longitude)
            progress_bar['value'] = i + 1
            root.update_idletasks()  # Update the GUI

        df.to_excel(output_file, index=False)
        result_label.config(text='Geocoding completed. Geocoded data saved to ' + output_file)
    except Exception as e:
        error_message = 'Error: ' + str(e)
        print(error_message)
        result_label.config(text=error_message)

# Function to handle file selection for input Excel file
def browse_file():
    file_path = filedialog.askopenfilename()
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(0, file_path)

# Create the main GUI window
root = tk.Tk()
root.title("Geocoding Tool")

# Create and position GUI elements
file_path_label = tk.Label(root, text="Input Excel File:")
file_path_label.grid(row=0, column=0, padx=10, pady=5)

file_path_entry = tk.Entry(root, width=40)
file_path_entry.grid(row=0, column=1, padx=10, pady=5)

browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=5, pady=5)

latitude_label = tk.Label(root, text="Latitude:")
latitude_label.grid(row=1, column=0, padx=10, pady=5)

latitude_entry = tk.Entry(root)
latitude_entry.grid(row=1, column=1, padx=10, pady=5)

longitude_label = tk.Label(root, text="Longitude:")
longitude_label.grid(row=2, column=0, padx=10, pady=5)

longitude_entry = tk.Entry(root)
longitude_entry.grid(row=2, column=1, padx=10, pady=5)

geocode_button = tk.Button(root, text="Geocode", command=geocode_location)
geocode_button.grid(row=3, column=0, columnspan=2, pady=10)

address_label = tk.Label(root, text="Address:")
address_label.grid(row=4, column=0, padx=10, pady=5)

address_entry = tk.Entry(root, width=40)
address_entry.grid(row=4, column=1, padx=10, pady=5)

output_file_label = tk.Label(root, text="Output Excel File:")
output_file_label.grid(row=5, column=0, padx=10, pady=5)

output_file_entry = tk.Entry(root, width=40)
output_file_entry.grid(row=5, column=1, padx=10, pady=5)

save_file_button = tk.Button(root, text="Save As", command=save_file)
save_file_button.grid(row=5, column=2, padx=5, pady=5)

geocode_from_excel_button = tk.Button(root, text="Geocode from Excel", command=geocode_from_excel)
geocode_from_excel_button.grid(row=6, column=0, columnspan=3, pady=10)

result_label = tk.Label(root, text="")
result_label.grid(row=7, column=0, columnspan=3, pady=10)

# Add a progress bar
progress_bar = tk.ttk.Progressbar(root, mode='determinate', length=400)
progress_bar.grid(row=8, column=0, columnspan=3, pady=10)

# Run the GUI loop
root.mainloop()
