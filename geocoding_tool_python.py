import pandas as pd
from geopy.geocoders import Nominatim
import time
import requests

MAX_RETRIES = 3

# Function to perform geocoding with retries
def geocode_location(latitude, longitude, row_index):
    geolocator = Nominatim(user_agent="geocoding_tool")
    retries = 0
    while retries < MAX_RETRIES:
        try:
            location = geolocator.reverse((latitude, longitude), language='en', timeout=10)
            address = location.address if location else None
            print(f"Row {row_index}: Geocoded successfully - Address: {address}")
            return address
        except requests.exceptions.ReadTimeout:
            print(f"Row {row_index}: Read timeout. Retrying ({retries + 1}/{MAX_RETRIES})...")
            retries += 1
            time.sleep(1)  # Wait for a short duration before retrying
    print(f"Row {row_index}: Max retries exceeded. Unable to fetch data.")
    return None

# Load the Excel file into a DataFrame
def load_excel_data(file_path):
    return pd.read_excel(file_path)

# Save the geocoded addresses to a new Excel file
def save_geocoded_data(df, output_file):
    df.to_excel(output_file, index=False)

# Main function to initiate the geocoding process
def main():
    input_file = 'test1.xlsx'  # Change to your input Excel file
    output_file = 'output_1.xlsx'  # Change to your desired output file name

    # Load the data from the Excel file
    df = load_excel_data(input_file)

    # Geocode each row and obtain the address
    df['Address'] = df.apply(lambda row: geocode_location(row['Latitude'], row['Longitude'], row.name), axis=1)

    # Save the geocoded data to a new Excel file
    save_geocoded_data(df, output_file)
    print('Geocoding completed. Geocoded data saved to', output_file)

if __name__ == "__main__":
    main()
