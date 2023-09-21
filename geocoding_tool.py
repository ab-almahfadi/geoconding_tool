import pandas as pd
from geopy.geocoders import Nominatim

# Function to perform geocoding
def geocode_location(latitude, longitude):
    geolocator = Nominatim(user_agent="geocoding_tool")
    location = geolocator.reverse((latitude, longitude), language='en')
    return location.address if location else None

# Load the Excel file into a DataFrame
def load_excel_data(file_path):
    return pd.read_excel(file_path)

# Save the geocoded addresses to a new Excel file
def save_geocoded_data(df, output_file):
    df.to_excel(output_file, index=False)

# Main function to initiate the geocoding process
def main():
    input_file = 'input.xlsx'  # Change to your input Excel file
    output_file = 'output_geocoded.xlsx'  # Change to your desired output file name

    # Load the data from the Excel file
    df = load_excel_data(input_file)

    # Geocode each row and obtain the address
    df['Address'] = df.apply(lambda row: geocode_location(row['Latitude'], row['Longitude']), axis=1)

    # Save the geocoded data to a new Excel file
    save_geocoded_data(df, output_file)
    print('Geocoding completed. Geocoded data saved to', output_file)

if __name__ == "__main__":
    main()
