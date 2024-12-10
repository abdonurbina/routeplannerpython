from geopy.geocoders import Nominatim
import pandas as pd


def geocode_address(address):
    """Get latitude and longitude for a given address."""
    geolocator = Nominatim(user_agent="route-planner")
    location = geolocator.geocode(address)
    if location:
        return location.latitude, location.longitude
    else:
        return None, None


def process_addresses(file_path, output_file):
    """Read addresses from an Excel file, geocode them, and save results."""
    data = pd.read_excel(file_path)
    if 'Address' not in data.columns:
        raise ValueError("Excel file must contain an 'Address' column.")

    latitudes, longitudes = [], []
    for address in data['Address']:
        lat, lon = geocode_address(address)
        latitudes.append(lat)
        longitudes.append(lon)
        print(f"Processed: {address} -> ({lat}, {lon})")

    # Add latitude and longitude to the dataframe
    data['Latitude'] = latitudes
    data['Longitude'] = longitudes

    # Save to a new Excel file
    data.to_excel(output_file, index=False)
    print(f"Geocoded addresses saved to {output_file}")


# Example usage
if __name__ == "__main__":
    input_file = "addresses.xlsx"  # Excel file with an 'Address' column
    output_file = "geocoded_addresses.xlsx"
    process_addresses(input_file, output_file)

