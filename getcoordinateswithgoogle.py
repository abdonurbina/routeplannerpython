import googlemaps
import pandas as pd


def geocode_address_google(api_key, address):
    """Get latitude and longitude using Google Maps API."""
    gmaps = googlemaps.Client(key=api_key)
    geocode_result = gmaps.geocode(address)
    if geocode_result:
        location = geocode_result[0]['geometry']['location']
        return location['lat'], location['lng']
    else:
        return None, None


def process_addresses_google(api_key, input_file, output_file):
    """Read addresses from Excel, geocode them, and save results."""
    data = pd.read_excel(input_file)
    if 'Address' not in data.columns:
        raise ValueError("Excel file must contain an 'Address' column.")

    latitudes, longitudes = [], []
    for address in data['Address']:
        lat, lon = geocode_address_google(api_key, address)
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
    # Replace with your actual API key
    google_api_key = ""

    input_file = "addresses.xlsx"  # Excel file with 'Address' column
    output_file = "geocoded_addresses_google.xlsx"

    process_addresses_google(google_api_key, input_file, output_file)
