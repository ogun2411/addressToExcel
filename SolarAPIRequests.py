import pandas as pd
import requests
import os
import json

# Function to make API request with Solar API
def fetch_data_from_solar_api(latitude, longitude, name, api_key, output_dir, df):
    try:
        # Construct the full API URL with parameters
        full_url = (
            f"https://solar.googleapis.com/v1/buildingInsights:findClosest"
            f"?location.latitude={latitude}&location.longitude={longitude}"
            f"&requiredQuality=HIGH&key={api_key}"
        )
        response = requests.get(full_url)

        # Check for a successful response
        if response.status_code == 200:
            json_data = response.json()

            # Create a safe filename using the location name
            safe_name = "".join(c if c.isalnum() or c in (" ", "_", "-") else "_" for c in name)
            output_file = os.path.join(output_dir, f"{safe_name}.json")

            # Save JSON data to a file
            with open(output_file, 'w') as file:
                json.dump(json_data, file, indent=4)
            print(f"Data for {name} saved to {output_file}.")

            # Parse specific data from the JSON response
            num_panels = json_data.get('solarPotential', {}).get('maxArrayPanelsCount', None)
            solar_area = json_data.get('solarPotential', {}).get('maxArrayAreaMeters2', None)

            # Extract the yearly energy corresponding to maxArrayPanelsCount
            solar_panel_configs = json_data.get('solarPotential', {}).get('solarPanelConfigs', [])
            yearly_energy = None
            for config in solar_panel_configs:
                if config.get('panelsCount') == num_panels:
                    yearly_energy = config.get('yearlyEnergyDcKwh', None)
                    break  # Stop searching once the correct configuration is found

            # Debug the parsed values
            print(f"Parsed data for {name}: NumPanels={num_panels}, SolarArea={solar_area}, YearlyEnergy={yearly_energy}")
            
            # Update the DataFrame with the parsed values
            df.loc[df['Name'] == name, 'NumPanels'] = num_panels
            df.loc[df['Name'] == name, 'YearlyEnergy'] = yearly_energy
            df.loc[df['Name'] == name, 'SolarArea'] = solar_area
        else:
            print(f"Failed to fetch data for {name}. HTTP Status: {response.status_code}")
    except Exception as e:
        print(f"Error fetching data for {name}: {e}")

# Main function to process the Excel data
def process_coordinates_with_solar_api(df, api_key, output_dir, output_excel):
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Iterate through each row
    for _, row in df.iterrows():
        # Determine which coordinates to use based on the flag
        flag = row['Flag']
        if flag == 0:
            latitude = row['Geocoding Latitude']
            longitude = row['Geocoding Longitude']
        elif flag == 1:
            latitude = row['Actual Latitude']
            longitude = row['Actual Longitude']
        else:
            print(f"Invalid flag value for row {_}. Skipping...")
            continue

        # Get the location name
        name = row['Name']

        # Validate that coordinates and name are not empty
        if pd.notnull(latitude) and pd.notnull(longitude) and pd.notnull(name):
            fetch_data_from_solar_api(latitude, longitude, name, api_key, output_dir, df)
        else:
            print(f"Missing data for row {_}. Skipping...")

    # Save updated DataFrame to an Excel file
    df.to_excel(output_excel, index=False)
    print(f"Updated data saved to {output_excel}")

# Example usage with the provided API URL
api_key = "AIzaSyD4zRtuXdscMigE7PVeKXsdDCcY9UBjeJw"  # Replace with your actual API key
output_directory = "solar_api_responses"  # Directory to save the JSON responses
output_excel = "updated_solar_api_data.xlsx"  # File to save updated data

file_path = "SolarApiLocations_Deliverable1.xlsx"  # Update with your file path
df = pd.read_excel(file_path)
df_cleaned = df.rename(
    columns={
        'Unnamed: 0': 'Name',
        'Unnamed: 1': 'Address',
        'Geocoding API': 'Geocoding Latitude',
        'Unnamed: 3': 'Geocoding Longitude',
        'Actual': 'Actual Latitude',
        'Unnamed: 5': 'Actual Longitude',
        'Unnamed: 8': 'Flag',
        'Unnamed: 9': 'NumPanels',
        'Unnamed: 10': 'YearlyEnergy',
        'Unnamed: 11': 'SolarArea'
    }
)

# Drop the first row, which contains headers instead of data
df_cleaned = df_cleaned[1:]

# Convert necessary columns to numeric for consistency
df_cleaned['Geocoding Latitude'] = pd.to_numeric(df_cleaned['Geocoding Latitude'], errors='coerce')
df_cleaned['Geocoding Longitude'] = pd.to_numeric(df_cleaned['Geocoding Longitude'], errors='coerce')
df_cleaned['Actual Latitude'] = pd.to_numeric(df_cleaned['Actual Latitude'], errors='coerce')
df_cleaned['Actual Longitude'] = pd.to_numeric(df_cleaned['Actual Longitude'], errors='coerce')
df_cleaned['Flag'] = pd.to_numeric(df_cleaned['Flag'], errors='coerce')

# Process the data
process_coordinates_with_solar_api(df_cleaned, api_key, output_directory, output_excel)
