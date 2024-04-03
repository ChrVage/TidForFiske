import requests
from xml.etree import ElementTree
import pandas as pd
from datetime import datetime, timedelta
import pytz
import json

def get_high_tides(latitude, longitude, num_days):
    # Use pytz.utc for the UTC timezone
    utc = pytz.utc
    
    # Current time in UTC, used as the starting point (fromtime)
    fromtime = datetime.now(utc).strftime("%Y-%m-%dT%H:%M:%S+00:00")
    
    # Calculate totime by adding num_days to the current date
    totime = (datetime.now(utc) + timedelta(days=num_days)).strftime("%Y-%m-%dT%H:%M:%S+00:00")
    
    url = "https://api.sehavniva.no/tideapi.php"
    params = {
        "lat": latitude,
        "lon": longitude,
        "fromtime": fromtime,
        "totime": totime,
        "datatype": "tab",
        "refcode": "cd",
        "lang": "nn",
        "interval": 10,
        "dst": 1,
        "tide_request": "locationdata"
    }
    
    response = requests.get(url, params=params)
    if response.status_code == 200:
        root = ElementTree.fromstring(response.content)
        
        high_tides = []
        for waterlevel in root.findall(".//waterlevel[@flag='high']"):
            time_str = waterlevel.get("time")
            time_obj = datetime.fromisoformat(time_str)
            date = time_obj.strftime('%Y-%m-%d')
            time_end = time_obj.strftime('%H:%M:%S')
            time_start = (time_obj - timedelta(hours=1)).strftime('%H:%M:%S')
            height = waterlevel.get("value")
            
            high_tides.append({
                "time_start": f"{date} {time_start}",
                "time_end": f"{date} {time_end}",
                "data": height,
                'type': 'high_tide'
            })
        
        return high_tides
    else:
        return "Failed to fetch data"

def get_metno_locationforecast(latitude, longitude):
    url = "https://api.met.no/weatherapi/locationforecast/2.0/compact"

    headers = {
        'User-Agent': 'TidForFiske (Chr@Vage.com)'  
    }

    params = {
        "lat": latitude,
        "lon": longitude
    }
    
    response = requests.get(url, headers=headers, params=params)
    
    if response.status_code == 200:
        data = response.json()
        
        forecast = []
        for time in data["properties"]["timeseries"]:

            current_forecast = {
                "time": time["time"],
                "air_temperature": time["data"]["instant"]["details"]["air_temperature"],
                "wind_speed": time["data"]["instant"]["details"]["wind_speed"],
                "wind_from_direction": time["data"]["instant"]["details"]["wind_from_direction"]
            }
            
            # Safely get the precipitation amount if it exists in the 'next_1_hours' block
            next_1_hours_data = time["data"].get("next_1_hours", {}).get("details", {})
            if 'precipitation_amount' in next_1_hours_data:
                current_forecast["precipitation_amount"] = next_1_hours_data["precipitation_amount"]
            else:
                current_forecast["precipitation_amount"] = None

            forecast.append(current_forecast)
 
        return forecast
    else:
        return "Failed to fetch data"

def get_metno_oceanforecast(latitude, longitude):
    url = "https://api.met.no/weatherapi/oceanforecast/2.0/complete"
    
    headers = {
        'User-Agent': 'TidForFiske (Chr@Vage.com)'  
    }

    params = {
        "lat": latitude,
        "lon": longitude
    }
    
    response = requests.get(url, headers=headers, params=params)
    
    if response.status_code == 200:
        data = response.json()
        
        forecast = []
        for time in data["properties"]["timeseries"]:

            forecast.append({
                "time": time["time"],
                "sea_surface_wave_height": time["data"]["instant"]["details"]["sea_surface_wave_height"],
                "sea_surface_wave_from_direction": time["data"]["instant"]["details"]["sea_surface_wave_from_direction"],
                "sea_water_speed": time["data"]["instant"]["details"]["sea_water_speed"],
                "sea_water_temperature": time["data"]["instant"]["details"]["sea_water_temperature"],
                "sea_water_to_direction": time["data"]["instant"]["details"]["sea_water_to_direction"]
            })


        return forecast
    else:
        return "Failed to fetch data"

def get_metno_sunrise(latitude, longitude, NumDays=9):
    url = "https://api.met.no/weatherapi/sunrise/3.0/sun"
    
    headers = {
        'User-Agent': 'TidForFiske (Chr@Vage.com)' 
    }

    # Use pytz to specify CET timezone correctly
    cet = pytz.timezone('CET')

    results = []

    for i in range(NumDays):
        # Get the current date in UTC and convert it to CET
        date_utc = datetime.now(pytz.utc) + timedelta(days=i)
        date_cet = date_utc.astimezone(cet).strftime("%Y-%m-%d")

        params = {
            "lat": latitude,
            "lon": longitude,
            "date": date_cet
        }

        response = requests.get(url, headers=headers, params=params)
        if response.status_code == 200:
            data = response.json()
            
            date = params["date"]
            sunrise_time = data['properties']['sunrise']['time']
            sunrise_azimuth = data['properties']['sunrise']['azimuth']
            sunset_time = data['properties']['sunset']['time']
            sunset_azimuth = data['properties']['sunset']['azimuth']
            
            sunrise_time_cet = datetime.fromisoformat(sunrise_time).astimezone(cet).strftime("%H:%M:%S")
            sunset_time_cet = datetime.fromisoformat(sunset_time).astimezone(cet).strftime("%H:%M:%S")

            results.append({
                'date': date,
                'sunrise_time': sunrise_time_cet,
                'sunrise_azimuth': sunrise_azimuth,
                'sunset_time': sunset_time_cet,
                'sunset_azimuth': sunset_azimuth,
            })

        else:
            print("Failed to fetch data")

    return results

def create_excel_from_data(data, file_name='output.xlsx'):
    # Convert the list of dictionaries to a pandas DataFrame
    df = pd.DataFrame(data)
    
    # Write the DataFrame to an Excel file, default sheet name will be used ("Sheet1")
    df.to_excel(file_name, index=False)
    print(f"Excel file '{file_name}' has been created.")

with open('TidForFiske_config.json', 'r', encoding='utf-8') as config_file:
    config = json.load(config_file)

# Accessing the configuration variables
EarlyTime = config.get("EarlyTime", "06:00")
LateTime = config.get("LateTime", "23:00")
latitude = config.get("latitude", 60.0)
longitude = config.get("longitude", 5.0)
CreateCalendar = config.get("CreateCalendar", True)
NumDays = config.get("NumDays", 9)

latitude, longitude = 59.94297433675198, 5.071140027978568 # Punkt 66

high_tides = get_high_tides(latitude, longitude, NumDays)
locationforecast = get_metno_locationforecast(latitude, longitude)
oceanforecast = get_metno_oceanforecast(latitude, longitude)
sun_times = get_metno_sunrise(latitude, longitude, NumDays)

sunrise = [{
    'time_start': f"{item['date']} {item['sunrise_time']}",
    'time_end': datetime.strptime(item['date'] + ' ' + item['sunrise_time'], "%Y-%m-%d %H:%M:%S") + timedelta(hours=1),
    'data': item['sunrise_azimuth'],
    'type': 'sunrise'
} for item in sun_times]

# Convert 'time_end' back to string if necessary
for item in sunrise:
    item['time_end'] = item['time_end'].strftime("%Y-%m-%d %H:%M:%S")

sunset = [{
    'time_start': datetime.strptime(item['date'] + ' ' + item['sunset_time'], "%Y-%m-%d %H:%M:%S") - timedelta(hours=1),
    'time_end': f"{item['date']} {item['sunset_time']}",
    'data': item['sunset_azimuth'],
    'type': 'sunset'
} for item in sun_times]

# Convert 'time_start' back to string if necessary
for item in sunset:
    item['time_start'] = item['time_start'].strftime("%Y-%m-%d %H:%M:%S")

fisketider = high_tides + sunrise + sunset

# Assuming fisketider is already filled with your data
df = pd.DataFrame(fisketider)

# Convert EarlyTime and LateTime to timedeltas for comparison
early_time_obj = datetime.strptime(EarlyTime, '%H:%M').time()
late_time_obj = datetime.strptime(LateTime, '%H:%M').time()

# Ensure 'time_start' and 'time_end' are in datetime.time format for comparison
df['time_start'] = pd.to_datetime(df['time_start'], format='%Y-%m-%d %H:%M:%S')
df['time_end'] = pd.to_datetime(df['time_end'], format='%Y-%m-%d %H:%M:%S')

# Filter out rows where 'time_end' is earlier than 'EarlyTime' or 'time_start' is later than 'LateTime'
filtered_df = df[(df['time_end'].dt.time >= early_time_obj) & (df['time_start'].dt.time <= late_time_obj)]

# Sort by 'date', then by 'time_start'
sorted_filtered_df = filtered_df.sort_values(by=['time_start'])

# Convert back to your desired format if necessary, for example, to a list of dictionaries
filtered_sorted_data = sorted_filtered_df.to_dict('records')

# Finally, create an Excel file with the filtered and sorted data
create_excel_from_data(filtered_sorted_data, "zz_fisketider_filtered_sorted.xlsx")
# create_excel_from_data(oceanforecast, "zz_oceanforecast.xlsx")
create_excel_from_data(locationforecast, "zz_locationforecast.xlsx")
