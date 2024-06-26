# The intention of this code is to fetch data about the time for high tides, sunrise/sunset times and times with low current and low wind speed.
# The times will collected with starttime. Length will be added from the TidForFiske_config.json. According to mode, a flag will be added to the data.
# To establish if the time is good for fishing, the following parameters will be used:
#   Wave height, wind speed, wind direction, temperature, cloud cover, precipitation and currents.
# Also the forecasts are more reliable the closer to the current time they are.
# Parameters to use in the code is collected from TidForFiske_config.json. The parameters are 
# latitude, longitude to establish which location the api's should get data about
#   time_start, time_end to set limits for how early and late the calendar times should be
#   duration_prep, duration_fish, duration_home to define how long time the different parts of the fishing trip will take
#   NumDays to define how many days the forecasts should be fetched for
#   Mode to define if the code should be used for fishing or finding quiet times (current and wind speed)
# The code will fetch data from the following api's:
#   Collecting times:
#   https://api.sehavniva.no/tideapi.php to get data about high tides
#   https://api.met.no/weatherapi/sunrise/2.0/.json to get sunrise and sunset times
#   Determining if the time is suitable for fishing:
#   https://api.met.no/weatherapi/locationforecast/2.0/compact to get weather forecast
#   https://api.met.no/weatherapi/oceanforecast/2.0/complete to get ocean forecast that will be used to determine if the time is suitable for fishing.
#
# The ouput of this program is an .ics file that can be imported into a calendar program. The .ics file will contain the times for high tides, sunrise and sunset times and times with low current and low wind speed.
# The subject of the event will contain a character (percentage) that indicates if the time is suitable for fishing or not. If the character is low enough the event will be removed from the file.
# The description of the event will contain the weather forecast and ocean forecast for the time. 
# To compile the character, the following parameters will be used:
#   Wave height: < 0.5m = 100%, > 2m = 0%
#   Wind speed: < 3m/s = 100%, > 13m/s = 0%
#   Wind direction: 30 degrees = 100%, 210 degrees = 80%
#   Temperature: < 10 degrees = 100%, > 15 degrees = 80%
#   Cloud cover: Not used
#   Precipitation: < 0.1mm = 100%, > 5mm = 50%
#   Currents: < 0.5m/s = 100%, > 1m/s = 50%
# The character will be calculated as the average of the parameters. If the character is below a certain threshold the event will be removed from the file.
# Another output will be an Excel file with all the times, sorted by time. The Excel file will contain the same information as the .ics file.
# There should be defined a few functions:
#   fetch_data(url, params=None, headers=None) to fetch data from the api's. The function will return the data in a flattened format. With all times converted to CET.
#   get_time_range(num_days, timezone='UTC') to generate the time range that the data should be fetched for. 
#   set_fish_times(sun_data, high_tides, duration_prep, duration_fish, duration_home) to set the times for the fishing trip. 
#       The function will return a list of dictionaries with the time for the high tide, the time for the sunrise and sunset with indication if it is sunset or sunrise or if it is high tide.
#   set_silent_times(weather_forecast, ocean_forecast, duration_prep, duration_fish, duration_home) to set the times for the silent trip.
#   set_forecasts(timelist, weather_forecast, ocean_forecast) to set the forecasts for the times. The function will return a list of dictionaries with the time for the forecast and the forecast data. 
#   create_ics(data, file_name='output.ics') to create an .ics file from the data. The .ics file will contain the times for high tides, sunrise and sunset times and times with low current and low wind speed.
# The urls should be defined and all function calls should be written in main().


# The following code is a template for the functions that should be implemented in the code. The code is not complete and should be completed with the necessary code to fetch data from the api's and process the data.

# import requests
# import pandas as pd
# from datetime import datetime, timedelta
# import pytz
# import json
# from xml.etree import ElementTree as ET


# # This function is used to fetch data from the api's. The function will return the data in the format that the api returns.
# def fetch_data(url, params=None, headers=None):
#     """Unified data fetching function to handle both XML and JSON responses."""
#     response = requests.get(url, params=params, headers=headers)

#     if response.status_code == 200:
#         content_type = response.headers.get('Content-Type', '')
#         if 'application/json' in content_type:
#             return response.json()
#         elif 'text/xml' in content_type or 'application/xml' in content_type:
#             # Properly parse XML data
#             return ET.fromstring(response.content)
#         else:
#             # This path now explicitly tries to parse XML as a fallback
#             try:
#                 return ET.fromstring(response.content)
#             except ET.ParseError:
#                 # If parsing fails, return None or handle as needed
#                 return None
#     return None

# # This function is used to generate the time range that the data should be fetched for. The function will return the time in the format that the api's expect.
# def get_time_range(num_days, timezone='UTC'):
#     """Generate from and to time strings."""
#     utc_now = datetime.now(pytz.utc)
#     return utc_now.strftime("%Y-%m-%dT%H:%M:%S+00:00"), (utc_now + timedelta(days=num_days)).strftime("%Y-%m-%dT%H:%M:%S+00:00")

# # This function is used to process the high tide data. The function will return a list of dictionaries with the time for the high tide, the height of the tide and the type of the tide.
# def process_high_tides(data, timezone):
#     """Process high tides data."""
#     high_tides = []
#     for waterlevel in data.findall(".//waterlevel[@flag='high']"):
#         time_obj = datetime.fromisoformat(waterlevel.get("time").replace('Z', '+00:00')).astimezone(timezone)
#         high_tides.append({
#             "time_start": time_obj - timedelta(hours=1),
#             "time_end": time_obj,
#             "height": waterlevel.get("value"),
#             "type": "high_tide"
#         })
#     return high_tides

# # This function is used to process the weather and ocean forecast data. The function will return a list of dictionaries with the time for the forecast and the forecast data.
# def process_forecasts(data, forecast_type):
#     """Process weather and ocean forecasts."""
#     forecasts = []
#     for entry in data["properties"]["timeseries"]:
#         forecast = {"time": entry["time"]}
#         if forecast_type == 'weather':
#             forecast.update(entry["data"]["instant"]["details"])
#         elif forecast_type == 'ocean':
#             forecast.update(entry["data"]["instant"]["details"])
#         forecasts.append(forecast)
#     return forecasts


# def get_data(latitude, longitude, num_days):
#     """Fetch and combine data from different sources."""
#     timezone = pytz.timezone('CET')
#     fromtime, totime = get_time_range(num_days)
    
#     # URLs and parameters
#     tide_api_url = "https://api.sehavniva.no/tideapi.php"
#     weather_api_url = "https://api.met.no/weatherapi/locationforecast/2.0/compact"
#     ocean_api_url = "https://api.met.no/weatherapi/oceanforecast/2.0/complete"
#     sunrise_api_url = "https://api.met.no/weatherapi/sunrise/2.0/.json"
    
#     common_params = {"lat": latitude, "lon": longitude, "fromtime": fromtime, "totime": totime}
#     headers = {'User-Agent': 'TidForFiske (Chr@Vage.com)'}
    
#     # Fetch data
#     tide_data = fetch_data(tide_api_url, params=common_params)
#     weather_data = fetch_data(weather_api_url, headers=headers, params={"lat": latitude, "lon": longitude})
#     ocean_data = fetch_data(ocean_api_url, headers=headers, params={"lat": latitude, "lon": longitude})
#     sun_data = fetch_data(sunrise_api_url, headers=headers, params={"lat": latitude, "lon": longitude})
    
#     # Process data
#     high_tides = process_high_tides(tide_data, timezone) if tide_data else []
#     weather_forecast = process_forecasts(weather_data, 'weather') if weather_data else []
#     ocean_forecast = process_forecasts(ocean_data, 'ocean') if ocean_data else []
#     sun_data = process_forecasts(sun_data, 'sunrise') if sun_data else []
    
#     return high_tides, weather_forecast, ocean_forecast, sun_data

# def create_excel(data, file_name='output.xlsx'):
#     """Generate Excel file from data."""
#     df = pd.DataFrame(data)
#     df.to_excel(file_name, index=False)
#     print(f"Excel file '{file_name}' has been created.")

# # Configuration and data fetching
# config = json.load(open('TidForFiske_config.json', 'r', encoding='utf-8'))

# latitude = config.get("latitude", 59.9)
# longitude = config.get("longitude", 5.0)
# time_start = config.get("time_start", 6)
# time_end = config.get("time_end", 18)
# duration_prep = config.get("duration_prep", 1)
# duration_fish = config.get("duration_fish", 4)
# duration_home = config.get("duration_home", 1)
# num_days = config.get("NumDays", 30)
# mode = config.get("mode", "fishing")

# high_tides, weather_forecast, ocean_forecast, sun_data = get_data(latitude, longitude, num_days)

# # Further processing and Excel creation can follow based on the obtained data
# create_excel(high_tides, 'high_tides.xlsx')
# create_excel(sun_data, 'sun_data.xlsx')
# create_excel(weather_forecast, 'weather_forecast.xlsx')
# create_excel(ocean_forecast, 'ocean_forecast.xlsx')

# Code start:

import requests
import pandas as pd
from datetime import datetime, timedelta
import pytz
import json
from xml.etree import ElementTree as ET
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as ticker
from ics import Calendar, Event
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

def fetch_data(url, params=None, headers=None):
    """Fetch data from the API."""
    response = requests.get(url, params=params, headers=headers)

    if response.status_code == 200:
        content_type = response.headers.get('Content-Type', '')
        if 'application/json' in content_type:
            return response.json()
        elif 'text/xml' in content_type or 'application/xml' in content_type:
            return ET.fromstring(response.content)
        else:
            try:
                return ET.fromstring(response.content)
            except ET.ParseError:
                return None
    return None


def set_fish_times(sun_data, high_tides, duration_prep, duration_fish, duration_home):
    """Set the times for the fishing trip."""
    fish_times = []
    for tide in high_tides:
        if tide["type"] == "high_tide":
            fish_start = tide["time_start"] - timedelta(hours=duration_prep)
            fish_end = tide["time_end"] + timedelta(hours=duration_fish + duration_home)
            fish_times.append({"time": fish_start, "type": "prep"})
            fish_times.append({"time": tide["time_start"], "type": "high_tide"})
            fish_times.append({"time": tide["time_end"], "type": "high_tide"})
            fish_times.append({"time": fish_end, "type": "home"})
    return fish_times

def set_silent_times(weather_forecast, ocean_forecast, duration_prep, duration_fish, duration_home):
    """Set the times for the silent trip."""
    silent_times = []
    for forecast in weather_forecast:
        if forecast["temperature"] < 10 and forecast["cloud_cover"] < 0.5 and forecast["precipitation"] < 0.1:
            silent_start = forecast["time"] - timedelta(hours=duration_prep)
            silent_end = forecast["time"] + timedelta(hours=duration_fish + duration_home)
            silent_times.append({"time": silent_start, "type": "prep"})
            silent_times.append({"time": forecast["time"], "type": "silent"})
            silent_times.append({"time": silent_end, "type": "home"})
    return silent_times

def set_forecasts(timelist, weather_forecast, ocean_forecast):
    """Set the forecasts for the times."""
    forecasts = []
    for time in timelist:
        for forecast in weather_forecast:
            if forecast["time"] == time["time"]:
                forecasts.append({"time": time["time"], "type": time["type"], "forecast": forecast})
                break
        for forecast in ocean_forecast:
            if forecast["time"] == time["time"]:
                forecasts.append({"time": time["time"], "type": time["type"], "forecast": forecast})
                break
    return forecasts

def create_ics(data, file_name='output.ics'):
    """Create an .ics file from the data."""
    c = Calendar()
    for item in data:
        event = Event()
        event.name = item["type"]
        event.begin = item["time"]
        event.description = str(item["forecast"])
        c.events.add(event)
    with open(file_name, 'w') as f:
        f.write(str(c))
    print(f"ICS file '{file_name}' has been created.")

def main():
    # Configuration and data fetching
    config = json.load(open('TidForFiske_config.json', 'r', encoding='utf-8'))

    latitude = config.get("latitude", 59.9)
    longitude = config.get("longitude", 5.0)
    time_start = config.get("time_start", 6)
    time_end = config.get("time_end", 18)
    duration_prep = config.get("duration_prep", 1)
    duration_fish = config.get("duration_fish", 4)
    duration_home = config.get("duration_home", 1)
    num_days = config.get("NumDays", 30)
    mode = config.get("mode", "fishing")

    # URLs and parameters
    tide_api_url = "https://api.sehavniva.no/tideapi.php"
    weather_api_url = "https://api.met.no/weatherapi/locationforecast/2.0/compact"
    ocean_api_url = "https://api.met.no/weatherapi/oceanforecast/2.0/complete"
    sunrise_api_url = "https://api.met.no/weatherapi/sunrise/2.0/.json"




    high_tides, weather_forecast, ocean_forecast, sun_data = get_data(latitude, longitude, num_days)

    if mode == "fishing":
        fish_times = set_fish_times(sun_data, high_tides, duration_prep, duration_fish, duration_home)
        forecasts = set_forecasts(fish_times, weather_forecast, ocean_forecast)
        create_ics(forecasts, 'fishing.ics')
    elif mode == "silent":
        silent_times = set_silent_times(weather_forecast, ocean_forecast, duration_prep, duration_fish, duration_home)
        forecasts = set_forecasts(silent_times, weather_forecast, ocean_forecast)
        create_ics(forecasts, 'silent.ics')
    else:
        print("Invalid mode. Please choose 'fishing' or 'silent'.")

    create_excel(high_tides, 'high_tides.xlsx')
    create_excel(sun_data, 'sun_data.xlsx')
    create_excel(weather_forecast, 'weather_forecast.xlsx')
    create_excel(ocean_forecast, 'ocean_forecast.xlsx')

if __name__ == "__main__":
    main()