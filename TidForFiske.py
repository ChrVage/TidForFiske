import requests
import pandas as pd
from datetime import datetime, timedelta
import pytz
import json

def fetch_data(url, params=None, headers=None):
    """Unified data fetching function to minimize repetitive code."""
    response = requests.get(url, params=params, headers=headers)
    return response.json() if response.status_code == 200 else None

def get_time_range(num_days, timezone='UTC'):
    """Generate from and to time strings."""
    utc_now = datetime.now(pytz.utc)
    return utc_now.strftime("%Y-%m-%dT%H:%M:%S+00:00"), (utc_now + timedelta(days=num_days)).strftime("%Y-%m-%dT%H:%M:%S+00:00")

def process_high_tides(data, timezone):
    """Process high tides data."""
    high_tides = []
    for waterlevel in data.findall(".//waterlevel[@flag='high']"):
        time_obj = datetime.fromisoformat(waterlevel.get("time").replace('Z', '+00:00')).astimezone(timezone)
        high_tides.append({
            "time_start": time_obj - timedelta(hours=1),
            "time_end": time_obj,
            "height": waterlevel.get("value"),
            "type": "high_tide"
        })
    return high_tides

def process_forecasts(data, forecast_type):
    """Process weather and ocean forecasts."""
    forecasts = []
    for entry in data["properties"]["timeseries"]:
        forecast = {"time": entry["time"]}
        if forecast_type == 'weather':
            forecast.update(entry["data"]["instant"]["details"])
        elif forecast_type == 'ocean':
            forecast.update(entry["data"]["instant"]["details"])
        forecasts.append(forecast)
    return forecasts

def get_combined_data(latitude, longitude, num_days):
    """Fetch and combine data from different sources."""
    timezone = pytz.timezone('CET')
    fromtime, totime = get_time_range(num_days)
    
    # URLs and parameters
    tide_api_url = "https://api.sehavniva.no/tideapi.php"
    weather_api_url = "https://api.met.no/weatherapi/locationforecast/2.0/compact"
    ocean_api_url = "https://api.met.no/weatherapi/oceanforecast/2.0/complete"
    sunrise_api_url = "https://api.met.no/weatherapi/sunrise/2.0/.json"
    
    common_params = {"lat": latitude, "lon": longitude, "fromtime": fromtime, "totime": totime}
    headers = {'User-Agent': 'TidForFiske (Chr@Vage.com)'}
    
    # Fetch data
    tide_data = fetch_data(tide_api_url, params=common_params)
    weather_data = fetch_data(weather_api_url, headers=headers, params={"lat": latitude, "lon": longitude})
    ocean_data = fetch_data(ocean_api_url, headers=headers, params={"lat": latitude, "lon": longitude})
    
    # Process data
    high_tides = process_high_tides(tide_data, timezone) if tide_data else []
    weather_forecast = process_forecasts(weather_data, 'weather') if weather_data else []
    ocean_forecast = process_forecasts(ocean_data, 'ocean') if ocean_data else []
    
    return high_tides, weather_forecast, ocean_forecast

def create_excel(data, file_name='output.xlsx'):
    """Generate Excel file from data."""
    df = pd.DataFrame(data)
    df.to_excel(file_name, index=False)
    print(f"Excel file '{file_name}' has been created.")

# Configuration and data fetching
config = json.load(open('TidForFiske_config.json', 'r', encoding='utf-8'))
latitude, longitude, num_days = config.get("latitude", 59.9), config.get("longitude", 5.0), config.get("NumDays", 30)

high_tides, weather_forecast, ocean_forecast = get_combined_data(latitude, longitude, num_days)

# Further processing and Excel creation can follow based on the obtained data
create_excel(high_tides + weather_forecast + ocean_forecast, "combined_data.xlsx")
