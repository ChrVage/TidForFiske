# Tide Forecast Processing Tool

This Python script provides functionalities to fetch and process tide, weather, and sunrise/sunset forecasts for specified geographical coordinates. It utilizes APIs from `sehavniva.no`, `api.met.no`, and processes data to generate a comprehensive outlook, focusing on high tide times, weather forecasts, and ocean forecasts.

## Features

- Fetch high tide times from the `sehavniva.no` tide API.
- Obtain weather forecasts including air temperature, wind speed, and direction from the `api.met.no` locationforecast API.
- Retrieve ocean forecasts detailing sea surface wave height, sea water speed, and temperature from the `api.met.no` oceanforecast API.
- Gather sunrise and sunset times, along with azimuth details from the `api.met.no` sunrise API.
- Process and compile these forecasts into a structured format.
- Create an Excel file summarizing the fetched and processed data, tailored for further analysis or record-keeping.

## Usage

Ensure you have Python 3.x installed along with the required libraries: `requests`, `pandas`, `pytz`, and `openpyxl` (for Excel file creation).

1. Configure the desired geographical coordinates (latitude, longitude) and the number of forecast days in `TidForFiske_config.json`.
2. Run the script:
   ```bash
   python <script_name>.py
