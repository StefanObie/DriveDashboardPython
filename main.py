import pandas as pd
import urllib.parse
import requests
import config
import openpyxl
import glob
import os

def ceil_time_to_minute(time):
    return time + pd.Timedelta(minutes=1) - pd.Timedelta(seconds=time.second, microseconds=time.microsecond)

def load_file():
    # Read the latest CSV file from the MovementReports directory
    movement_reports_dir = 'MovementReports'
    csv_files = glob.glob(os.path.join(movement_reports_dir, '*.csv'))
    if not csv_files:
        raise FileNotFoundError(f"No CSV files found in {movement_reports_dir}")
    file_path = max(csv_files, key=lambda f: os.path.getctime(os.path.abspath(f)))

    # Specific file paths for testing
    # file_path = 'MovementReports\\StefanMovementReport17Apr25.csv'
    # file_path = 'MovementReports\\StefanMovementReportApr2025.csv' # 1-3 April missing (max data retention 38 days)
    # file_path = 'MovementReports\\StefanMovementReport29May2025.csv' # 29-31 May missing
    # file_path = 'MovementReports\\StefanMovementReport30May2025.csv' # 31 May missing
    # file_path = 'MovementReports\\StefanMovementReportMay2025.csv'

    print(f"Loading file: {file_path}\n")
    df = pd.read_csv(file_path, encoding='latin1')
    return df

def preprocessing(df):
    # Remove empty columns
    columns_to_remove = ["DriverID","SkillSet","MsgTypeId","LocationTolerance","STATUS1"]
    df = df.drop(columns=columns_to_remove, errors='ignore')

    # Split the 'Location' column into 'Longitude' and 'Latitude'
    df[['Longitude', 'Latitude']] = df['Location'].str.extract(r'Long\s*:\s*([\d\.\-]+)\.\s*Lat\s*:\s*([\d\.\-]+)')
    df = df.drop(columns=['Location'], errors='ignore')

    # Convert datetime columns to datetime objects 4/1/2025 4:00:26 AM
    df['DateTime'] = pd.to_datetime(df['Report Group Date'], format='%m/%d/%Y %I:%M:%S %p', errors='coerce')

    # Trip Numbering
    df['TripNumber'] = (df['VehicleStatus'] == 'Start up').cumsum()

    return df

def no_drive_days(df, full_month=False):
    df_no_drive = df.copy()
    df_no_drive['Date'] = df_no_drive['DateTime'].dt.date

    if (full_month): # Full Month
        first_date = df_no_drive['Date'].min().replace(day=1)
        days_in_month = pd.Period(first_date, freq='M').days_in_month
        last_date = first_date + pd.Timedelta(days=days_in_month - 1)
    else: # Month-to-Date
        # first_date = df_no_drive['Date'].min() # Report Start Date
        first_date = df_no_drive['Date'].min().replace(day=1) # Month-to-Date
        last_date = df_no_drive['Date'].max()
    print(f"Report Dates: {first_date} - {last_date}")

    df_no_drive = df_no_drive[df["VehicleStatus"] == 'Start up']
    df_no_drive = df_no_drive.drop_duplicates(subset=['Date'], keep='first')

    num_no_drive_days = (last_date - first_date).days + 1 - len(df_no_drive)
    print(f"No-Drive Days: {num_no_drive_days}\n")
    return num_no_drive_days

def driving_violations(df, violation='Harsh Braking'):
    df_braking = df[df['VehicleStatus'] == violation]
    df_braking = df_braking.sort_values(by='DateTime') 
    df_braking = df_braking[~(df_braking['DateTime'].diff().dt.total_seconds().abs() <= 120)] # Remove false alarms for Harsh Braking
    print(f"{violation}: {len(df_braking)} ({len(df_braking) *8} Points)\n")

    return len(df_braking) *8

def get_speed_limit(lat, lon):
    # Get the speed limit for a given latitude and longitude using HERE API.
    base_url = "https://revgeocode.search.hereapi.com/v1/revgeocode"
    params = {
        "at": f"{lat},{lon},50",
        "maxResults": "1",
        "apiKey": config.HERE_API_KEY,
        "showNavAttributes": "speedLimits",
        "types": "street"
    }
    
    url = f"{base_url}?{urllib.parse.urlencode(params)}"
    print(f"Calling HERE API...")
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if 'items' in data and len(data['items']) > 0:
            speed_limits = data['items'][0].get('navigationAttributes', {}).get('speedLimits', [])
            if speed_limits and 'maxSpeed' in speed_limits[0]:
                return speed_limits[0]['maxSpeed']
    return None

def speed_violation(df, call_here_api_for_speedlimit=False, default_speed_limit=60):
    df_speed = df[(df['VehicleStatus'] == 'Speed Violation') & 
                (df['MOBILESPEED'] >= default_speed_limit+10)] # Everything under 70 km/h is not a violation
    
    print(f"Speed Violations") # Heading
    
    if len(df_speed) > 10 and call_here_api_for_speedlimit: # Do not call HERE API if there are too many speed violations
        print(df_speed[['DateTime', 'MOBILESPEED', 'Latitude', 'Longitude']])
        confirm = input(f"There are {len(df_speed)} speed violation(s). Type y (yes) to continue with HERE API calls: ")
        if confirm.strip().lower() != 'y':
            print(f"Using default speed limit of {default_speed_limit} km/h.")
            call_here_api_for_speedlimit = False

    speed_penalty_total = 0
    for _, row in df_speed.iterrows():
        speed = row['MOBILESPEED']

        if call_here_api_for_speedlimit:
            speedlimit = get_speed_limit(row['Latitude'], row['Longitude'])
        else: # Use a fixed speed limit
            speedlimit = default_speed_limit # Default speed limit

        if speedlimit is None:
            print(f"[ERROR] No speed limit found for coordinates: {row['Latitude']}, {row['Longitude']}")
            continue # Skip to the next iteration

        if speed - speedlimit >= 10:
            if speed - speedlimit <= 15: # 10..15
                penalty = 3
            elif speed - speedlimit <= 25: # 16..25
                penalty = 8
            elif speed - speedlimit > 25: # 26..
                penalty = 15

            speed_penalty_total += penalty
            print(f"\tSpeed Violation: {speed} km/h, Speed Limit: {speedlimit} km/h, Penalty: {penalty} points")

    print(f"Speed Violation Total: {speed_penalty_total} Points\n")
    return speed_penalty_total

def night_time_driving(df):
    # Night Time Driving
    df['Hour'] = df['DateTime'].dt.hour
    df_night_driving = df[((df['Hour'] >= 23) | (df['Hour'] <= 4.5)) & 
                        (df['VehicleStatus'] != 'Health Check; (Ignition off)')].copy()
        
    print(f"Night Time Driving")
    night_penalty_total = 0

    # For each TripNumber in the night driving dataframe
    trip_nums = df_night_driving['TripNumber'].unique()
    for trip in trip_nums:
        trip_df = df[df['TripNumber'] == trip]

        # Find difference between "Start up" and "Ignition off"
        if len(trip_df) > 1:
            start_time = trip_df[trip_df['VehicleStatus'] == 'Start up']['DateTime'].min()
            end_time = trip_df[trip_df['VehicleStatus'] == 'Ignition off']['DateTime'].max()
            duration = end_time - start_time

            # If time is between 23h and 4h30, add a penalty for each minute.
            t = start_time + pd.Timedelta(minutes=1, seconds=-1)
            night_penalty = 0
            while t <= end_time:
                if t.hour in [23, 4]:
                    night_penalty += 2
                elif t.hour in [0, 3]:
                    night_penalty += 4
                elif t.hour in [1, 2]:
                    night_penalty += 6
                # print(t, night_penalty)
                t += pd.Timedelta(minutes=1)
            # print(night_penalty)

            night_penalty_total += night_penalty 
            print(f"\tTrip Number: {trip}, Duration: {duration.components.hours}h:{duration.components.minutes:02d}m, Night Penalty: {night_penalty} points")

    print(f"Night Time Driving Total: {night_penalty_total}\n")
    return night_penalty_total

def distance(df):
    d = df['MOBILEODO'].iloc[-1]
    print(f"Month-to-Date Distance: {d:.0f} km")
    return d

def sheetname_lookup(df, dev=False):
    return 'DEV' if dev else config.SHEETNAME_LOOKUP.get(df['VehicleReg'].iloc[0], 'DEV')

def write_to_excel(drive_pen=0, night_pen=0, no_drive=0, dist=0, sheetname='DEV'):
    output_file = 'DriveTemplateDevelopmentPython.xlsm' if sheetname == 'DEV' else 'DriveSummaryPython.xlsx'

    try:
        wb = openpyxl.load_workbook(output_file)
    except FileNotFoundError:
        print(f"File {output_file} not found.")
        return

    if sheetname in wb.sheetnames:
        ws = wb[sheetname]
    else:
        print(f"Sheet {sheetname} not found in the workbook.")
        return

    ws['J6'] = drive_pen
    ws['J7'] = night_pen
    ws['J8'] = no_drive
    ws['J9'] = dist
    ws['J14'] = pd.Timestamp.now().strftime('%Y/%m/%d')

    wb.save(output_file)
    wb.close()

    print(f"Data written to Excel successfully.")

def main():
    # Config
    full_month = True
    call_here_api_for_speedlimit = True
    dev = False
    save_to_excel = True

    # Load and preprocess data
    df = load_file()
    df = preprocessing(df)

    no_drive = no_drive_days(df, full_month)

    drive_pen = 0
    drive_pen += driving_violations(df, violation='Harsh Braking')
    drive_pen += driving_violations(df, violation='Harsh Acceleration')
    drive_pen += driving_violations(df, violation='Harsh Cornering')
    drive_pen += speed_violation(df, call_here_api_for_speedlimit)
    
    night_pen = night_time_driving(df)
    dist = distance(df)

    sheetname = sheetname_lookup(df, dev)
    if save_to_excel:
        write_to_excel(drive_pen, night_pen, no_drive, dist, sheetname)

if main() == '__main__':
    main()