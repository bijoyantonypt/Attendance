from apscheduler.schedulers.background import BackgroundScheduler
import requests
import pandas as pd
from datetime import datetime

def fetch_and_store_data():
    try:
        # Fetch attendance data
        attendance_response = requests.get('http://127.0.0.1:5000/api/attendance')
        attendance_response.raise_for_status()
        attendance_data = attendance_response.json()

        # Fetch user data
        users_response = requests.get('http://127.0.0.1:5000/api/users')
        users_response.raise_for_status()
        users_data = users_response.json()

        # Save data to CSV files
        attendance_df = pd.DataFrame(attendance_data)
        users_df = pd.DataFrame(users_data)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        attendance_df.to_csv(f'attendance_{timestamp}.csv', index=False)
        users_df.to_csv(f'users_{timestamp}.csv', index=False)

        print(f"Data fetched and saved at {timestamp}")
    except Exception as e:
        print(f"Failed to fetch and store data: {e}")

# Schedule the job to run every hour
scheduler = BackgroundScheduler()
scheduler.add_job(fetch_and_store_data, 'interval', hours=1)
scheduler.start()

try:
    # Keep the scheduler running
    while True:
        pass
except (KeyboardInterrupt, SystemExit):
    scheduler.shutdown()
