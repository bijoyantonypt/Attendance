import dash
from dash import dcc, html, dash_table
import plotly.express as px
import pandas as pd
import requests
from datetime import datetime, timedelta

app = dash.Dash(__name__)

# Fetch data from the Flask API
def fetch_data(endpoint):
    try:
        response = requests.get(f'http://127.0.0.1:5000/api/{endpoint}')
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"Failed to fetch data: {e}")
        return []

# Fetch attendance data
attendance_data = fetch_data('attendance')
attendance_df = pd.DataFrame(attendance_data)

# Fetch user data
users_data = fetch_data('users')
users_df = pd.DataFrame(users_data)

# Merge attendance data with user data
if not attendance_df.empty and not users_df.empty:
    attendance_df = pd.merge(attendance_df, users_df, on='user_id', how='left')

# Convert timestamp to datetime
if 'timestamp' in attendance_df.columns:
    attendance_df['timestamp'] = pd.to_datetime(attendance_df['timestamp'])

# Calculate daily attendance
daily_attendance = pd.DataFrame(columns=['date', 'count'])
if 'timestamp' in attendance_df.columns:
    daily_attendance = attendance_df.groupby(attendance_df['timestamp'].dt.date).size().reset_index(name='count')
    daily_attendance.columns = ['date', 'count']

# Create the dashboard layout
app.layout = html.Div([
    html.H1("Attendance Dashboard"),

    html.Div([
        html.H2("Daily Attendance"),
        dcc.Graph(
            id='daily-attendance',
            figure=px.line(daily_attendance, x='date', y='count', title='Daily Attendance')
        )
    ]),

    html.Div([
        html.H2("Attendance Records"),
        dash_table.DataTable(
            id='attendance-table',
            columns=[{"name": i, "id": i} for i in attendance_df.columns],
            data=attendance_df.to_dict('records'),
            page_size=10
        )
    ]),

    html.Div([
        html.H2("User List"),
        dash_table.DataTable(
            id='users-table',
            columns=[{"name": i, "id": i} for i in users_df.columns],
            data=users_df.to_dict('records'),
            page_size=10
        )
    ])
])

if __name__ == '__main__':
    app.run(debug=True)
