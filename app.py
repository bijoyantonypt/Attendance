from flask import Flask, jsonify
from flask_cors import CORS
from zk import ZK
import pandas as pd
from datetime import datetime

app = Flask(__name__)
CORS(app)

DEVICE_IP = '192.168.29.24'
DEVICE_PORT = 4370
TIMEOUT = 10

def connect_device():
    try:
        zk = ZK(DEVICE_IP, port=DEVICE_PORT, timeout=TIMEOUT, password=0)
        conn = zk.connect()
        return conn
    except Exception as e:
        print(f"Connection failed: {e}")
        return None

def fetch_attendance(conn):
    try:
        attendance = conn.get_attendance()
        logs = []
        for att in attendance:
            logs.append({
                "user_id": att.user_id,
                "timestamp": att.timestamp.isoformat(),
                "status": att.status,
                "punch": att.punch
            })
        return logs
    except Exception as e:
        print(f"Failed to fetch attendance: {e}")
        return []

def fetch_users(conn):
    try:
        users = conn.get_users()
        user_list = []
        for user in users:
            user_list.append({
                "user_id": user.uid,
                "name": user.name,
                "card_number": getattr(user, 'card_number', None),
                "pin": getattr(user, 'pin', None),
                "role": getattr(user, 'role', None),
            })
        return user_list
    except Exception as e:
        print(f"Failed to fetch users: {e}")
        return []

@app.route('/api/attendance', methods=['GET'])
def get_attendance():
    conn = connect_device()
    if not conn:
        return jsonify({"error": "Failed to connect to device"}), 500

    try:
        attendance = fetch_attendance(conn)
        return jsonify(attendance)
    finally:
        conn.disconnect()

@app.route('/api/users', methods=['GET'])
def get_users():
    conn = connect_device()
    if not conn:
        return jsonify({"error": "Failed to connect to device"}), 500

    try:
        users = fetch_users(conn)
        return jsonify(users)
    finally:
        conn.disconnect()

if __name__ == '__main__':
    app.run(debug=True)
