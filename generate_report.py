"""
ESSL X990 Local Attendance Report Generator
Connects to the device over LAN and exports attendance to Excel.
"""

from zk import ZK, const
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import os
import sys

# ----------- CONFIGURATION -----------
DEVICE_IP = "192.168.1.41"   # <-- CHANGE THIS to your X990's actual IP
DEVICE_PORT = 4370
TIMEOUT = 10
OUTPUT_DIR = "reports"
# --------------------------------------


def connect_device():
    print(f"Connecting to ESSL device at {DEVICE_IP}:{DEVICE_PORT} ...")
    zk = ZK(DEVICE_IP, port=DEVICE_PORT, timeout=TIMEOUT,
            password=0, force_udp=False, ommit_ping=False)
    try:
        conn = zk.connect()
        print("✅ Connected successfully.")
        return conn
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        print("Check: device IP, LAN cable, firewall, and that you're on the same network.")
        sys.exit(1)


def fetch_data(conn):
    print("Fetching users...")
    users = conn.get_users()
    user_map = {u.user_id: u.name for u in users}

    print("Fetching attendance logs...")
    attendance = conn.get_attendance()
    print(f"✅ Retrieved {len(attendance)} attendance records.")
    return users, attendance, user_map


def status_label(status):
    mapping = {0: "Check-In", 1: "Check-Out", 2: "Break-Out", 3: "Break-In",
               4: "OT-In", 5: "OT-Out"}
    return mapping.get(status, f"Unknown({status})")


def export_to_excel(attendance, user_map):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    timestamp_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = os.path.join(OUTPUT_DIR, f"attendance_report_{timestamp_str}.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"

    headers = ["User ID", "Name", "Date", "Time", "Status", "Verified"]
    ws.append(headers)

    # Style header row
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # Sort by timestamp
    attendance_sorted = sorted(attendance, key=lambda x: x.timestamp)

    for record in attendance_sorted:
        name = user_map.get(record.user_id, "Unknown")
        ws.append([
            record.user_id,
            name,
            record.timestamp.strftime("%Y-%m-%d"),
            record.timestamp.strftime("%H:%M:%S"),
            status_label(record.status),
            "Yes" if getattr(record, "punch", None) is not None else "N/A"
        ])

    # Auto-fit column widths (approx)
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

    wb.save(filename)
    print(f"✅ Report saved: {os.path.abspath(filename)}")
    return filename


def main():
    conn = connect_device()
    try:
        users, attendance, user_map = fetch_data(conn)
        if not attendance:
            print("⚠️  No attendance records found on device.")
        else:
            export_to_excel(attendance, user_map)
    finally:
        conn.disconnect()
        print("Disconnected from device.")


if __name__ == "__main__":
    main()
