"""
AttendanceTracker — main GUI entry point.
On launch: shows Device IP field, auto-fetches attendance data,
updates both databases, computes payroll, and displays the dashboard.
"""

import os
import sys
import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import attendance_engine
import db_manager
import payroll_manager
import charts


def get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = get_base_dir()
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")


def load_config():
    default = {"device_ip": "192.168.1.34", "drive_sync_folder": "",
               "default_hourly_rate": 100.0, "standard_workday_hours": 9}
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r") as f:
            default.update(json.load(f))
    return default


def save_config(cfg):
    with open(CONFIG_PATH, "w") as f:
        json.dump(cfg, f, indent=2)


class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Tracker & Payroll Dashboard")
        self.root.geometry("1100x750")
        self.cfg = load_config()

        self.attendance_db = db_manager.get_db_path(BASE_DIR)
        self.payroll_db = payroll_manager.get_db_path(BASE_DIR)

        self._build_top_bar()
        self._build_tabs()
        self.status_var.set("Ready.")

        # Auto-fetch on startup
        self.root.after(500, self.fetch_and_refresh)

    # ---------------- Top control bar ----------------
    def _build_top_bar(self):
        bar = ttk.Frame(self.root, padding=10)
        bar.pack(fill="x")

        ttk.Label(bar, text="Device IP:").pack(side="left")
        self.ip_var = tk.StringVar(value=self.cfg.get("device_ip", "192.168.1.34"))
        ttk.Entry(bar, textvariable=self.ip_var, width=16).pack(side="left", padx=5)

        ttk.Button(bar, text="Fetch Latest & Update", command=self.fetch_and_refresh).pack(side="left", padx=5)

        ttk.Label(bar, text="   Google Drive Folder:").pack(side="left")
        self.drive_var = tk.StringVar(value=self.cfg.get("drive_sync_folder", ""))
        ttk.Entry(bar, textvariable=self.drive_var, width=35).pack(side="left", padx=5)
        ttk.Button(bar, text="Browse...", command=self.browse_drive_folder).pack(side="left")

        self.status_var = tk.StringVar(value="")
        ttk.Label(self.root, textvariable=self.status_var, foreground="#2980b9",
                  padding=(10, 0)).pack(fill="x")

    def browse_drive_folder(self):
        path = filedialog.askdirectory(title="Select your local Google Drive synced folder")
        if path:
            self.drive_var.set(path)
            self.cfg["drive_sync_folder"] = path
            save_config(self.cfg)

    # ---------------- Tabs ----------------
    def _build_tabs(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_overview = ttk.Frame(self.notebook)
        self.tab_employee = ttk.Frame(self.notebook)
        self.tab_payroll = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_overview, text="📊 Dashboard Overview")
        self.notebook.add(self.tab_employee, text="🔍 Employee Drill-Down")
        self.notebook.add(self.tab_payroll, text="💰 Payroll (Per-Day Salary)")

    # ---------------- Fetch + Update pipeline ----------------
    def fetch_and_refresh(self):
        ip = self.ip_var.get().strip()
        self.cfg["device_ip"] = ip
        save_config(self.cfg)
        self.status_var.set(f"Connecting to {ip} ...")
        threading.Thread(target=self._background_fetch, args=(ip,), daemon=True).start()

    def _background_fetch(self, ip):
        try:
            records = attendance_engine.fetch_and_process(ip, self.cfg.get("standard_workday_hours", 9))
            db_manager.update_attendance_db(self.attendance_db, records)

            daily_df, monthly_df = db_manager.load_dataframes(self.attendance_db)
            payroll_manager.compute_and_store_daily_pay(
                self.payroll_db, daily_df, self.cfg.get("default_hourly_rate", 100.0)
            )

            drive_folder = self.drive_var.get().strip()
            msg1, msg2 = "", ""
            if drive_folder:
                ok1, msg1 = db_manager.sync_file_to_drive(self.attendance_db, drive_folder)
                ok2, msg2 = db_manager.sync_file_to_drive(self.payroll_db, drive_folder)

            self.root.after(0, lambda: self._on_fetch_success(drive_folder, msg1, msg2))
        except Exception as e:
            self.root.after(0, lambda: self._on_fetch_error(str(e)))

    def _on_fetch_success(self, drive_folder, msg1, msg2):
        self.status_var.set(f"Updated successfully. {msg1} {msg2}".strip())
        self.refresh_dashboard()

    def _on_fetch_error(self, error_msg):
        self.status_var.set(f"Error: {error_msg}")
        messagebox.showerror("Fetch Failed", f"Could not fetch data:\n{error_msg}")

    # ---------------- Populate dashboard tabs ----------------
    def refresh_dashboard(self):
        daily_df, monthly_df = db_manager.load_dataframes(self.attendance_db)
        daily_pay_df, monthly_payroll_df = payroll_manager.load_payroll_data(self.payroll_db)

        self._populate_overview(monthly_df)
        self._populate_employee_tab(daily_df, monthly_df)
        self._populate_payroll_tab(daily_pay_df, monthly_payroll_df)

    def _clear_frame(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()

    def _populate_overview(self, monthly_df):
        self._clear_frame(self.tab_overview)
        frame = ttk.Frame(self.tab_overview)
        frame.pack(fill="both", expand=True)

        fig1 = charts.hours_trend_figure(monthly_df)
        canvas1 = FigureCanvasTkAgg(fig1, master=frame)
        canvas1.get_tk_widget().pack(side="left", fill="both", expand=True)
        canvas1.draw()

        fig2 = charts.excess_deficit_figure(monthly_df)
        canvas2 = FigureCanvasTkAgg(fig2, master=frame)
        canvas2.get_tk_widget().pack(side="left", fill="both", expand=True)
        canvas2.draw()

    def _populate_employee_tab(self, daily_df, monthly_df):
        self._clear_frame(self.tab_employee)
        left = ttk.Frame(self.tab_employee, width=180)
        left.pack(side="left", fill="y")
        right = ttk.Frame(self.tab_employee)
        right.pack(side="left", fill="both", expand=True)

        ttk.Label(left, text="Employees").pack()
        listbox = tk.Listbox(left)
        listbox.pack(fill="y", expand=True)
        names = sorted(daily_df["name"].unique()) if not daily_df.empty else []
        for n in names:
            listbox.insert("end", n)

        def on_select(event):
            if not listbox.curselection():
                return
            name = listbox.get(listbox.curselection()[0])
            self._show_employee_detail(right, daily_df, name)

        listbox.bind("<<ListboxSelect>>", on_select)

    def _show_employee_detail(self, container, daily_df, name):
        self._clear_frame(container)
        emp_df = daily_df[daily_df["name"] == name].sort_values("date")

        cols = ["date", "clock_in", "clock_out", "total_hours", "excess_hours", "deficit_hours", "flags"]
        tree = ttk.Treeview(container, columns=cols, show="headings", height=10)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=100)
        for _, row in emp_df.iterrows():
            tree.insert("", "end", values=[row[c] for c in cols])
        tree.pack(fill="x")

        fig = charts.employee_hours_figure(emp_df, name)
        canvas = FigureCanvasTkAgg(fig, master=container)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        canvas.draw()

    def _populate_payroll_tab(self, daily_pay_df, monthly_payroll_df):
        self._clear_frame(self.tab_payroll)

        top = ttk.Frame(self.tab_payroll)
        top.pack(fill="x", pady=5)
        ttk.Label(top, text="Select Employee:").pack(side="left")

        names = sorted(daily_pay_df["name"].unique()) if not daily_pay_df.empty else []
        emp_var = tk.StringVar(value=names[0] if names else "")
        combo = ttk.Combobox(top, textvariable=emp_var, values=names, state="readonly")
        combo.pack(side="left", padx=5)

        ttk.Button(top, text="Update Hourly Rate",
                   command=lambda: self._update_rate_dialog(emp_var.get())).pack(side="left", padx=5)

        table_frame = ttk.Frame(self.tab_payroll)
        table_frame.pack(fill="both", expand=True)

        cols = ["date", "total_hours", "hourly_rate", "daily_pay"]
        tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=10)
        for c in cols:
            tree.heading(c, text=c)
        tree.pack(side="left", fill="both", expand=True)

        def refresh_table(*_):
            tree.delete(*tree.get_children())
            emp_df = daily_pay_df[daily_pay_df["name"] == emp_var.get()].sort_values("date")
            for _, row in emp_df.iterrows():
                tree.insert("", "end", values=[row[c] for c in cols])

        combo.bind("<<ComboboxSelected>>", refresh_table)
        refresh_table()

        ttk.Label(self.tab_payroll, text="Monthly Payroll Summary (All Employees)",
                  font=("Arial", 10, "bold")).pack(pady=(10, 0))
        summary_cols = ["name", "year_month", "days_paid", "total_hours", "total_pay"]
        summary_tree = ttk.Treeview(self.tab_payroll, columns=summary_cols, show="headings", height=8)
        for c in summary_cols:
            summary_tree.heading(c, text=c)
        summary_tree.pack(fill="x")
        if not monthly_payroll_df.empty:
            for _, row in monthly_payroll_df.iterrows():
                summary_tree.insert("", "end", values=[row[c] for c in summary_cols])

    def _update_rate_dialog(self, name):
        if not name:
            return
        new_rate = simpledialog.askfloat("Update Hourly Rate", f"New hourly rate for {name}:")
        if new_rate is not None:
            import sqlite3
            conn = sqlite3.connect(self.payroll_db)
            payroll_manager.update_rate(conn, name, new_rate)
            conn.close()
            self.status_var.set(f"Updated rate for {name} to {new_rate}. Recomputing payroll...")
            self.fetch_and_refresh()


if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()
