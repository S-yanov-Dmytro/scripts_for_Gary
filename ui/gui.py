import tkinter as tk
from tkinter import ttk
from config import sites
from utils.file_operations import load_xls_file
from automation.coral import run_automation_coral
from automation.ladbrokes import run_automation_ladbrokes
from automation.betway import run_automation_betway
from automation.netbet import run_automation_netbet
from automation.betvictor import run_automation_betvictor

import threading

class AutomationApp:
    def __init__(self, root):
        self.root = root
        self.selected_site = "Ladbrokes"
        self.data_list = []
        self.file_path = ""
        self.setup_ui()

    def setup_ui(self):
        self.root.title("Registration Automation")
        self.root.geometry("600x500")

        frame = tk.Frame(self.root, padx=30, pady=30, bg="#ffffff", relief=tk.GROOVE, bd=3)
        frame.pack(pady=30, padx=30, fill=tk.BOTH, expand=True)

        self.site_var = tk.StringVar(value=self.selected_site)
        site_label = tk.Label(frame, text="Select site:", font=("Arial", 14), bg="#ffffff")
        site_label.pack(pady=5)

        site_dropdown = ttk.Combobox(frame, textvariable=self.site_var, values=list(sites.keys()), state="readonly", font=("Arial", 12))
        site_dropdown.pack(pady=5, fill=tk.X)
        site_dropdown.bind("<<ComboboxSelected>>", self.select_site)

        btn_load = tk.Button(frame, text="Load Excel", command=self.load_file, font=("Arial", 14), bg="#007BFF", fg="white", height=2)
        btn_load.pack(pady=10)

        self.btn_start = tk.Button(frame, text="Start", command=self.start_thread, font=("Arial", 14), bg="#28A745", fg="white", height=2, state="disabled")
        self.btn_start.pack(pady=10)

        self.lbl_status = tk.Label(frame, text="", font=("Arial", 14), bg="#ffffff")
        self.lbl_status.pack(pady=10)

    def select_site(self, event):
        self.selected_site = self.site_var.get()

    def load_file(self):
        data, path = load_xls_file()
        if data:
            self.data_list = data
            self.file_path = path
            self.lbl_status.config(text="Data loaded!", fg="green")
            self.btn_start.config(state="normal")

    def start_thread(self):
        func_map = {
            "Coral": run_automation_coral,
            "betway": run_automation_betway,
            "Ladbrokes": run_automation_ladbrokes,
            "netbet": run_automation_netbet,
            "betvictor": run_automation_betvictor
        }

        func = func_map.get(self.selected_site, run_automation_ladbrokes)
        thread = threading.Thread(target=func, args=(self.data_list, self.file_path, self.lbl_status), daemon=True)
        thread.start()
