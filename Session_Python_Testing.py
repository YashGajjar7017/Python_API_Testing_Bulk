import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import requests
import threading

# API Mapping directly extracted from the provided specification sheet
API_ENDPOINTS = [
    {
        "sr": 1,
        "name": "Authentication API (Write API Login)",
        "url": "http://192.168.4.1/api/login",
        "method": "POST",
        "roles": {"Viewer": ["Write"], "Operator": ["Write"], "System Admin": ["Write"], "Security Admin": ["Write"]}
    },
    {
        "sr": 2,
        "name": "Authentication API (Read API Login)",
        "url": "http://192.168.4.1/api/auth/status",
        "method": "GET",
        "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}
    },
    {
        "sr": 3,
        "name": "UUID Checking Single-phase",
        "url": "http://192.168.4.1/api/config/parameters",
        "method": "POST",
        "roles": {"System Admin": ["Write"]}
    },
    {
        "sr": 4,
        "name": "UUID Checking Single-phase(Get)",
        "url": "http://192.168.4.1/api/config/parameters?table=1",
        "method": "GET",
        "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}
    },
    {
        "sr": 5,
        "name": "ISP Configuration API",
        "url": "http://192.168.4.1/api/config/ISP",
        "method": "POST",
        "roles": {"Security Admin": ["Write"]}
    },
    {
        "sr": 6,
        "name": "Get ISP Configuration",
        "url": "http://192.168.4.1/api/config/ISP",
        "method": "GET",
        "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}
    },
    {
        "sr": 7,
        "name": "Remote Server Configuration API",
        "url": "http://192.168.4.1/api/config/remote-server",
        "method": "POST",
        "roles": {"Security Admin": ["Write"]}
    },
    {
        "sr": 8,
        "name": "Remote Server Configuration Read API",
        "url": "http://192.168.4.1/api/config/remote-server",
        "method": "GET",
        "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}
    },
    {
        "sr": 10,
        "name": "Secure Broker Connection Trigger & Status API",
        "url": "http://192.168.4.1/api/device/broker/connect",
        "method": "POST",
        "roles": {"Security Admin": ["Write"]}
    },
    {
        "sr": 11,
        "name": "Read API – Broker Connection Status",
        "url": "http://192.168.4.1/api/device/broker/status",
        "method": "GET",
        "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}
    },
    {
        "sr": 12,
        "name": "Read API – Get Inverter Communication Configuration",
        "url": "http://192.168.4.1/api/config/inverter-communication",
        "method": "GET",
        "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}
    },
    {
        "sr": 13,
        "name": "Write API – Get Inverter Communication Configuration",
        "url": "http://192.168.4.1/api/config/inverter-communication",
        "method": "POST",
        "roles": {"System Admin": ["Write"]}
    },
    {
        "sr": 15,
        "name": "Offline Historical Data Download API",
        "url": "http://192.168.4.1/api/history?day=2026-04-21&vd=5&o",
        "method": "GET",
        "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}
    },
    {
        "sr": 16,
        "name": "WIFI Connection Check",
        "url": "http://192.168.4.1/api/device/config/update",
        "method": "GET",
        "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}
    },
    {
        "sr": 17,
        "name": "WIFI Connection Check_2",
        "url": "http://192.168.4.1/api/device/config/update",
        "method": "POST",
        "roles": {"Security Admin": ["Write"]}
    },
    {
        "sr": 18,
        "name": "Certificate RootCA",
        "url": "http://192.168.4.1/write.html?filename=rootCA.pem",
        "method": "POST",
        "roles": {"Security Admin": ["Write"]}
    },
    {
        "sr": 19,
        "name": "Certificate Key",
        "url": "http://192.168.4.1/write.html?filename=key.pem",
        "method": "POST",
        "roles": {"Security Admin": ["Write"]}
    },
    {
        "sr": 20,
        "name": "Certificate Client",
        "url": "http://192.168.4.1/write.html?filename=client.pem",
        "method": "POST",
        "roles": {"Security Admin": ["Write"]}
    },
    {
        "sr": 21,
        "name": "Restart",
        "url": "http://192.168.4.1/restart",
        "method": "POST",
        "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}
    },
    {
        "sr": 23,
        "name": "MQTTServer Get",
        "url": "http://192.168.4.1/api/config/mqtt-server",
        "method": "GET",
        "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}
    },
    {
        "sr": 24,
        "name": "MQTTServer Post",
        "url": "http://192.168.4.1/api/config/mqtt-server",
        "method": "POST",
        "roles": {"Security Admin": ["Write"]}
    },
    {
        "sr": 25,
        "name": "Firmware Update",
        "url": "http://192.168.4.1/update",
        "method": "POST",
        "roles": {"Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read", "Write"]}
    },
    {
        "sr": 26,
        "name": "Modbus Poll Acess",
        "url": "http://192.168.4.1/api/modbus",
        "method": "POST",
        "roles": {"Operator": ["Read", "Write"]}
    },
    {
        "sr": 27,
        "name": "Fiddler Request",
        "url": "http://192.168.4.1:85/list.html",
        "method": "GET",
        "roles": {}
    }
]

class APITesterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Role-Based API Runner")
        self.root.geometry("1100x700")

        # Top Control Frame
        ctrl_frame = ttk.LabelFrame(root, text=" Configuration & Controls ", padding=10)
        ctrl_frame.pack(fill="x", padx=10, pady=5)

        # Role Selection
        ttk.Label(ctrl_frame, text="Select Role:").grid(row=0, column=0, sticky="w", padx=5)
        self.role_var = tk.StringVar(value="Viewer")
        roles = ["Viewer", "Operator", "System Admin", "Security Admin"]
        self.role_cb = ttk.Combobox(ctrl_frame, textvariable=self.role_var, values=roles, state="readonly", width=18)
        self.role_cb.grid(row=0, column=1, padx=5, pady=5)
        self.role_cb.bind("<<ComboboxSelected>>", self.on_role_change)

        # Session Key Value
        ttk.Label(ctrl_frame, text="Session Value (Cookie):").grid(row=0, column=2, sticky="w", padx=10)
        self.session_var = tk.StringVar(value="3fa9a12da510c17d")
        self.session_entry = ttk.Entry(ctrl_frame, textvariable=self.session_var, width=35)
        self.session_entry.grid(row=0, column=3, padx=5, pady=5)

        # Run Button
        self.run_btn = ttk.Button(ctrl_frame, text="Run Authorized APIs", command=self.start_api_execution)
        self.run_btn.grid(row=0, column=4, padx=15, pady=5)

        # Table Display
        table_frame = ttk.Frame(root, padding=10)
        table_frame.pack(fill="both", expand=True)

        columns = ("sr", "name", "method", "url", "allowed", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("sr", text="Sr.")
        self.tree.heading("name", text="API Name")
        self.tree.heading("method", text="Method")
        self.tree.heading("url", text="URL")
        self.tree.heading("allowed", text="Role Access")
        self.tree.heading("status", text="Status")

        self.tree.column("sr", width=40, anchor="center")
        self.tree.column("name", width=260)
        self.tree.column("method", width=70, anchor="center")
        self.tree.column("url", width=340)
        self.tree.column("allowed", width=100, anchor="center")
        self.tree.column("status", width=150, anchor="center")

        # Scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Console Log Box
        log_frame = ttk.LabelFrame(root, text=" Execution Output Logs ", padding=5)
        log_frame.pack(fill="x", padx=10, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, state="disabled", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True)

        # Initial Population
        self.populate_table()

    def log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def populate_table(self):
        self.tree.delete(*self.tree.get_children())
        selected_role = self.role_var.get()

        for api in API_ENDPOINTS:
            permissions = api["roles"].get(selected_role, [])
            allowed_str = "/".join(permissions) if permissions else "NO ACCESS"

            self.tree.insert(
                "",
                "end",
                iid=str(api["sr"]),
                values=(
                    api["sr"],
                    api["name"],
                    api["method"],
                    api["url"],
                    allowed_str,
                    "Pending" if permissions else "Skipped (Unauthorized)"
                )
            )

    def on_role_change(self, event=None):
        self.populate_table()

    def start_api_execution(self):
        session_val = self.session_var.get().strip()
        if not session_val:
            messagebox.showwarning("Input Error", "Please provide a valid session ID.")
            return

        self.run_btn.config(state="disabled")
        self.log("=" * 60)
        self.log(f"Starting execution for Role: {self.role_var.get()}")
        self.log(f"Session Header Set: Cookie: session={session_val}")
        self.log("=" * 60)

        threading.Thread(target=self.execute_apis, daemon=True).start()

    def execute_apis(self):
        selected_role = self.role_var.get()
        session_id = self.session_var.get().strip()

        # Cookie header setting
        headers = {
            "Cookie": f"session={session_id}",
            "User-Agent": "API-Tester-GUI"
        }

        for api in API_ENDPOINTS:
            sr_id = str(api["sr"])
            permissions = api["roles"].get(selected_role, [])

            if not permissions:
                self.tree.item(sr_id, values=(api["sr"], api["name"], api["method"], api["url"], "NO ACCESS", "Skipped"))
                continue

            self.tree.item(sr_id, values=(api["sr"], api["name"], api["method"], api["url"], "/".join(permissions), "Running..."))
            self.log(f"[EXEC] [{api['method']}] {api['name']} -> {api['url']}")

            try:
                if api["method"] == "GET":
                    res = requests.get(api["url"], headers=headers, timeout=5)
                else:  # POST
                    res = requests.post(api["url"], headers=headers, timeout=5)

                status_msg = f"{res.status_code} {res.reason}"
                self.tree.item(sr_id, values=(api["sr"], api["name"], api["method"], api["url"], "/".join(permissions), status_msg))
                self.log(f"  └─ Status: {status_msg}")

            except requests.exceptions.RequestException as e:
                self.tree.item(sr_id, values=(api["sr"], api["name"], api["method"], api["url"], "/".join(permissions), "Failed/Timeout"))
                self.log(f"  └─ Error: {e}")

        self.log("=" * 60)
        self.log("Execution Finished.")
        self.log("=" * 60)
        self.root.after(0, lambda: self.run_btn.config(state="normal"))


if __name__ == "__main__":
    root = tk.Tk()
    app = APITesterApp(root)
    root.mainloop()