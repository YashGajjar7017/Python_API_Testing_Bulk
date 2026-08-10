import os
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import requests

# Excel export library
import openpyxl

# PDF export library
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Default Role-Password Mapping
USER_CREDENTIALS = {
    "Viewer": {"username": "viewer", "password": "viewer_001"},
    "Operator": {"username": "operator", "password": "operator_001"},
    "System Admin": {"username": "sysadmin", "password": "sysadmin_001"},
    "Security Admin": {"username": "secadmin", "password": "secadmin_001"}
}

# Full Endpoints Specification
API_ENDPOINTS = [
    {"sr": 1, "name": "Authentication API (Write API Login)", "url": "http://192.168.4.1/api/login", "method": "POST", "roles": {"Viewer": ["Write"], "Operator": ["Write"], "System Admin": ["Write"], "Security Admin": ["Write"]}},
    {"sr": 2, "name": "Authentication API (Read API Login)", "url": "http://192.168.4.1/api/auth/status", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}},
    {"sr": 3, "name": "UUID Checking Single-phase", "url": "http://192.168.4.1/api/config/parameters", "method": "POST", "roles": {"System Admin": ["Write"]}},
    {"sr": 4, "name": "UUID Checking Single-phase(Get)", "url": "http://192.168.4.1/api/config/parameters?table=1", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}},
    {"sr": 5, "name": "ISP Configuration API", "url": "http://192.168.4.1/api/config/ISP", "method": "POST", "roles": {"Security Admin": ["Write"]}},
    {"sr": 6, "name": "Get ISP Configuration", "url": "http://192.168.4.1/api/config/ISP", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}},
    {"sr": 7, "name": "Remote Server Configuration API", "url": "http://192.168.4.1/api/config/remote-server", "method": "POST", "roles": {"Security Admin": ["Write"]}},
    {"sr": 8, "name": "Remote Server Configuration Read API", "url": "http://192.168.4.1/api/config/remote-server", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}},
    {"sr": 10, "name": "Secure Broker Connection Trigger & Status API", "url": "http://192.168.4.1/api/device/broker/connect", "method": "POST", "roles": {"Security Admin": ["Write"]}},
    {"sr": 11, "name": "Read API – Broker Connection Status", "url": "http://192.168.4.1/api/device/broker/status", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}},
    {"sr": 12, "name": "Read API – Get Inverter Communication Configuration", "url": "http://192.168.4.1/api/config/inverter-communication", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}},
    {"sr": 13, "name": "Write API – Get Inverter Communication Configuration", "url": "http://192.168.4.1/api/config/inverter-communication", "method": "POST", "roles": {"System Admin": ["Write"]}},
    {"sr": 15, "name": "Offline Historical Data Download API", "url": "http://192.168.4.1/api/history?day=2026-04-21&vd=5&o", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}},
    {"sr": 16, "name": "WIFI Connection Check", "url": "http://192.168.4.1/api/device/config/update", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}},
    {"sr": 17, "name": "WIFI Connection Check_2", "url": "http://192.168.4.1/api/device/config/update", "method": "POST", "roles": {"Security Admin": ["Write"]}},
    {"sr": 18, "name": "Certificate RootCA", "url": "http://192.168.4.1/write.html?filename=rootCA.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}},
    {"sr": 19, "name": "Certificate Key", "url": "http://192.168.4.1/write.html?filename=key.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}},
    {"sr": 20, "name": "Certificate Client", "url": "http://192.168.4.1/write.html?filename=client.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}},
    {"sr": 21, "name": "Restart", "url": "http://192.168.4.1/restart", "method": "POST", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}},
    {"sr": 23, "name": "MQTTServer Get", "url": "http://192.168.4.1/api/config/mqtt-server", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}},
    {"sr": 24, "name": "MQTTServer Post", "url": "http://192.168.4.1/api/config/mqtt-server", "method": "POST", "roles": {"Security Admin": ["Write"]}},
    {"sr": 25, "name": "Firmware Update", "url": "http://192.168.4.1/update", "method": "POST", "roles": {"Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read", "Write"]}},
    {"sr": 26, "name": "Modbus Poll Acess", "url": "http://192.168.4.1/api/modbus", "method": "POST", "roles": {"Operator": ["Read", "Write"]}},
    {"sr": 27, "name": "Fiddler Request", "url": "http://192.168.4.1:85/list.html", "method": "GET", "roles": {}}
]

class RMSDeviceTesterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RMS Device Direct AP Tester")
        self.root.geometry("1280x750")

        self.execution_results = []

        # Header Frame: Connection & User Setup
        setup_frame = ttk.LabelFrame(root, text=" Device Connection & Authentication ", padding=10)
        setup_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(setup_frame, text="Device AP Name:").grid(row=0, column=0, sticky="w", padx=5)
        self.ap_name_var = tk.StringVar(value="RMS-0001")
        ttk.Entry(setup_frame, textvariable=self.ap_name_var, width=15).grid(row=0, column=1, padx=5)

        ttk.Label(setup_frame, text="Select Role:").grid(row=0, column=2, sticky="w", padx=10)
        self.role_var = tk.StringVar(value="Viewer")
        self.role_cb = ttk.Combobox(setup_frame, textvariable=self.role_var, values=list(USER_CREDENTIALS.keys()), state="readonly", width=15)
        self.role_cb.grid(row=0, column=3, padx=5)
        self.role_cb.bind("<<ComboboxSelected>>", self.on_role_change)

        ttk.Label(setup_frame, text="Password:").grid(row=0, column=4, sticky="w", padx=10)
        self.password_var = tk.StringVar(value=USER_CREDENTIALS["Viewer"]["password"])
        ttk.Entry(setup_frame, textvariable=self.password_var, width=15, show="*").grid(row=0, column=5, padx=5)

        ttk.Label(setup_frame, text="Session Token:").grid(row=0, column=6, sticky="w", padx=10)
        self.session_token_var = tk.StringVar(value="")
        ttk.Entry(setup_frame, textvariable=self.session_token_var, width=25).grid(row=0, column=7, padx=5)

        self.run_btn = ttk.Button(setup_frame, text="Authenticate & Run APIs", command=self.start_process)
        self.run_btn.grid(row=0, column=8, padx=15)

        # Middle Paned Window (Treeview on Left, Text Output on Right)
        paned = ttk.PanedWindow(root, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=10, pady=5)

        # Left Table Frame
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=3)

        cols = ("sr", "name", "method", "url", "status")
        self.tree = ttk.Treeview(left_frame, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("sr", text="Sr.")
        self.tree.heading("name", text="API Name")
        self.tree.heading("method", text="Method")
        self.tree.heading("url", text="URL")
        self.tree.heading("status", text="Status")

        self.tree.column("sr", width=35, anchor="center")
        self.tree.column("name", width=220)
        self.tree.column("method", width=60, anchor="center")
        self.tree.column("url", width=260)
        self.tree.column("status", width=120, anchor="center")

        tree_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=tree_scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_select_tree_item)

        # Right Response View Frame
        right_frame = ttk.LabelFrame(paned, text=" API Response Output ", padding=5)
        paned.add(right_frame, weight=2)

        self.output_text = scrolledtext.ScrolledText(right_frame, font=("Consolas", 9), wrap="word")
        self.output_text.pack(fill="both", expand=True)

        # Bottom Bar (Export Buttons)
        bottom_frame = ttk.Frame(root, padding=5)
        bottom_frame.pack(fill="x", padx=10, pady=5)

        self.status_lbl = ttk.Label(bottom_frame, text="Status: Ready", font=("Arial", 9, "bold"))
        self.status_lbl.pack(side="left", padx=5)

        ttk.Button(bottom_frame, text="Export PDF", command=self.export_pdf).pack(side="right", padx=5)
        ttk.Button(bottom_frame, text="Export Excel", command=self.export_excel).pack(side="right", padx=5)

        self.populate_table()

    def on_role_change(self, event=None):
        role = self.role_var.get()
        self.password_var.set(USER_CREDENTIALS[role]["password"])
        self.populate_table()

    def populate_table(self):
        self.tree.delete(*self.tree.get_children())
        selected_role = self.role_var.get()

        for api in API_ENDPOINTS:
            permissions = api["roles"].get(selected_role, [])
            status = "Pending" if permissions else "Unauthorized"
            self.tree.insert("", "end", iid=str(api["sr"]), values=(api["sr"], api["name"], api["method"], api["url"], status))

    def on_select_tree_item(self, event):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        sr_id = int(selected_item[0])
        for res in self.execution_results:
            if res["sr"] == sr_id:
                self.output_text.delete("1.0", tk.END)
                self.output_text.insert(tk.END, f"API: {res['name']}\nURL: {res['url']}\nMethod: {res['method']}\nStatus: {res['status']}\n\n--- RESPONSE BODY ---\n{res['body']}")
                break

    def start_process(self):
        self.run_btn.config(state="disabled")
        self.execution_results.clear()
        threading.Thread(target=self.run_execution_pipeline, daemon=True).start()

    def run_execution_pipeline(self):
        ap_name = self.ap_name_var.get().strip()
        role = self.role_var.get()
        password = self.password_var.get().strip()

        # Step 1: Login Request to acquire Session Token
        self.update_status(f"Authenticating as {role} at http://192.168.4.1/api/login ...")
        login_url = "http://192.168.4.1/api/login"
        payload = {"username": USER_CREDENTIALS[role]["username"], "password": password}

        session_token = ""
        try:
            resp = requests.post(login_url, json=payload, timeout=5)
            if resp.status_code == 200:
                # Try JSON response or cookie header
                data = resp.json() if resp.headers.get("content-type") == "application/json" else {}
                session_token = data.get("sessionToken") or resp.cookies.get("session") or "3fa9a12da510c17d"
            else:
                session_token = "3fa9a12da510c17d"  # Fallback session token
        except Exception as e:
            session_token = "3fa9a12da510c17d"  # Fallback session token on failure

        self.session_token_var.set(session_token)
        headers = {"Cookie": f"session={session_token}"}

        # Step 2: Iterate and call each API with 30s delay
        apis = [a for a in API_ENDPOINTS if a["roles"].get(role)]

        for idx, api in enumerate(apis):
            sr_id = str(api["sr"])
            self.update_status(f"Running API {idx+1}/{len(apis)}: {api['name']}...")
            self.tree.item(sr_id, values=(api["sr"], api["name"], api["method"], api["url"], "Running..."))

            body_output = ""
            status_str = ""

            try:
                if api["method"] == "GET":
                    r = requests.get(api["url"], headers=headers, timeout=5)
                else:
                    r = requests.post(api["url"], headers=headers, timeout=5)

                status_str = f"{r.status_code} {r.reason}"
                body_output = r.text
            except Exception as ex:
                status_str = "Error/Offline"
                body_output = f"Connection Error: {str(ex)}"

            self.tree.item(sr_id, values=(api["sr"], api["name"], api["method"], api["url"], status_str))
            
            res_entry = {
                "sr": api["sr"],
                "name": api["name"],
                "method": api["method"],
                "url": api["url"],
                "status": status_str,
                "body": body_output
            }
            self.execution_results.append(res_entry)

            # Auto display latest output
            self.root.after(0, lambda res=res_entry: self.display_output(res))

            # 30 seconds delay between API calls
            if idx < len(apis) - 1:
                for remaining in range(30, 0, -1):
                    self.update_status(f"Waiting 30s delay... ({remaining}s remaining before next API)")
                    time.sleep(1)

        self.update_status("Execution Completed!")
        self.root.after(0, lambda: self.run_btn.config(state="normal"))

    def display_output(self, res):
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, f"Device: {self.ap_name_var.get()}\nAPI: {res['name']}\nURL: {res['url']}\nStatus: {res['status']}\n\n--- RAW OUTPUT ---\n{res['body']}")

    def update_status(self, text):
        self.root.after(0, lambda: self.status_lbl.config(text=f"Status: {text}"))

    # Export Report to Excel
    def export_excel(self):
        if not self.execution_results:
            messagebox.showinfo("Export Error", "No execution results available to export.")
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "API Execution Results"

        ws.append(["Device AP Name", self.ap_name_var.get()])
        ws.append(["User Role", self.role_var.get()])
        ws.append(["Session Token", self.session_token_var.get()])
        ws.append([])
        ws.append(["Sr.", "API Name", "Method", "URL", "Status", "Response Output"])

        for item in self.execution_results:
            ws.append([item["sr"], item["name"], item["method"], item["url"], item["status"], item["body"][:500]])

        filename = f"{self.ap_name_var.get()}_API_Report.xlsx"
        wb.save(filename)
        messagebox.showinfo("Export Successful", f"Excel report saved as {filename}")

    # Export Report to PDF
    def export_pdf(self):
        if not self.execution_results:
            messagebox.showinfo("Export Error", "No execution results available to export.")
            return

        filename = f"{self.ap_name_var.get()}_API_Report.pdf"
        doc = SimpleDocTemplate(filename, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []

        title_style = ParagraphStyle("TitleStyle", parent=styles["Heading1"], fontSize=16, leading=20)
        story.append(Paragraph(f"RMS Device API Test Summary Report", title_style))
        story.append(Spacer(1, 10))

        info_text = f"<b>Device Name:</b> {self.ap_name_var.get()}<br/>" \
                    f"<b>User Role:</b> {self.role_var.get()}<br/>" \
                    f"<b>Session Token:</b> {self.session_token_var.get()}"
        story.append(Paragraph(info_text, styles["Normal"]))
        story.append(Spacer(1, 15))

        table_data = [["Sr.", "API Name", "Method", "Status"]]
        for item in self.execution_results:
            table_data.append([str(item["sr"]), item["name"], item["method"], item["status"]])

        t = Table(table_data, colWidths=[30, 260, 60, 150])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ]))

        story.append(t)
        doc.build(story)
        messagebox.showinfo("Export Successful", f"PDF report saved as {filename}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RMSDeviceTesterApp(root)
    root.mainloop()