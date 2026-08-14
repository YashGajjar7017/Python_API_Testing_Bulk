import os
import time
import json
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

# Default Role Passwords
USER_CREDENTIALS = {
    "Viewer": "viewer_001",
    "Operator": "operator_001",
    "System Admin": "sysadmin_001",
    "Security Admin": "secadmin_001"
}

# Raw Certificate payloads
ROOT_CA_PEM = """-----BEGIN CERTIFICATE-----
MIICyTCCAbGgAwIBAgIUNNDf2Nc1apj4ZNrhqdlEh4lGvwQwDQYJKoZIhvcNAQEL
BQAwDzENMAsGA1UEAwwEcm9vdDAeFw0yNTA5MDIxNjE4MjlaFw0zMjA4MzExNjE4
MjlaMA8xDTALBgNVBAMMBHJvb3QwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEK
AoIBAQDjWfgumybA/tq/saHcDEE/+KMgWdfOI+iRhE4Ud/YLXeaoN7ruVBjuwNMT
8s2CJxuKP4p++psr1T+NtxK8QyKTs+BQc1MHGI/SC5oQ2wBBUlLIvQ/pifPOuVFO
yEogGmCvv8XikQmkmo/alyUnf9KmVEdCs4PDk8Uf67QgYLXpeA/fxshk7kPhclU8
TRbICMVBXbL+VKX1H5xexbAbTCYBUbJP9p+zUVy8IS8A4ZFNfDWuV2L690o5XfvG
toCxnrvFusEgB/JViXqnEKu44Xj+jXsi/wZz0parY2eEQlfI4ZGyx/SUDoEGBsoa
pdowg5zLLFgwSEb0QDA+8pzq4c97AgMBAAGjHTAbMAwGA1UdEwQFMAMBAf8wCwYD
VR0PBAQDAgEGMA0GCSqGSIb3DQEBCwUAA4IBAQAYtxxiW7tvBl4uJIUwou+fbPPo
J7bTELpUWbwe1HA4S1jHBDu/tOl1hCf+f2XsqsmS1De4V5ksN3Ez8XDmOzIYRqJZ
XMFpdXXSn6Fb2dE2MKmb9D7DYCAuoOYr9cLICJ1NJsqKs9sl61JeEWrDWxfm2xjM
6Cq06ikeEX1K6z5Bc4qrvu5rMmKBK2c+XwyguGgobj6qfGQi8rs+qiAlIoadW8Ow
BrxJwmhrm5Uq/ww/UCugaZTdgWXG4DOKz56yJMqewuF0GHk477qiRBLAd4BqGI0a
m9s3tRSU1+2IvvPvYDZXALwu+s0YF/T54OLQ3QBK6+Hylyjck9uo7iOjQQCr
-----END CERTIFICATE-----"""

KEY_PEM = """-----BEGIN RSA PRIVATE KEY-----
MIIEpAIBAAKCAQEAxJGD5KA+QYaFRL91WE04omsAGoka5we25yQEqcMQfiaI4kkl
oe0w+3jxrgQ/mRz35dqwZ4GF0b5UGDUoYDiCc63emGym3FUtywUWRcEjZ7/pD0SC
Kn/UWHRL3WhT/NR55LRwErygAiCiHOtoc/Hz5tTQuf0LxRY+XFn6fIGrwMJCM+AT
/aJKcpRniRhn5G+1SrIM7o7C9pHmXCZo7AIEWXyZi6GEEkTq1EKB2+RHWGZGFt8q
/3c5vJbfP28h5T+4poBAuoc/aRoHSSzjQw/8Ipl78ruTP1NTJVA591EwEYqCVLZJ
0WSFox0i6MWkPcs4ijL8oPqlpH/HIwg0Bxxi7wIDAQABAoIBAEcVqHDBUkI393HJ
GxJ3uodh0oPGGSh7LbWg3N7XX9t+6/mJIkbQk+oc7qlTd/xS8YL59tk3icEz5w1q
C9PgG6zhr6HHiQAc/ZZxnIBVPAyE5i9TQw9DcvMZ+2VyGQ1ReVMvvWfJNsQOMRHs
P7XORbmwNFtOGzeGt76F+4OTER1/f6zoAm1T0uTyO7azxeuRD50tppXQ8EBO677u
5kWLGx0ilW6hv/C6RFzXiUYy4+M9QNAK560qrbDr/ZX9G3UweZ0hjQUbAbqjPtKc
nEK+32didktk0UXP0lNKbCsuhHfXuJzEPRBghdf/taE7ROVbZhf3wsi3TAD0Q8rR
xZQ/GcECgYEA9cf+87IsEMckfsaKWm8LA3oCWiiUq2Tqlb2JVqYJ8WF4cZeuShc6
7FFD+Db47rBYWi7L45o0rGQV5TAEF+Oo6uZcc52z7+0wO+WcM7yZ9PhEhs60eqo+
2Wmmtt1P5Fe33/uRZxF/69sTvKoZ5sxs5EerU7sp2cBkanWGLrY3MTcCgYEAzL23
ZUgRP6rG4JVz3NCdIac8xVumdvS8UBTBZS6qbKnWPiXn6/c2h1BbaMnQ7Kjn/ESV
gc9vzUvIaeawMo1zNtrv4v1e4YFNt0cZGf5egBOgz3zTwltB7XNIY6E2ktGVnw8q
OT4rTmFGN4c0jYURTCTTrIWdanUE+Q2KI6slmAkCgYEAwLI6q2nHrqfTbynOEzT4
V5NzAKjMVsxaN2hQowSuyvb2bWjUlvY7lkfomFTROqI7wwjphdrC7V3S5MuOigRN
zU0qsuKzzqoRpQRSQSXfNvbnEyJA0eNkPyTcNoaxOn+jhBJCX6KSvqIWiVZ71D32
KcwfARi+qSB0GhtmRn1KLGUCgYBGqbfRSfVCDxQC0TIi5RFGWyz7RK2IYuFXlkJF
RLznMhrSakzIQRNd0lsqKHVmKMmgZJ9hJRIGPgCWpFrtp6o7JYcjxaTombMT7YL5
WLSO6bdXxVJkwxBW/rHeSPbH53QHaLl+9jkGHUaZxZ1atrGIoTE5WwhAhNDJkGW7
98bK2QKBgQDrRsxyO4sbmVVGLQ1+WBdyF02YTvUorGqXWuJ2WZKs4MuSmb11287c0JxVy+xAgXwQXEHKE7dLtOvR0rwQ+6Ki8hOgM+ebhiZjZOab09XaztgCvQC7akVf
ItIB0WQvjPpMSemyDNVf1q15YM/3zFzgClER99I1ZCwfh4BrUPBWgQ==
-----END RSA PRIVATE KEY-----"""

CLIENT_PEM = """-----BEGIN CERTIFICATE-----
MIIC+jCCAeKgAwIBAgICAvowDQYJKoZIhvcNAQELBQAwDzENMAsGA1UEAwwEcm9v
dDAeFw0yNTA3MDEwNzE2MDNaFw0zMzA2MjkwNzE2MDNaMEAxLTArBgNVBAMMJDg2
NjczODA4MzYwODc0MyRvbmdyaWRyb29mdG9wJDUxMDAxNzEPMA0GA1UECgwGNTEw
MDE3MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAxJGD5KA+QYaFRL91
WE04omsAGoka5we25yQEqcMQfiaI4kkloe0w+3jxrgQ/mRz35dqwZ4GF0b5UGDUo
YDiCc63emGym3FUtywUWRcEjZ7/pD0SCKn/UWHRL3WhT/NR55LRwErygAiCiHOto
c/Hz5tTQuf0LxRY+XFn6fIGrwMJCM+AT/aJKcpRniRhn5G+1SrIM7o7C9pHmXCZo
7AIEWXyZi6GEEkTq1EKB2+RHWGZGFt8q/3c5vJbfP28h5T+4poBAuoc/aRoHSSzj
Qw/8Ipl78ruTP1NTJVA591EwEYqCVLZJ0WSFox0i6MWkPcs4ijL8oPqlpH/HIwg0
Bxxi7wIDAQABoy8wLTAJBgNVHRMEAjAAMAsGA1UdDwQEAwIFoDATBgNVHSUEDDAK
BggrBgEFBQcDAjANBgkqhkiG9w0BAQsFAAOCAQEAx0QpfVs05nXkudzIP3BVerP5
QNSoQLX6YtIHfQWZ3AU8fKYym0VXkXaZyQmgkZl5yB0UPvbs2RtLhGdl/PceFATv
ctygTTYmOTjpljTAEVSkjXVtZSyqp01aZKeXQMcGPonfQdMOfwhXTelRAbHfCe4f
7r+Yws0klLdrESJs8x8bS8+pLjhkmoeneNJKRXvVUiDlf3B76/m6LlTWFgjUWy/4
ZKHDY4/D6ZqOj5y6kBsMcemDdLsPnKjHdhOdk+5u4q7DxeaBTW6tPGbnxb9+jrMo
E+FqGv4I0wdDgg9bxgzppc+ZscC0AkC5y6UCN+/1S8++CcSOSOuGeL25pdnxhg==
-----END CERTIFICATE-----"""

# Complete Master API Mapping
API_ENDPOINTS = [
    {"sr": 1, "name": "Authentication API (Write API Login)", "url": "http://192.168.4.1/api/login", "method": "POST", "roles": {"Viewer": ["Write"], "Operator": ["Write"], "System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": None},
    {"sr": 2, "name": "Authentication API (Read API Login)", "url": "http://192.168.4.1/api/auth/status", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 3, "name": "UUID Checking Single-phase", "url": "http://192.168.4.1/api/config/parameters", "method": "POST", "roles": {"System Admin": ["Write"]}, "payload": {"vdinterval": 5, "table": 1, "parameters": []}},
    {"sr": 4, "name": "UUID Checking Single-phase(Get)", "url": "http://192.168.4.1/api/config/parameters?table=1", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 5, "name": "ISP Configuration API", "url": "http://192.168.4.1/api/config/isp", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"apn": "airtelgprs.com", "apn2": "airtelgprs.com", "current_sim": "1"}},
    {"sr": 6, "name": "Get ISP Configuration", "url": "http://192.168.4.1/api/config/ISP", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 7, "name": "Remote Server Configuration API", "url": "http://192.168.4.1/api/config/remote-server", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"server_url": "rms.iotscada-pmsg.com", "server_port": 8883, "solution_type": "ongridrooftop", "client_id": "d:866738083608743$ongridrooftop$510017", "username": "866738083608743$ongridrooftop$510017", "password": "31c1074a", "server_url1": "rms.iotscada-pmsg.com", "server_port1": 8883, "solution_type1": "ongridrooftop", "client_id1": "d:866082075799828$ongridrooftop$500092", "username1": "866082075799828$ongridrooftop$500092", "password1": "466b856f", "imei": "866738083608743", "imei1": "866082075799828"}},
    {"sr": 8, "name": "Remote Server Configuration Read API", "url": "http://192.168.4.1/api/config/remote-server", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 10, "name": "Secure Broker Connection Trigger & Status API", "url": "http://192.168.4.1/api/device/broker/connect", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"action": "connect"}},
    {"sr": 11, "name": "Read API – Broker Connection Status", "url": "http://192.168.4.1/api/device/broker/status", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 12, "name": "Read API – Get Inverter Communication Configuration", "url": "http://192.168.4.1/api/config/inverter-communication", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 13, "name": "Write API – Get Inverter Communication Configuration", "url": "http://192.168.4.1/api/config/parameters", "method": "POST", "roles": {"System Admin": ["Write"]}, "payload": {"asn": "Yash", "baudrate": 9600, "parity": 1, "stopBit": 1, "databits": 8, "reqCount_1": 2}},
    {"sr": 15, "name": "Offline Historical Data Download API", "url": "http://192.168.4.1/api/history?day=2026-04-21&vd=5&o", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 16, "name": "WIFI Connection Check", "url": "http://192.168.4.1/api/device/config/update", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 17, "name": "WIFI Connection Check_2", "url": "http://192.168.4.1/api/device/config/update", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"ssid": "test_wifi"}},
    {"sr": 18, "name": "Certificate RootCA", "url": "http://192.168.4.1/write.html?filename=rootCA.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": ROOT_CA_PEM},
    {"sr": 19, "name": "Certificate Key", "url": "http://192.168.4.1/write.html?filename=key.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": KEY_PEM},
    {"sr": 20, "name": "Certificate Client", "url": "http://192.168.4.1/write.html?filename=client.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": CLIENT_PEM},
    {"sr": 21, "name": "Restart", "url": "http://192.168.4.1/restart", "method": "POST", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": {"action": "reboot"}},
    {"sr": 23, "name": "MQTTServer Get", "url": "http://192.168.4.1/api/config/mqtt-server", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 24, "name": "MQTTServer Post", "url": "http://192.168.4.1/api/config/mqtt-server", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"http_url": "api.iotscada-pmsg.com", "http_port": 443, "imei": "866738083608743", "username": "866738083608743", "password": "31c1074a"}},
    {"sr": 25, "name": "Firmware Update", "url": "http://192.168.4.1/update", "method": "POST", "roles": {"Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read", "Write"]}, "payload": {"file": "fw_v1.bin"}},
    {"sr": 26, "name": "Modbus Poll Access", "url": "http://192.168.4.1/api/modbus", "method": "POST", "roles": {"Operator": ["Read", "Write"]}, "payload": {"poll": True}},
    {"sr": 27, "name": "Fiddler Request", "url": "http://192.168.4.1:85/list.html", "method": "GET", "roles": {}, "payload": None}
]

class RMSDeviceTesterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RMS Device Direct AP Suite")
        self.root.geometry("1350x850")

        self.execution_results = []
        self.cancel_event = threading.Event()
        
        # 1-minute Timer variables
        self.session_expire_time = None
        self.timer_running = False

        # 1. Header Frame: Controls & Credentials
        setup_frame = ttk.LabelFrame(root, text=" Connection & Login Setup ", padding=8)
        setup_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(setup_frame, text="Username (AP Name):").grid(row=0, column=0, sticky="w", padx=3)
        self.ap_name_var = tk.StringVar(value="RMS-2088")
        ttk.Entry(setup_frame, textvariable=self.ap_name_var, width=12).grid(row=0, column=1, padx=3)

        ttk.Label(setup_frame, text="Role:").grid(row=0, column=2, sticky="w", padx=3)
        self.role_var = tk.StringVar(value="Viewer")
        self.role_cb = ttk.Combobox(setup_frame, textvariable=self.role_var, values=list(USER_CREDENTIALS.keys()), state="readonly", width=13)
        self.role_cb.grid(row=0, column=3, padx=3)
        self.role_cb.bind("<<ComboboxSelected>>", self.on_role_change)

        ttk.Label(setup_frame, text="Password:").grid(row=0, column=4, sticky="w", padx=3)
        self.password_var = tk.StringVar(value=USER_CREDENTIALS["Viewer"])
        ttk.Entry(setup_frame, textvariable=self.password_var, width=13).grid(row=0, column=5, padx=3)

        ttk.Button(setup_frame, text="1. Fetch Session Token (/api/login)", command=self.fetch_session_token).grid(row=0, column=6, padx=8)

        ttk.Label(setup_frame, text="Session Token:").grid(row=0, column=7, sticky="w", padx=3)
        self.session_token_var = tk.StringVar(value="")
        ttk.Entry(setup_frame, textvariable=self.session_token_var, width=20).grid(row=0, column=8, padx=3)

        self.run_btn = ttk.Button(setup_frame, text="2. Run Selected APIs", command=self.start_process)
        self.run_btn.grid(row=0, column=9, padx=5)

        self.cancel_btn = ttk.Button(setup_frame, text="Cancel", command=self.cancel_process, state="disabled")
        self.cancel_btn.grid(row=0, column=10, padx=3)

        # 2. Preview Frame: Request Details (Method, URL, Header, Body)
        preview_frame = ttk.LabelFrame(root, text=" Request Preview (Outgoing API Header & Raw/JSON Body) ", padding=5)
        preview_frame.pack(fill="x", padx=10, pady=5)

        self.preview_text = scrolledtext.ScrolledText(preview_frame, height=5, font=("Consolas", 9), wrap="word")
        self.preview_text.pack(fill="both", expand=True)
        self.update_request_preview("N/A", "N/A", {}, None)

        # 3. Middle Paned Window
        paned = ttk.PanedWindow(root, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=10, pady=5)

        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=3)

        # Selection Control Toolbar for Tick box
        btn_bar = ttk.Frame(left_frame)
        btn_bar.pack(fill="x", pady=2)
        ttk.Button(btn_bar, text="Select All", command=self.select_all_apis).pack(side="left", padx=2)
        ttk.Button(btn_bar, text="Deselect All", command=self.deselect_all_apis).pack(side="left", padx=2)
        ttk.Label(btn_bar, text=" (Click '[ ]' or '[☑]' cell to toggle selection)").pack(side="left", padx=5)

        # Separated Read & Write Columns + Checkbox Column
        cols = ("select", "sr", "name", "method", "url", "read_acc", "write_acc", "status")
        self.tree = ttk.Treeview(left_frame, columns=cols, show="headings", selectmode="browse")
        
        self.tree.heading("select", text="Run?")
        self.tree.heading("sr", text="Sr.")
        self.tree.heading("name", text="API Name")
        self.tree.heading("method", text="Method")
        self.tree.heading("url", text="URL")
        self.tree.heading("read_acc", text="Read Access")
        self.tree.heading("write_acc", text="Write Access")
        self.tree.heading("status", text="Status")

        self.tree.column("select", width=45, anchor="center")
        self.tree.column("sr", width=35, anchor="center")
        self.tree.column("name", width=210)
        self.tree.column("method", width=60, anchor="center")
        self.tree.column("url", width=220)
        self.tree.column("read_acc", width=85, anchor="center")
        self.tree.column("write_acc", width=85, anchor="center")
        self.tree.column("status", width=120, anchor="center")

        tree_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=tree_scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")
        
        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<<TreeviewSelect>>", self.on_select_tree_item)

        right_frame = ttk.LabelFrame(paned, text=" Received API Response Output ", padding=5)
        paned.add(right_frame, weight=2)

        self.output_text = scrolledtext.ScrolledText(right_frame, font=("Consolas", 9), wrap="word")
        self.output_text.pack(fill="both", expand=True)

        # 4. Bottom Status & Export Buttons
        bottom_frame = ttk.Frame(root, padding=5)
        bottom_frame.pack(fill="x", padx=10, pady=5)

        self.status_lbl = ttk.Label(bottom_frame, text="Status: Ready", font=("Arial", 9, "bold"))
        self.status_lbl.pack(side="left", padx=5)

        # Live 1-minute countdown label
        self.timer_lbl = ttk.Label(bottom_frame, text="Token Timer: 60s", font=("Arial", 9, "bold"), foreground="blue")
        self.timer_lbl.pack(side="right", padx=15)

        ttk.Button(bottom_frame, text="Export PDF", command=self.export_pdf).pack(side="right", padx=5)
        ttk.Button(bottom_frame, text="Export Excel", command=self.export_excel).pack(side="right", padx=5)

        self.populate_table()

    def start_1min_timer(self):
        self.session_expire_time = time.time() + 60
        if not self.timer_running:
            self.timer_running = True
            self.update_timer_loop()

    def update_timer_loop(self):
        if not self.timer_running:
            return

        if self.session_expire_time:
            remaining = int(self.session_expire_time - time.time())
            if remaining > 0:
                self.timer_lbl.config(text=f"Token Timer: {remaining}s remaining", foreground="blue")
                self.root.after(1000, self.update_timer_loop)
            else:
                self.timer_lbl.config(text="Token Timer: EXPIRED (0s)", foreground="red")
                self.update_status("Token Expired (1-min limit reached). Testing stopped! Fetch a new token to continue.")
                
                # Halt active execution pipeline
                self.cancel_event.set()
                self.run_btn.config(state="normal")
                self.cancel_btn.config(state="disabled")
                self.timer_running = False
        else:
            self.timer_lbl.config(text="Token Timer: Off", foreground="gray")

    def on_role_change(self, event=None):
        role = self.role_var.get()
        self.password_var.set(USER_CREDENTIALS[role])
        self.populate_table()

    def populate_table(self):
        self.tree.delete(*self.tree.get_children())
        selected_role = self.role_var.get()

        for api in API_ENDPOINTS:
            permissions = api["roles"].get(selected_role, [])
            read_access = "YES" if "Read" in permissions else "-"
            write_access = "YES" if "Write" in permissions else "-"

            if permissions:
                status = "Pending"
                tick = "☑"
            else:
                status = "Unauthorized"
                tick = "☐"

            self.tree.insert("", "end", iid=str(api["sr"]), values=(
                tick, api["sr"], api["name"], api["method"], api["url"], read_access, write_access, status
            ))

    def on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == "#1":  # Tickbox Column
                item_id = self.tree.identify_row(event.y)
                if item_id:
                    vals = list(self.tree.item(item_id, "values"))
                    vals[0] = "☑" if vals[0] == "☐" else "☐"
                    self.tree.item(item_id, values=vals)

    def select_all_apis(self):
        for item_id in self.tree.get_children():
            vals = list(self.tree.item(item_id, "values"))
            vals[0] = "☑"
            self.tree.item(item_id, values=vals)

    def deselect_all_apis(self):
        for item_id in self.tree.get_children():
            vals = list(self.tree.item(item_id, "values"))
            vals[0] = "☐"
            self.tree.item(item_id, values=vals)

    def update_request_preview(self, method, url, headers, payload):
        self.preview_text.config(state="normal")
        self.preview_text.delete("1.0", tk.END)
        preview_str = f"TARGET: [{method}] {url}\n"
        preview_str += f"HEADERS: {json.dumps(headers, indent=2)}\n"
        
        if payload is not None:
            if isinstance(payload, dict):
                preview_str += f"BODY (JSON):\n{json.dumps(payload, indent=2)}"
            else:
                preview_str += f"BODY (RAW TEXT):\n{str(payload)}"
        else:
            preview_str += "BODY: None"

        self.preview_text.insert(tk.END, preview_str)
        self.preview_text.config(state="normal")

    def fetch_session_token(self):
        self.update_status("Fetching Session Token (/api/login)...")
        def _thread():
            ap_username = self.ap_name_var.get().strip()
            password = self.password_var.get().strip()
            login_url = "http://192.168.4.1/api/login"
            payload = {"username": ap_username, "password": password}
            headers = {"Content-Type": "application/json"}

            self.root.after(0, lambda: self.update_request_preview("POST", login_url, headers, payload))
            try:
                resp = requests.post(login_url, json=payload, headers=headers, timeout=5)
                if resp.status_code == 200:
                    token = ""
                    try:
                        data = resp.json()
                        token = data.get("sessionToken") or data.get("session") or ""
                    except Exception:
                        pass

                    if not token:
                        token = resp.cookies.get("session") or "3fa9a12da510c17d"

                    self.session_token_var.set(token)
                    self.update_status(f"Login Success! Session Token: {token}")
                    
                    # Start / Reset the 1-minute countdown timer
                    self.root.after(0, self.start_1min_timer)
                else:
                    self.update_status("Login Failed. Check credentials.")
            except Exception as ex:
                self.update_status(f"Login Error: {str(ex)}")

        threading.Thread(target=_thread, daemon=True).start()

    def on_select_tree_item(self, event):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        sr_id = int(selected_item[0])

        # 1. Update Top Request Preview Box
        for api in API_ENDPOINTS:
            if api["sr"] == sr_id:
                token = self.session_token_var.get().strip()
                headers = {"Cookie": f"session={token}"} if token else {}
                
                payload = api.get("payload")
                if api["name"] == "Authentication API (Write API Login)":
                    payload = {"username": self.ap_name_var.get().strip(), "password": self.password_var.get().strip()}

                self.update_request_preview(api["method"], api["url"], headers, payload)
                break

        # 2. Update Right Response Output Box from execution history
        found_result = False
        for res in self.execution_results:
            if res["sr"] == sr_id:
                self.display_output(res)
                found_result = True
                break

        if not found_result:
            self.output_text.config(state="normal")
            self.output_text.delete("1.0", tk.END)
            self.output_text.insert(tk.END, f"API Sr. #{sr_id} selected.\nNo response received yet. Click 'Run Selected APIs' to execute.")

    def start_process(self):
        token = self.session_token_var.get().strip()
        if not token:
            messagebox.showwarning("Missing Token", "Please fetch or enter a session token first.")
            return

        # Check token expiration
        if self.session_expire_time and time.time() >= self.session_expire_time:
            messagebox.showwarning("Token Expired", "1-minute session timer has expired! Please fetch a new token before testing.")
            return

        selected_sr_list = []
        for item_id in self.tree.get_children():
            vals = self.tree.item(item_id, "values")
            if vals[0] == "☑":
                selected_sr_list.append(int(vals[1]))

        if not selected_sr_list:
            messagebox.showwarning("No Selection", "Please check at least one API tick box to run.")
            return

        self.cancel_event.clear()
        self.run_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.execution_results.clear()
        
        threading.Thread(target=self.run_execution_pipeline, args=(selected_sr_list,), daemon=True).start()

    def cancel_process(self):
        self.cancel_event.set()
        self.update_status("Cancelling execution pipeline...")
        self.cancel_btn.config(state="disabled")

    def run_execution_pipeline(self, selected_sr_list):
        role = self.role_var.get()
        apis_to_run = [a for a in API_ENDPOINTS if a["sr"] in selected_sr_list]

        for idx, api in enumerate(apis_to_run):
            if self.cancel_event.is_set():
                break

            # Check 1-minute expiration before sending every API request
            if self.session_expire_time and time.time() >= self.session_expire_time:
                self.update_status("Token 1-minute timer expired! Execution stopped. Waiting for a new token.")
                break

            token = self.session_token_var.get().strip()
            sr_id = str(api["sr"])

            permissions = api["roles"].get(role, [])
            read_acc = "YES" if "Read" in permissions else "-"
            write_acc = "YES" if "Write" in permissions else "-"

            self.update_status(f"Running API {idx+1}/{len(apis_to_run)}: {api['name']}...")
            self.tree.item(sr_id, values=("☑", api["sr"], api["name"], api["method"], api["url"], read_acc, write_acc, "Running..."))

            payload = api.get("payload")
            if api["name"] == "Authentication API (Write API Login)":
                payload = {"username": self.ap_name_var.get().strip(), "password": self.password_var.get().strip()}

            if isinstance(payload, dict):
                headers = {"Cookie": f"session={token}", "Content-Type": "application/json"}
            else:
                headers = {"Cookie": f"session={token}", "Content-Type": "text/plain"}

            self.root.after(0, lambda a=api, p=payload, h=headers: self.update_request_preview(a["method"], a["url"], h, p))

            body_output = ""
            status_str = ""

            try:
                if api["method"] == "GET":
                    r = requests.get(api["url"], headers=headers, timeout=5)
                else:
                    if isinstance(payload, dict):
                        r = requests.post(api["url"], json=payload, headers=headers, timeout=5)
                    else:
                        r = requests.post(api["url"], data=payload, headers=headers, timeout=5)

                status_str = f"{r.status_code} {r.reason}"
                body_output = r.text
            except Exception as ex:
                status_str = "Error/Offline"
                body_output = f"Connection Error: {str(ex)}"

            self.tree.item(sr_id, values=("☑", api["sr"], api["name"], api["method"], api["url"], read_acc, write_acc, status_str))

            res_entry = {
                "sr": api["sr"],
                "name": api["name"],
                "method": api["method"],
                "url": api["url"],
                "read_acc": read_acc,
                "write_acc": write_acc,
                "status": status_str,
                "payload": payload,
                "body": body_output
            }
            self.execution_results.append(res_entry)

            # Instantly display output in the right panel
            self.root.after(0, lambda res=res_entry: self.display_output(res))

            # Delay before next request
            if idx < len(apis_to_run) - 1:
                for remaining in range(10, 0, -1):
                    if self.cancel_event.is_set() or (self.session_expire_time and time.time() >= self.session_expire_time):
                        break
                    self.update_status(f"Waiting 10s delay... ({remaining}s remaining before next API)")
                    time.sleep(1)

        if not self.cancel_event.is_set() and (not self.session_expire_time or time.time() < self.session_expire_time):
            self.update_status("Selected Execution Pipeline Completed!")

        self.root.after(0, lambda: self.run_btn.config(state="normal"))
        self.root.after(0, lambda: self.cancel_btn.config(state="disabled"))

    def display_output(self, res):
        self.output_text.config(state="normal")
        self.output_text.delete("1.0", tk.END)
        output_data = (
            f"Device Name : {self.ap_name_var.get()}\n"
            f"API Name    : {res['name']}\n"
            f"Target URL  : {res['url']}\n"
            f"Read Access : {res['read_acc']} | Write Access: {res['write_acc']}\n"
            f"Status Code : {res['status']}\n\n"
            f"================ RESPONSE BODY ================\n"
            f"{res['body']}"
        )
        self.output_text.insert(tk.END, output_data)
        self.output_text.see(tk.END)

    def update_status(self, text):
        self.root.after(0, lambda: self.status_lbl.config(text=f"Status: {text}"))

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
        ws.append(["Sr.", "API Name", "Method", "URL", "Read Access", "Write Access", "Status", "Response Output"])

        for item in self.execution_results:
            ws.append([item["sr"], item["name"], item["method"], item["url"], item["read_acc"], item["write_acc"], item["status"], str(item["body"])[:500]])

        filename = f"{self.ap_name_var.get()}_API_Report.xlsx"
        wb.save(filename)
        messagebox.showinfo("Export Successful", f"Excel report saved as {filename}")

    def export_pdf(self):
        if not self.execution_results:
            messagebox.showinfo("Export Error", "No execution results available to export.")
            return

        filename = f"{self.ap_name_var.get()}_API_Report.pdf"
        doc = SimpleDocTemplate(filename, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []

        title_style = ParagraphStyle("TitleStyle", parent=styles["Heading1"], fontSize=16, leading=20)
        story.append(Paragraph("RMS Device API Test Summary Report", title_style))
        story.append(Spacer(1, 10))

        info_text = f"<b>Device Name:</b> {self.ap_name_var.get()}<br/>" \
                    f"<b>User Role:</b> {self.role_var.get()}<br/>" \
                    f"<b>Session Token:</b> {self.session_token_var.get()}"
        story.append(Paragraph(info_text, styles["Normal"]))
        story.append(Spacer(1, 15))

        table_data = [["Sr.", "API Name", "Method", "Read", "Write", "Status"]]
        for item in self.execution_results:
            table_data.append([str(item["sr"]), item["name"], item["method"], item["read_acc"], item["write_acc"], item["status"]])

        t = Table(table_data, colWidths=[30, 220, 50, 50, 50, 100])
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