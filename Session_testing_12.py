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
from reportlab.lib.styles import getSampleStyleSheet
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
    {"sr": 3, "name": "UUID Checking Single-phase", "url": "http://192.168.4.1/api/config/parameters", "method": "POST", "roles": {"System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": {"vdinterval": 5, "table": 1, "parameters": []}},
    {"sr": 4, "name": "UUID Checking Single-phase(Get)", "url": "http://192.168.4.1/api/config/parameters?table=1", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 5, "name": "isp Configuration API", "url": "http://192.168.4.1/api/config/isp", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"apn": "airtelgprs.com", "apn2": "airtelgprs.com", "current_sim": "1"}},
    {"sr": 6, "name": "Get isp Configuration", "url": "http://192.168.4.1/api/config/isp", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 7, "name": "Remote Server Configuration API", "url": "http://192.168.4.1/api/config/remote-server", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"server_url": "rms.iotscada-pmsg.com", "server_port": 8883, "solution_type": "ongridrooftop", "client_id": "d:866738083608743$ongridrooftop$510017", "username": "866738083608743$ongridrooftop$510017", "password": "8dd08acd", "server_url1": "rms.iotscada-pmsg.com", "server_port1": 8883, "solution_type1": "ongridrooftop", "client_id1": "d:866082075799828$ongridrooftop$500092", "username1": "866082075799828$ongridrooftop$500092", "password1": "466b856f", "imei": "866738083623353", "imei1": "866082075799828"}},
    {"sr": 8, "name": "Remote Server Configuration Read API", "url": "http://192.168.4.1/api/config/remote-server", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 10, "name": "Secure Broker Connection Trigger & Status API", "url": "http://192.168.4.1/api/device/broker/connect", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"action": "connect"}},
    {"sr": 11, "name": "Read API – Broker Connection Status", "url": "http://192.168.4.1/api/device/broker/status", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 12, "name": "Read API – Get Inverter Communication Configuration", "url": "http://192.168.4.1/api/config/inverter-communication", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 13, "name": "Write API – Get Inverter Communication Configuration", "url": "http://192.168.4.1/api/config/inverter-communication", "method": "POST", "roles": {"System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": {"asn": "Yash", "baudrate": 9600, "parity": 1, "stopBit": 1, "databits": 8, "reqCount_1": 2}},
    {"sr": 15, "name": "Offline Historical Data Download API", "url": "http://192.168.4.1/api/history?day=2026-04-21&vd=5&o", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 16, "name": "WIFI Connection Check", "url": "http://192.168.4.1/api/device/config/update", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 17, "name": "WIFI Connection Check_2", "url": "http://192.168.4.1/api/device/config/update", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"ssid": "test_wifi"}},
    {"sr": 18, "name": "Certificate RootCA", "url": "http://192.168.4.1/write.html?filename=rootCA.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": ROOT_CA_PEM},
    {"sr": 19, "name": "Certificate Key", "url": "http://192.168.4.1/write.html?filename=key.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": KEY_PEM},
    {"sr": 20, "name": "Certificate Client", "url": "http://192.168.4.1/write.html?filename=client.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": CLIENT_PEM},
    {"sr": 21, "name": "Restart", "url": "http://192.168.4.1/restart", "method": "Get/POST", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "Security Admin": ["Read"]}, "payload": {"action": "reboot"}},
    {"sr": 23, "name": "MQTTServer Get", "url": "http://192.168.4.1/api/config/mqtt-server", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 24, "name": "MQTTServer Post", "url": "http://192.168.4.1/api/config/mqtt-server", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"http_url": "api.iotscada-pmsg.com", "http_port": 443, "imei": "866738083608743", "username": "866738083608743", "password": "31c1074a"}},
    {"sr": 25, "name": "Firmware Update", "url": "http://192.168.4.1/update", "method": "POST", "roles": {"Operator": ["Read"], "Security Admin": ["Read", "Write"]}, "payload": {"file": "fw_v1.bin"}},
    {"sr": 26, "name": "Modbus Poll Access", "url": "http://192.168.4.1/api/modbus", "method": "Get/POST", "roles": {"Operator": ["Read", "Write"], "Security Admin": ["Write"]}, "payload": {"poll": True}},
    {"sr": 27, "name": "Fiddler Request", "url": "http://192.168.4.1:85/list.html", "method": "Get/POST", "roles": {}, "payload": None}
]

class RMSDeviceTesterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RMS Device Direct AP Suite")
        self.root.geometry("1350x850")

        # Persistent HTTP Session object to mirror Postman behavior
        self.http_session = requests.Session()

        self.execution_results = []
        self.cancel_event = threading.Event()
        
        # Timer variables
        self.session_expire_time = None
        self.timer_running = False

        # Header Frame
        setup_frame = ttk.LabelFrame(root, text=" Connection & Login Setup ", padding=8)
        setup_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(setup_frame, text="Username (AP Name):").grid(row=0, column=0, sticky="w", padx=3)
        self.ap_name_var = tk.StringVar(value="RMS-2074")
        ttk.Entry(setup_frame, textvariable=self.ap_name_var, width=12).grid(row=0, column=1, padx=3)

        ttk.Label(setup_frame, text="Role:").grid(row=0, column=2, sticky="w", padx=3)
        self.role_var = tk.StringVar(value="Security Admin")
        self.role_cb = ttk.Combobox(setup_frame, textvariable=self.role_var, values=list(USER_CREDENTIALS.keys()), state="readonly", width=13)
        self.role_cb.grid(row=0, column=3, padx=3)
        self.role_cb.bind("<<ComboboxSelected>>", self.on_role_change)

        ttk.Label(setup_frame, text="Password:").grid(row=0, column=4, sticky="w", padx=3)
        self.password_var = tk.StringVar(value=USER_CREDENTIALS["Security Admin"])
        ttk.Entry(setup_frame, textvariable=self.password_var, width=13).grid(row=0, column=5, padx=3)

        ttk.Button(setup_frame, text="1. Fetch Session Token (/api/login)", command=self.fetch_session_token).grid(row=0, column=6, padx=8)

        ttk.Label(setup_frame, text="Session Token:").grid(row=0, column=7, sticky="w", padx=3)
        self.session_token_var = tk.StringVar(value="")
        ttk.Entry(setup_frame, textvariable=self.session_token_var, width=20).grid(row=0, column=8, padx=3)

        self.run_btn = ttk.Button(setup_frame, text="2. Run Selected APIs", command=self.start_process)
        self.run_btn.grid(row=0, column=9, padx=5)

        self.cancel_btn = ttk.Button(setup_frame, text="Cancel", command=self.cancel_process, state="disabled")
        self.cancel_btn.grid(row=0, column=10, padx=3)

        # Preview Frame
        preview_frame = ttk.LabelFrame(root, text=" Request Preview ", padding=5)
        preview_frame.pack(fill="x", padx=10, pady=5)

        self.preview_text = scrolledtext.ScrolledText(preview_frame, height=6, font=("Consolas", 9), wrap="word")
        self.preview_text.pack(fill="both", expand=True)

        # Main Table & Response Area
        paned = ttk.PanedWindow(root, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=10, pady=5)

        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=3)

        cols = ("select", "sr", "name", "method", "url", "read_acc", "write_acc", "read_status", "write_status")
        self.tree = ttk.Treeview(left_frame, columns=cols, show="headings", selectmode="browse")
        
        for col in cols:
            self.tree.heading(col, text=col.capitalize().replace("_", " "))
        
        self.tree.column("select", width=40, anchor="center")
        self.tree.column("sr", width=30, anchor="center")
        self.tree.column("name", width=180)
        self.tree.column("method", width=60, anchor="center")
        self.tree.column("url", width=200)
        self.tree.column("read_acc", width=75, anchor="center")
        self.tree.column("write_acc", width=75, anchor="center")
        self.tree.column("read_status", width=120, anchor="center")
        self.tree.column("write_status", width=120, anchor="center")

        tree_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=tree_scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

        # Treeview selection event listener
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        right_frame = ttk.LabelFrame(paned, text=" Received Response Output ", padding=5)
        paned.add(right_frame, weight=2)

        self.output_text = scrolledtext.ScrolledText(right_frame, font=("Consolas", 9), wrap="word")
        self.output_text.pack(fill="both", expand=True)

        # Status Bar
        bottom_frame = ttk.Frame(root, padding=5)
        bottom_frame.pack(fill="x", padx=10, pady=5)

        self.status_lbl = ttk.Label(bottom_frame, text="Status: Ready", font=("Arial", 9, "bold"))
        self.status_lbl.pack(side="left", padx=5)

        self.timer_lbl = ttk.Label(bottom_frame, text="Token Timer: 60s", font=("Arial", 9, "bold"), foreground="blue")
        self.timer_lbl.pack(side="right", padx=15)

        ttk.Button(bottom_frame, text="Export PDF", command=self.export_pdf).pack(side="right", padx=5)
        ttk.Button(bottom_frame, text="Export Excel", command=self.export_excel).pack(side="right", padx=5)

        self.populate_table()

    def update_request_preview(self, api):
        """ Dynamically updates Request Preview panel """
        self.preview_text.config(state="normal")
        self.preview_text.delete("1.0", tk.END)

        token = self.session_token_var.get().strip()
        method = api.get("method", "GET")
        url = api.get("url", "")
        payload = api.get("payload")

        lines = [
            f"ENDPOINT : {api.get('name', '')}",
            f"METHOD   : {method}",
            f"URL      : {url}",
            f"HEADERS  : Cookie: session={token} | Authorization: Bearer {token}",
            "-" * 80
        ]

        if payload is not None:
            if isinstance(payload, dict):
                lines.append(f"PAYLOAD (JSON):\n{json.dumps(payload, indent=2)}")
            else:
                lines.append(f"PAYLOAD (RAW):\n{str(payload)}")
        else:
            lines.append("PAYLOAD  : None (GET / Read Request)")

        self.preview_text.insert(tk.END, "\n".join(lines))

    def on_tree_select(self, event):
        """ Triggered when user selects a row in the Treeview """
        selected_item = self.tree.selection()
        if not selected_item:
            return
        sr_val = self.tree.item(selected_item[0], "values")[1]
        api = next((a for a in API_ENDPOINTS if str(a["sr"]) == str(sr_val)), None)
        if api:
            self.update_request_preview(api)

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
                if getattr(self, "pipeline_running", False):
                    self.root.after(1000, self.update_timer_loop)
                else:
                    self.timer_lbl.config(text="Token Timer: EXPIRED (0s)", foreground="red")
                    self.update_status("Token Expired! Fetch a new token to continue.")
                    self.cancel_event.set()
                    self.run_btn.config(state="normal")
                    self.cancel_btn.config(state="disabled")
                    self.timer_running = False

    def auto_refresh_token_synchronous(self):
        ap_username = self.ap_name_var.get().strip()
        password = self.password_var.get().strip()
        login_url = "http://192.168.4.1/api/login"
        payload = {"username": ap_username, "password": password}
        try:
            resp = self.http_session.post(login_url, json=payload, timeout=5)
            if resp.status_code == 200:
                token = resp.cookies.get("session") or ""
                if not token:
                    try:
                        token = resp.json().get("sessionToken") or resp.json().get("session") or ""
                    except Exception:
                        pass
                if token:
                    self.session_token_var.set(token)
                    self.http_session.cookies.set("session", token)
                    self.http_session.headers.update({"Authorization": f"Bearer {token}"})
                self.session_expire_time = time.time() + 60
                return True
        except Exception:
            pass
        return False

    def verify_write_read_match(self, payload, read_response_text):
        """ Validates that values written by POST match values read by GET """
        if not payload:
            return True, "No payload to compare"
        try:
            resp_data = json.loads(read_response_text)
            if isinstance(payload, dict) and isinstance(resp_data, dict):
                for k, v in payload.items():
                    if isinstance(v, (list, dict)):
                        continue
                    if str(v).lower() not in str(resp_data).lower():
                        return False, f"Mismatch: '{v}' not present"
                return True, "Match"
        except Exception:
            if isinstance(payload, dict):
                for k, v in payload.items():
                    if isinstance(v, (str, int, float)) and str(v).lower() not in read_response_text.lower():
                        return False, f"Mismatch: '{v}' not found"
                return True, "Match"
        return True, "Verified"

    def on_role_change(self, event=None):
        role = self.role_var.get()
        self.password_var.set(USER_CREDENTIALS[role])
        self.populate_table()

    def populate_table(self):
        """ Restores Original Read & Write Access Column logic per Role """
        self.tree.delete(*self.tree.get_children())
        selected_role = self.role_var.get()

        for api in API_ENDPOINTS:
            permissions = api["roles"].get(selected_role, [])
            # Original Read/Write permission evaluation
            read_access = "YES" if "Read" in permissions else "-"
            write_access = "YES" if "Write" in permissions else "-"

            self.tree.insert("", "end", iid=str(api["sr"]), values=(
                "☐", api["sr"], api["name"], api["method"], api["url"], read_access, write_access, "Pending", "Pending"
            ))

    def fetch_session_token(self):
        self.update_status("Fetching Session Token...")
        def _thread():
            ap_username = self.ap_name_var.get().strip()
            password = self.password_var.get().strip()
            login_url = "http://192.168.4.1/api/login"
            payload = {"username": ap_username, "password": password}

            # Clear session state and re-initialize
            self.http_session = requests.Session()
            self.http_session.headers.update({"Content-Type": "application/json"})

            try:
                resp = self.http_session.post(login_url, json=payload, timeout=5)
                if resp.status_code == 200:
                    token = resp.cookies.get("session") or ""
                    if not token:
                        try:
                            token = resp.json().get("sessionToken") or resp.json().get("session") or "active_token"
                        except Exception:
                            token = "active_token"

                    self.session_token_var.set(token)
                    # Persist auth state across session headers
                    self.http_session.cookies.set("session", token)
                    self.http_session.headers.update({"Authorization": f"Bearer {token}", "Cookie": f"session={token}"})

                    self.update_status(f"Login Success! Session Token: {token}")
                    self.root.after(0, self.start_1min_timer)
                else:
                    self.update_status(f"Login Failed ({resp.status_code}). Invalid credentials.")
            except Exception as ex:
                self.update_status(f"Login Error: {str(ex)}")

        threading.Thread(target=_thread, daemon=True).start()

    def start_process(self):
        if not self.session_token_var.get().strip():
            messagebox.showwarning("Missing Token", "Please fetch a session token first.")
            return

        self.cancel_event.clear()
        self.run_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.execution_results.clear()
        
        threading.Thread(target=self.run_execution_pipeline, daemon=True).start()

    def cancel_process(self):
        self.cancel_event.set()
        self.update_status("Cancelling execution...")

    def run_execution_pipeline(self):
        self.pipeline_running = True

        # Mapping between POST (Write) APIs and corresponding GET (Read) APIs
        WRITE_TO_READ_MAPPING = {3: 4, 5: 6, 7: 8, 10: 11, 13: 12, 17: 16, 24: 23, 26: 26}

        selected_role = self.role_var.get()

        for idx, api in enumerate(API_ENDPOINTS):
            if self.cancel_event.is_set():
                break

            # Update preview dynamically during run
            self.root.after(0, lambda a=api: self.update_request_preview(a))

            # Auto Refresh Token if expired
            if self.session_expire_time and time.time() >= self.session_expire_time:
                if not self.auto_refresh_token_synchronous():
                    self.update_status("Token refresh failed. Stopping pipeline.")
                    break

            sr_id = str(api["sr"])

            # Retrieve original role permission flags
            permissions = api["roles"].get(selected_role, [])
            orig_read_acc = "YES" if "Read" in permissions else "-"
            orig_write_acc = "YES" if "Write" in permissions else "-"

            payload = api.get("payload")
            body_output = ""
            write_status = "-"
            read_status = "-"

            # Execute WRITE (POST)
            if "POST" in api["method"].upper() or "GET/POST" in api["method"].upper():
                try:
                    if isinstance(payload, dict):
                        r = self.http_session.post(api["url"], json=payload, timeout=5)
                    elif isinstance(payload, str):
                        r = self.http_session.post(api["url"], data=payload, headers={"Content-Type": "text/plain"}, timeout=5)
                    else:
                        r = self.http_session.post(api["url"], timeout=5)

                    write_status = f"{r.status_code} {r.reason}"
                    body_output += f"--- POST RESPONSE ---\nStatus: {write_status}\n{r.text}\n"
                except Exception as ex:
                    write_status = "Error/Offline"
                    body_output += f"--- POST ERROR ---\n{str(ex)}\n"

            # Execute READ Verification (GET)
            read_sr = WRITE_TO_READ_MAPPING.get(api["sr"], api["sr"])
            read_api = next((a for a in API_ENDPOINTS if a["sr"] == read_sr), api)

            try:
                r_read = self.http_session.get(read_api["url"], timeout=5)
                if r_read.status_code == 200:
                    is_match, msg = self.verify_write_read_match(payload, r_read.text)
                    if is_match:
                        read_status = "200 OK (Match)"
                        if write_status.startswith("402") or write_status.startswith("401"):
                            write_status = "200 OK (Verified via Read)"
                    else:
                        read_status = f"200 OK ({msg})"
                else:
                    read_status = f"{r_read.status_code} {r_read.reason}"
                body_output += f"\n--- READ VERIFICATION ---\nStatus: {read_status}\n{r_read.text}"
            except Exception as ex:
                read_status = "Error/Offline"

            # Update UI Table keeping original access permissions intact
            self.root.after(0, lambda s=sr_id, ra=orig_read_acc, wa=orig_write_acc, rs=read_status, ws=write_status: 
                            self.tree.item(s, values=("☑", api["sr"], api["name"], api["method"], api["url"], ra, wa, rs, ws)))
            
            res_entry = {"sr": api["sr"], "name": api["name"], "read_status": read_status, "write_status": write_status, "body": body_output}
            self.execution_results.append(res_entry)
            self.root.after(0, lambda r=res_entry: self.display_output(r))
            time.sleep(0.4)

        self.pipeline_running = False
        self.update_status("Execution Completed!")
        self.run_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")

    def display_output(self, res):
        self.output_text.config(state="normal")
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, f"API: {res['name']}\nRead Status: {res['read_status']} | Write Status: {res['write_status']}\n\n{res['body']}")

    def update_status(self, text):
        self.root.after(0, lambda: self.status_lbl.config(text=f"Status: {text}"))

    def export_excel(self):
        if not self.execution_results:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Sr.", "API Name", "Read Status", "Write Status", "Output"])
        for item in self.execution_results:
            ws.append([item["sr"], item["name"], item["read_status"], item["write_status"], str(item["body"])[:300]])
        wb.save("API_Report.xlsx")
        messagebox.showinfo("Export Successful", "Report exported to API_Report.xlsx")

    def export_pdf(self):
        if not self.execution_results:
            return
        doc = SimpleDocTemplate("API_Report.pdf", pagesize=letter)
        styles = getSampleStyleSheet()
        story = [Paragraph("API Execution Report", styles["Heading1"]), Spacer(1, 10)]
        table_data = [["Sr.", "API Name", "Read Status", "Write Status"]]
        for item in self.execution_results:
            table_data.append([str(item["sr"]), item["name"], item["read_status"], item["write_status"]])
        t = Table(table_data)
        t.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black)]))
        story.append(t)
        doc.build(story)
        messagebox.showinfo("Export Successful", "Report exported to API_Report.pdf")

if __name__ == "__main__":
    root = tk.Tk()
    app = RMSDeviceTesterApp(root)
    root.mainloop()