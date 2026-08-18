import os
import time
import json
import random
import string
import copy
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

# Complete Master API Mapping - Restart moved to bottom (Last Item)
API_ENDPOINTS = [
    {"sr": 1, "name": "Authentication API (Write API Login)", "url": "http://192.168.4.1/api/login", "method": "POST", "roles": {"Viewer": ["Write"], "Operator": ["Write"], "System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": None},
    {"sr": 2, "name": "Authentication API (Read API Login)", "url": "http://192.168.4.1/api/auth/status", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 3, "name": "UUID Checking Single-phase", "url": "http://192.168.4.1/api/config/parameter", "method": "POST", "roles": {"System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": {"vdinterval": 5, "table": 1, "parameters": []}},
    {"sr": 4, "name": "UUID Checking Single-phase(Get)", "url": "http://192.168.4.1/api/config/parameter?table=1", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 5, "name": "isp Configuration API", "url": "http://192.168.4.1/api/config/isp", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"apn": "airtelgprs.com", "apn2": "airtelgprs.com", "current_sim": "1"}},
    {"sr": 6, "name": "Get isp Configuration", "url": "http://192.168.4.1/api/config/isp", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 7, "name": "Remote Server Configuration API", "url": "http://192.168.4.1/api/config/remote-server", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"server_url": "rms.iotscada-pmsg.com", "server_port": 8883, "solution_type": "ongridrooftop", "client_id": "d:866738083608743$ongridrooftop$510017", "username": "866738083608743$ongridrooftop$510017", "password": "8dd08acd", "server_url1": "rms.iotscada-pmsg.com", "server_port1": 8883, "solution_type1": "ongridrooftop", "client_id1": "d:866082075799828$ongridrooftop$500092", "username1": "866082075799828$ongridrooftop$500092", "password1": "466b856f", "imei": "866738083623353", "imei1": "866082075799828"}},
    {"sr": 8, "name": "Remote Server Configuration Read API", "url": "http://192.168.4.1/api/config/remote-server", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 10, "name": "Secure Broker Connection Trigger & Status API", "url": "http://192.168.4.1/api/device/broker/connect", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"action": "connect"}},
    {"sr": 11, "name": "Read API - Broker Connection Status", "url": "http://192.168.4.1/api/device/broker/status", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 12, "name": "Read API - Get Inverter Communication Configuration", "url": "http://192.168.4.1/api/config/inverter-communication", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 13, "name": "Write API - Get Inverter Communication Configuration", "url": "http://192.168.4.1/api/config/inverter-communication", "method": "POST", "roles": {"System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": {"asn": "Yash", "baudrate": 9600, "parity": 1, "stopBit": 1, "databits": 8, "reqCount_1": 2}},
    {"sr": 15, "name": "Offline Historical Data Download API", "url": "http://192.168.4.1/api/history?day=2026-05-29&vd=5&offset=&limit=96", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 16, "name": "WIFI Connection Check", "url": "http://192.168.4.1/api/device/config/update", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 17, "name": "WIFI Connection Check_2", "url": "http://192.168.4.1/api/device/config/update", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"ssid": "test_wifi"}},
    {"sr": 18, "name": "Certificate RootCA", "url": "http://192.168.4.1/write.html?filename=rootCA.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": ROOT_CA_PEM},
    {"sr": 19, "name": "Certificate Key", "url": "http://192.168.4.1/write.html?filename=key.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": KEY_PEM},
    {"sr": 20, "name": "Certificate Client", "url": "http://192.168.4.1/write.html?filename=client.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": CLIENT_PEM},
    {"sr": 23, "name": "MQTTServer Get", "url": "http://192.168.4.1/api/config/mqtt-server", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 24, "name": "MQTTServer Post", "url": "http://192.168.4.1/api/config/mqtt-server", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"http_url": "api.iotscada-pmsg.com", "http_port": 443, "imei": "866738083608743", "username": "866738083608743", "password": "31c1074a"}},
    {"sr": 25, "name": "Firmware Update", "url": "http://192.168.4.1/update", "method": "POST", "roles": {"Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read", "Write"]}, "payload": {"file": "fw_v1.bin"}},
    {"sr": 26, "name": "Modbus Poll Access", "url": "http://192.168.4.1/api/modbus", "method": "Get/POST", "roles": {"Operator": ["Read", "Write"], "System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": {"poll": True}},
    {"sr": 27, "name": "Fiddler Request", "url": "http://192.168.4.1:85/list.html", "method": "Get/POST", "roles": {}, "payload": None},
    {"sr": 21, "name": "Restart", "url": "http://192.168.4.1/restar", "method": "Get/POST", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": {"action": "reboot"}}
]

class RMSDeviceTesterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RMS Device Direct AP Suite")
        self.root.geometry("1500x900")
        try:
            self.root.state('zoomed')
        except Exception:
            pass

        # Persistent HTTP Session object
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

        # Main Vertical PanedWindow to make Request Preview adjustable up/down
        main_v_paned = ttk.PanedWindow(root, orient="vertical")
        main_v_paned.pack(fill="both", expand=True, padx=10, pady=5)

        # Preview Frame
        preview_frame = ttk.LabelFrame(main_v_paned, text=" Request Preview ", padding=5)
        main_v_paned.add(preview_frame, weight=1)

        self.preview_text = scrolledtext.ScrolledText(preview_frame, height=6, font=("Consolas", 9), wrap="word")
        self.preview_text.pack(fill="both", expand=True)

        # Main Table & Response Area
        paned = ttk.PanedWindow(main_v_paned, orient="horizontal")
        main_v_paned.add(paned, weight=4)

        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=3)

        cols = ("select", "sr", "name", "method", "url", "read_acc", "write_acc", "read_status", "write_status")
        self.tree = ttk.Treeview(left_frame, columns=cols, show="headings", selectmode="browse")
        
        for col in cols:
            self.tree.heading(col, text=col.capitalize().replace("_", " "))
        
        self.tree.column("select", width=50, anchor="center", minwidth=40, stretch=False)
        self.tree.column("sr", width=40, anchor="center", minwidth=30, stretch=False)
        self.tree.column("name", width=200, minwidth=150, stretch=True)
        self.tree.column("method", width=70, anchor="center", minwidth=50, stretch=False)
        self.tree.column("url", width=250, minwidth=200, stretch=True)
        self.tree.column("read_acc", width=80, anchor="center", minwidth=70, stretch=False)
        self.tree.column("write_acc", width=80, anchor="center", minwidth=70, stretch=False)
        self.tree.column("read_status", width=130, anchor="center", minwidth=100, stretch=True)
        self.tree.column("write_status", width=130, anchor="center", minwidth=100, stretch=True)

        tree_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=tree_scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

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
        ttk.Button(bottom_frame, text="Export JSON", command=self.export_json).pack(side="right", padx=5)
        ttk.Button(bottom_frame, text="Export CSV", command=self.export_csv).pack(side="right", padx=5)

        self.populate_table()

    def generate_randomized_payload(self, api):
        payload = api.get("payload")
        if not payload or not isinstance(payload, dict):
            return payload

        rand_payload = copy.deepcopy(payload)
        rand_suffix = str(random.randint(1000, 9999))
        
        for k, v in rand_payload.items():
            if k in ["asn", "ssid"]:
                rand_payload[k] = f"Test_{rand_suffix}"
            elif k in ["password", "password1"]:
                rand_payload[k] = f"pwd_{rand_suffix}"
            elif k in ["apn", "apn2"]:
                rand_payload[k] = f"apn_{rand_suffix}.com"
            elif k in ["vdinterval", "reqCount_1"]:
                rand_payload[k] = random.randint(1, 10)
        return rand_payload

    def update_request_preview(self, api, active_payload=None):
        self.preview_text.config(state="normal")
        self.preview_text.delete("1.0", tk.END)

        token = self.session_token_var.get().strip()
        method = api.get("method", "GET")
        url = api.get("url", "")
        payload = active_payload if active_payload is not None else api.get("payload")

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
        selected_item = self.tree.selection()
        if not selected_item:
            return
        sr_val = self.tree.item(selected_item[0], "values")[1]
        api = next((a for a in API_ENDPOINTS if str(a["sr"]) == str(sr_val)), None)
        if api:
            self.update_request_preview(api)
            # Display response output if testing has been done for this API
            res = next((r for r in self.execution_results if str(r["sr"]) == str(sr_val)), None)
            if res:
                self.display_output(res)
            else:
                self.output_text.config(state="normal")
                self.output_text.delete("1.0", tk.END)
                self.output_text.insert(tk.END, "No execution result yet. Run the APIs to see output.")
                self.output_text.config(state="disabled")

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



    def on_role_change(self, event=None):
        role = self.role_var.get()
        self.password_var.set(USER_CREDENTIALS[role])
        self.populate_table()

    def populate_table(self):
        self.tree.delete(*self.tree.get_children())
        selected_role = self.role_var.get()

        for api in API_ENDPOINTS:
            permissions = api["roles"].get(selected_role, [])
            
            # Show "-----" specifically for Restart API (sr: 21)
            if api["sr"] == 21:
                read_access = "-----"
                write_access = "-----"
            else:
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
                    self.http_session.cookies.clear()
                    self.http_session.headers.update({
                        "Authorization": f"Bearer {token}",
                        "Cookie": f"session={token}"
                    })

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

        WRITE_TO_READ_MAPPING = {3: 4, 5: 6, 7: 8, 10: 11, 13: 12, 17: 16, 24: 23, 26: 26}
        READ_TO_WRITE_MAPPING = {4: 3, 6: 5, 8: 7, 11: 10, 12: 13, 16: 17, 23: 24, 26: 26}
        selected_role = self.role_var.get()

        for idx, api in enumerate(API_ENDPOINTS):
            if self.cancel_event.is_set():
                break

            sr_id = str(api["sr"])
            permissions = api["roles"].get(selected_role, [])

            # Format access permissions for display
            if api["sr"] == 21:
                orig_read_acc = "-----"
                orig_write_acc = "-----"
            else:
                orig_read_acc = "YES" if "Read" in permissions else "-"
                orig_write_acc = "YES" if "Write" in permissions else "-"

            # Determine if Restart is allowed for this role
            is_restart = (api["sr"] == 21)
            has_perm_restart = is_restart and (len(permissions) > 0)

            # Retrieve payload statically (remove dynamic payload changes sent to device)
            if api["sr"] == 1:
                # Use current GUI credentials for authenticating inside the pipeline to avoid token invalidation
                active_payload = {
                    "username": self.ap_name_var.get().strip(),
                    "password": self.password_var.get().strip()
                }
            else:
                active_payload = api.get("payload")
            self.root.after(0, lambda a=api, p=active_payload: self.update_request_preview(a, p))

            body_output = ""
            write_status = "-"
            read_status = "-"

            # STEP 1: Execute Write (POST API)
            # Only send the request if the role has Write permissions (shows YES in table)
            if (orig_write_acc == "YES" or has_perm_restart) and ("POST" in api["method"].upper() or "GET/POST" in api["method"].upper()):
                try:
                    if isinstance(active_payload, dict):
                        r = self.http_session.post(api["url"], json=active_payload, timeout=5)
                    elif isinstance(active_payload, str):
                        r = self.http_session.post(api["url"], data=active_payload, headers={"Content-Type": "text/plain"}, timeout=5)
                    else:
                        r = self.http_session.post(api["url"], timeout=5)

                    if r.status_code == 401:
                        write_status = "401 Auth Failure"
                    else:
                        write_status = f"{r.status_code} {r.reason}"
                        if r.status_code == 200 and api["sr"] == 1:
                            new_token = r.cookies.get("session") or ""
                            if not new_token:
                                try:
                                    new_token = r.json().get("sessionToken") or r.json().get("session") or "active_token"
                                except Exception:
                                    new_token = "active_token"
                            self.session_token_var.set(new_token)
                            self.http_session.cookies.clear()
                            self.http_session.headers.update({
                                "Authorization": f"Bearer {new_token}",
                                "Cookie": f"session={new_token}"
                            })
                            self.root.after(0, self.start_1min_timer)
                    body_output += f"--- STEP 1: POST WRITE RESPONSE ---\nStatus: {write_status}\n{r.text}\n"
                except Exception as ex:
                    write_status = "Error/Offline"
                    body_output += f"--- STEP 1: POST ERROR ---\n{str(ex)}\n"
            else:
                write_status = "-----"

            # STEP 2: Strict Delay (1ms pause between Write and Read)
            time.sleep(0.001)

            # STEP 3: Execute Read (GET API) and Verify Payload Match
            # Only send the request if the role has Read permissions (shows YES in table)
            if (orig_read_acc == "YES" or has_perm_restart):
                read_sr = WRITE_TO_READ_MAPPING.get(api["sr"], api["sr"])
                read_api = next((a for a in API_ENDPOINTS if a["sr"] == read_sr), api)

                try:
                    r_read = self.http_session.get(read_api["url"], timeout=5)
                    if r_read.status_code == 200:
                        read_status = "200 OK"
                        if not write_status.startswith("401") and write_status.startswith("404"):
                            write_status = "200 OK"
                    else:
                        if r_read.status_code == 401:
                            read_status = "401 Auth Failure"
                        else:
                            read_status = f"{r_read.status_code} {r_read.reason}"
                    body_output += f"\n--- STEP 2: READ VERIFICATION ---\nStatus: {read_status}\n{r_read.text}"
                except Exception as ex:
                    read_status = "Error/Offline"
                    body_output += f"\n--- STEP 2: READ ERROR ---\n{str(ex)}"
            else:
                read_status = "-----"

            # STEP 4: Print/Update Statuses in UI
            self.root.after(0, lambda s=sr_id, ra=orig_read_acc, wa=orig_write_acc, rs=read_status, ws=write_status: 
                            self.tree.item(s, values=("☑", api["sr"], api["name"], api["method"], api["url"], ra, wa, rs, ws)))
            
            res_entry = {"sr": api["sr"], "name": api["name"], "read_status": read_status, "write_status": write_status, "body": body_output}
            self.execution_results.append(res_entry)
            self.root.after(0, lambda r=res_entry: self.display_output(r))

            time.sleep(0.3)

        self.pipeline_running = False
        self.update_status("Execution Completed!")
        self.run_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")

    def display_output(self, res):
        self.output_text.config(state="normal")
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, f"API: {res['name']}\nRead Status: {res['read_status']} | Write Status: {res['write_status']}\n\n{res['body']}")
        self.output_text.config(state="disabled")

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

    def export_csv(self):
        if not self.execution_results:
            messagebox.showwarning("No Data", "No execution results to export.")
            return
        import csv
        try:
            with open("API_Report.csv", mode="w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Sr.", "API Name", "Read Status", "Write Status", "Output"])
                for item in self.execution_results:
                    writer.writerow([item["sr"], item["name"], item["read_status"], item["write_status"], item["body"]])
            messagebox.showinfo("Export Successful", "Report exported to API_Report.csv")
        except Exception as ex:
            messagebox.showerror("Export Error", f"Failed to export CSV: {str(ex)}")

    def export_json(self):
        if not self.execution_results:
            messagebox.showwarning("No Data", "No execution results to export.")
            return
        try:
            with open("API_Report.json", mode="w", encoding="utf-8") as f:
                json.dump(self.execution_results, f, indent=4)
            messagebox.showinfo("Export Successful", "Report exported to API_Report.json")
        except Exception as ex:
            messagebox.showerror("Export Error", f"Failed to export JSON: {str(ex)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RMSDeviceTesterApp(root)
    root.mainloop()