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
import html

# Excel export library
import openpyxl

# PDF export library
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Default Role Passwords (Updated to default to Admin@123)
USER_CREDENTIALS = {
    "Viewer": "Admin@123",
    "Operator": "Admin@123",
    "System Admin": "Admin@123",
    "Security Admin": "Admin@123"
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
98bK2QKBgQDrRsxyO4sbmVVGLQ1+WBdyF02YTvUorGqXWuJ2WZKs4MuSmb11287c
0JxVy+xAgXwQXEHKE7dLtOvR0rwQ+6Ki8hOgM+ebhiZjZOab09XaztgCvQC7akVf
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

# Complete Master API Mapping - Added Sr 9 (Metadata) and Sr 28 (Device Config Update)
API_ENDPOINTS = [
    {"sr": 1, "name": "Authentication API (Write API Login)", "url": "http://192.168.4.1/api/login", "method": "POST", "roles": {"Viewer": ["Write"], "Operator": ["Write"], "System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": None},
    {"sr": 2, "name": "Authentication API (Read API Login)", "url": "http://192.168.4.1/api/auth/status", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 3, "name": "UUID Checking Single-phase", "url": "http://192.168.4.1/api/config/parameter", "method": "POST", "roles": {"System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": {"vdinterval": 5, "table": 1, "parameters": []}},
    {"sr": 4, "name": "UUID Checking Single-phase(Get)", "url": "http://192.168.4.1/api/config/parameter?table=1", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 5, "name": "isp Configuration API", "url": "http://192.168.4.1/api/config/isp", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"apn": "airtelgprs.com", "apn2": "airtelgprs.com", "current_sim": "1"}},
    {"sr": 6, "name": "Get isp Configuration", "url": "http://192.168.4.1/api/config/isp", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 7, "name": "Remote Server Configuration API", "url": "http://192.168.4.1/api/config/remote-server", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"server_url": "rms.iotscada-pmsg.com", "server_port": 8883, "solution_type": "ongridrooftop", "client_id": "d:866738083608743$ongridrooftop$510017", "username": "866738083608743$ongridrooftop$510017", "password": "8dd08acd", "server_url1": "rms.iotscada-pmsg.com", "server_port1": 8883, "solution_type1": "ongridrooftop", "client_id1": "d:866082075799828$ongridrooftop$500092", "username1": "866082075799828$ongridrooftop$500092", "password1": "466b856f", "imei": "866738083623353", "imei1": "866082075799828"}},
    {"sr": 8, "name": "Remote Server Configuration Read API", "url": "http://192.168.4.1/api/config/remote-server", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 9, "name": "Device Metadata API", "url": "http://192.168.4.1/api/device/metadata", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 10, "name": "Secure Broker Connection Trigger & Status API", "url": "http://192.168.4.1/api/device/broker/connect", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"action": "connect"}},
    {"sr": 11, "name": "Read API - Broker Connection Status", "url": "http://192.168.4.1/api/device/broker/status", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 12, "name": "Read API - Get Inverter Communication Configuration", "url": "http://192.168.4.1/api/config/inverter-communication", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 13, "name": "Write API - Get Inverter Communication Configuration", "url": "http://192.168.4.1/api/config/inverter-communication", "method": "POST", "roles": {"System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": {"asn": "Yash", "baudrate": 9600, "parity": 1, "stopBit": 1, "databits": 8, "reqCount_1": 2}},
    {"sr": 15, "name": "Offline Historical Data Download API", "url": "http://192.168.4.1/api/history?day=2026-08-21&vd=5&offset=10&limit=96", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 16, "name": "WIFI Connection Check", "url": "http://192.168.4.1/api/device/config/update", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 17, "name": "WIFI Connection Check_2", "url": "http://192.168.4.1/api/device/config/update", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"ssid": "test_wifi"}},
    {"sr": 18, "name": "Certificate RootCA", "url": "http://192.168.4.1/write.html?filename=rootCA.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": ROOT_CA_PEM},
    {"sr": 19, "name": "Certificate Key", "url": "http://192.168.4.1/write.html?filename=key.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": KEY_PEM},
    {"sr": 20, "name": "Certificate Client", "url": "http://192.168.4.1/write.html?filename=client.pem", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": CLIENT_PEM},
    {"sr": 23, "name": "MQTTServer Get", "url": "http://192.168.4.1/api/config/mqtt-server", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 24, "name": "MQTTServer Post", "url": "http://192.168.4.1/api/config/mqtt-server", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"http_url": "api.iotscada-pmsg.com", "http_port": 443, "imei": "866738083608743", "username": "866738083608743", "password": "31c1074a"}},
    {"sr": 25, "name": "Firmware Update", "url": "http://192.168.4.1/update1", "method": "POST", "roles": {"Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read", "Write"]}, "payload": {"file": "fw_v1.bin"}},
    {"sr": 26, "name": "Modbus Poll Access", "url": "http://192.168.4.1/api/modbus", "method": "Get/POST", "roles": {"Operator": ["Read", "Write"], "System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": {"poll": True}},
    {"sr": 27, "name": "Fiddler Request", "url": "http://192.168.4.1:85/list.html", "method": "Get/POST", "roles": {}, "payload": None},
    {"sr": 28, "name": "Device Config Update API", "url": "http://192.168.4.1/api/device/config/update", "method": "POST", "roles": {"Viewer": ["Write"], "Operator": ["Write"], "System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": {"type": 2, "device_username": "SystemAdmin", "device_password": "Admin@123"}},
    {"sr": 21, "name": "Restart", "url": "http://192.168.4.1/restart1", "method": "Get/POST", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": {"action": "reboot"}}
]

class RMSDeviceTesterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RMS Device Direct AP Suite Pro (v14)")
        self.root.geometry("1550x950")
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

        # Apply modern styling
        self.setup_styles()

        # Header Frame
        setup_frame = ttk.LabelFrame(root, text=" Connection & Login Setup ", padding=8)
        setup_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(setup_frame, text="Username (AP Name):").grid(row=0, column=0, sticky="w", padx=3)
        self.ap_name_var = tk.StringVar(value="RMS-3E30")
        ttk.Entry(setup_frame, textvariable=self.ap_name_var, width=12).grid(row=0, column=1, padx=3)

        ttk.Label(setup_frame, text="Role:").grid(row=0, column=2, sticky="w", padx=3)
        self.role_var = tk.StringVar(value="Security Admin")
        self.role_cb = ttk.Combobox(setup_frame, textvariable=self.role_var, values=list(USER_CREDENTIALS.keys()), state="readonly", width=13)
        self.role_cb.grid(row=0, column=3, padx=3)
        self.role_cb.bind("<<ComboboxSelected>>", self.on_role_change)

        ttk.Label(setup_frame, text="Password:").grid(row=0, column=4, sticky="w", padx=3)
        self.password_var = tk.StringVar(value=USER_CREDENTIALS["Security Admin"])
        ttk.Entry(setup_frame, textvariable=self.password_var, width=13).grid(row=0, column=5, padx=3)

        ttk.Button(setup_frame, text="1. Fetch Session Token (3-Step Flow)", command=self.fetch_session_token).grid(row=0, column=6, padx=8)

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

        # Preview Frame (Contains Request Details - double click table row to edit)
        preview_frame = ttk.LabelFrame(main_v_paned, text=" Request Details (Double-Click table row to edit API/Payload) ", padding=5)
        main_v_paned.add(preview_frame, weight=2)

        self.preview_text = scrolledtext.ScrolledText(
            preview_frame, height=6, font=("Consolas", 9), wrap="word",
            bg=self.BG_CARD, fg=self.FG_TEXT, insertbackground=self.FG_TEXT,
            highlightthickness=1, highlightbackground=self.BG_BORDER, relief="flat"
        )
        self.preview_text.pack(fill="both", expand=True)

        # Main Table & Response Area
        paned = ttk.PanedWindow(main_v_paned, orient="horizontal")
        main_v_paned.add(paned, weight=5)

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
        self.tree.bind("<Double-1>", self.popup_edit_dialog)

        right_frame = ttk.LabelFrame(paned, text=" Received Response Output ", padding=5)
        paned.add(right_frame, weight=2)

        self.output_text = scrolledtext.ScrolledText(
            right_frame, font=("Consolas", 9), wrap="word",
            bg=self.BG_CARD, fg=self.FG_TEXT, insertbackground=self.FG_TEXT,
            highlightthickness=1, highlightbackground=self.BG_BORDER, relief="flat"
        )
        self.output_text.pack(fill="both", expand=True)

        # Status Bar
        bottom_frame = ttk.Frame(root, padding=5)
        bottom_frame.pack(fill="x", padx=10, pady=5)

        self.status_lbl = ttk.Label(bottom_frame, text="Status: Ready", font=("Segoe UI", 9, "bold"))
        self.status_lbl.pack(side="left", padx=5)

        self.timer_lbl = ttk.Label(bottom_frame, text="Token Timer: 60s", font=("Segoe UI", 9, "bold"), foreground=self.COLOR_BLUE)
        self.timer_lbl.pack(side="right", padx=15)

        ttk.Button(bottom_frame, text="Add More URL", command=self.popup_add_api_dialog).pack(side="right", padx=5)
        ttk.Button(bottom_frame, text="Export PDF", command=self.export_pdf).pack(side="right", padx=5)
        ttk.Button(bottom_frame, text="Export Excel", command=self.export_excel).pack(side="right", padx=5)
        ttk.Button(bottom_frame, text="Export JSON", command=self.export_json).pack(side="right", padx=5)
        ttk.Button(bottom_frame, text="Export CSV", command=self.export_csv).pack(side="right", padx=5)

        self.populate_table()

    def setup_styles(self):
        """Configure modern light theme with white/gray background for ttk widgets"""
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # Color codes
        self.BG_MAIN = "#f4f4f9"
        self.BG_CARD = "#ffffff"
        self.BG_BORDER = "#d1d5db"
        self.FG_TEXT = "#1f2937"
        self.FG_MUTED = "#6b7280"
        self.COLOR_ACCENT = "#7c3aed" # Deep Purple
        self.COLOR_BLUE = "#2563eb"   # Royal Blue
        self.COLOR_GREEN = "#16a34a"  # Green
        self.COLOR_YELLOW = "#ca8a04" # Yellow
        self.COLOR_RED = "#dc2626"    # Red

        # Tkinter main background color
        self.root.configure(bg=self.BG_MAIN)

        # Apply styles
        self.style.configure(".", background=self.BG_MAIN, foreground=self.FG_TEXT, font=("Segoe UI", 9), bordercolor=self.BG_BORDER)
        self.style.configure("TFrame", background=self.BG_MAIN)
        self.style.configure("Card.TFrame", background=self.BG_CARD, relief="solid", borderwidth=1)
        self.style.configure("TLabelframe", background=self.BG_MAIN, foreground=self.COLOR_ACCENT, bordercolor=self.BG_BORDER, borderwidth=1, font=("Segoe UI", 10, "bold"))
        self.style.configure("TLabelframe.Label", background=self.BG_MAIN, foreground=self.COLOR_ACCENT, font=("Segoe UI", 10, "bold"))
        self.style.configure("TLabel", background=self.BG_MAIN, foreground=self.FG_TEXT)
        self.style.configure("Card.TLabel", background=self.BG_CARD, foreground=self.FG_TEXT)
        self.style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"), foreground=self.COLOR_BLUE)
        self.style.configure("Sub.TLabel", font=("Segoe UI", 8, "italic"), foreground=self.FG_MUTED)

        # Standard Buttons
        self.style.configure("TButton", background="#e5e7eb", foreground=self.FG_TEXT, borderwidth=1, font=("Segoe UI", 9, "bold"), focuscolor="")
        self.style.map("TButton",
            background=[("active", "#d1d5db"), ("disabled", "#f3f4f6")],
            foreground=[("disabled", "#9ca3af")]
        )

        # Accent Primary Actions
        self.style.configure("Accent.TButton", background=self.COLOR_BLUE, foreground="#ffffff", borderwidth=0, font=("Segoe UI", 9, "bold"))
        self.style.map("Accent.TButton", background=[("active", "#1d4ed8"), ("disabled", "#f3f4f6")], foreground=[("disabled", "#9ca3af")])

        # Input elements
        self.style.configure("TEntry", fieldbackground=self.BG_CARD, foreground=self.FG_TEXT, insertcolor=self.FG_TEXT, bordercolor=self.BG_BORDER)
        self.style.configure("TCombobox", fieldbackground=self.BG_CARD, background="#e5e7eb", foreground=self.FG_TEXT, bordercolor=self.BG_BORDER)
        self.style.map("TCombobox", fieldbackground=[("readonly", self.BG_CARD)], foreground=[("readonly", self.FG_TEXT)])

        # Treeview styling
        self.style.configure("Treeview", background=self.BG_CARD, fieldbackground=self.BG_CARD, foreground=self.FG_TEXT, bordercolor=self.BG_BORDER, rowheight=26, font=("Segoe UI", 9))
        self.style.configure("Treeview.Heading", background="#e5e7eb", foreground=self.FG_TEXT, relief="flat", font=("Segoe UI", 9, "bold"))
        self.style.map("Treeview", background=[("selected", "#d1d5db")], foreground=[("selected", self.COLOR_ACCENT)])

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
        self.preview_text.config(state="disabled")

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

    def popup_edit_dialog(self, event):
        """Displays a modal Toplevel window to edit the URL and Payload directly"""
        selected_item = self.tree.selection()
        if not selected_item:
            return
        
        sr_val = self.tree.item(selected_item[0], "values")[1]
        api = next((a for a in API_ENDPOINTS if str(a["sr"]) == str(sr_val)), None)
        if not api:
            return

        # Create Dialog Toplevel Window
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit API Endpoint - {api['name']}")
        dialog.geometry("620x460")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg=self.BG_MAIN)

        # Center Dialog Popup relative to main app window
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (width // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (height // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")

        # Modal layout frame
        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text=f"Modifying #{api['sr']}: {api['name']}", font=("Segoe UI", 11, "bold"), foreground=self.COLOR_ACCENT).pack(anchor="w", pady=(0, 15))

        # Endpoint URL field
        ttk.Label(main_frame, text="URL:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5, 2))
        url_var = tk.StringVar(value=api.get("url", ""))
        url_entry = ttk.Entry(main_frame, textvariable=url_var, font=("Segoe UI", 9))
        url_entry.pack(fill="x", pady=(0, 10))
        url_entry.focus_set()

        # Payload ScrolledText field
        ttk.Label(main_frame, text="Payload (JSON Format / raw certificate PEM):", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5, 2))
        payload_text = scrolledtext.ScrolledText(
            main_frame, height=9, font=("Consolas", 9), wrap="none",
            bg=self.BG_CARD, fg=self.FG_TEXT, insertbackground=self.FG_TEXT,
            highlightthickness=1, highlightbackground=self.BG_BORDER, relief="flat"
        )
        payload_text.pack(fill="both", expand=True, pady=(0, 15))

        # Prepopulate existing payload values
        payload = api.get("payload")
        if payload is not None:
            if isinstance(payload, (dict, list)):
                payload_text.insert(tk.END, json.dumps(payload, indent=2))
            else:
                payload_text.insert(tk.END, str(payload))

        # Bottom buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", side="bottom")

        def save_and_close():
            new_url = url_var.get().strip()
            new_payload_str = payload_text.get("1.0", tk.END).strip()

            if not new_url:
                messagebox.showerror("Validation Error", "URL field cannot be empty.", parent=dialog)
                return

            if not new_payload_str:
                new_payload = None
            elif new_payload_str.startswith("-----BEGIN"):
                new_payload = new_payload_str
            else:
                try:
                    new_payload = json.loads(new_payload_str)
                except Exception:
                    new_payload = new_payload_str

            # Update details in endpoint object
            api["url"] = new_url
            api["payload"] = new_payload

            # Sync Treeview display with updated URL
            values = list(self.tree.item(selected_item[0], "values"))
            values[4] = new_url
            self.tree.item(selected_item[0], values=values)

            # Re-render Request Details preview
            self.update_request_preview(api)

            self.update_status(f"Saved changes for '{api['name']}'")
            dialog.destroy()
            messagebox.showinfo("Update Successful", f"Successfully updated '{api['name']}' parameters.")

        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Save Changes", command=save_and_close, style="Accent.TButton").pack(side="right", padx=5)

    def popup_add_api_dialog(self):
        """Displays a modal Toplevel window to add a new API Endpoint"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add New API Endpoint")
        dialog.geometry("620x560")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg=self.BG_MAIN)

        # Center Dialog Popup relative to main app window
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (width // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (height // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")

        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="Add New API Endpoint Configuration", font=("Segoe UI", 11, "bold"), foreground=self.COLOR_ACCENT).pack(anchor="w", pady=(0, 15))

        # API Name
        ttk.Label(main_frame, text="API Name:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5, 2))
        name_var = tk.StringVar()
        name_entry = ttk.Entry(main_frame, textvariable=name_var, font=("Segoe UI", 9))
        name_entry.pack(fill="x", pady=(0, 10))
        name_entry.focus_set()

        # Endpoint URL
        ttk.Label(main_frame, text="URL:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5, 2))
        url_var = tk.StringVar(value="http://192.168.4.1/api/")
        url_entry = ttk.Entry(main_frame, textvariable=url_var, font=("Segoe UI", 9))
        url_entry.pack(fill="x", pady=(0, 10))

        # Method
        ttk.Label(main_frame, text="Method:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5, 2))
        method_var = tk.StringVar(value="GET")
        method_cb = ttk.Combobox(main_frame, textvariable=method_var, values=["GET", "POST", "Get/POST"], state="readonly", font=("Segoe UI", 9))
        method_cb.pack(fill="x", pady=(0, 10))

        # Permissions per role
        ttk.Label(main_frame, text="Role Permissions (Select Read/Write per role):", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5, 2))
        
        perm_frame = ttk.LabelFrame(main_frame, text=" Permissions Matrix ", padding=10)
        perm_frame.pack(fill="x", pady=(0, 10))

        # Columns headers
        ttk.Label(perm_frame, text="Role", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=2)
        ttk.Label(perm_frame, text="Read", font=("Segoe UI", 9, "bold")).grid(row=0, column=1, anchor="center", padx=20, pady=2)
        ttk.Label(perm_frame, text="Write", font=("Segoe UI", 9, "bold")).grid(row=0, column=2, anchor="center", padx=20, pady=2)

        roles_list = ["Viewer", "Operator", "System Admin", "Security Admin"]
        perm_vars = {} # maps role name to (read_var, write_var)

        for i, role in enumerate(roles_list):
            ttk.Label(perm_frame, text=role).grid(row=i+1, column=0, sticky="w", padx=10, pady=2)
            
            read_var = tk.BooleanVar(value=True)
            write_var = tk.BooleanVar(value=False)
            
            # Default helper permissions
            if role == "Security Admin":
                write_var.set(True)
            elif role == "System Admin":
                write_var.set(True)
                
            read_chk = ttk.Checkbutton(perm_frame, variable=read_var)
            read_chk.grid(row=i+1, column=1, padx=20, pady=2)
            
            write_chk = ttk.Checkbutton(perm_frame, variable=write_var)
            write_chk.grid(row=i+1, column=2, padx=20, pady=2)
            
            perm_vars[role] = (read_var, write_var)

        # Payload ScrolledText field
        ttk.Label(main_frame, text="Payload (JSON Format / raw certificate PEM, Optional):", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5, 2))
        payload_text = scrolledtext.ScrolledText(
            main_frame, height=6, font=("Consolas", 9), wrap="none",
            bg=self.BG_CARD, fg=self.FG_TEXT, insertbackground=self.FG_TEXT,
            highlightthickness=1, highlightbackground=self.BG_BORDER, relief="flat"
        )
        payload_text.pack(fill="both", expand=True, pady=(0, 15))

        # Bottom buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", side="bottom")

        def add_and_close():
            name = name_var.get().strip()
            url = url_var.get().strip()
            method = method_var.get().strip()
            payload_str = payload_text.get("1.0", tk.END).strip()

            if not name:
                messagebox.showerror("Validation Error", "API Name field cannot be empty.", parent=dialog)
                return
            if not url:
                messagebox.showerror("Validation Error", "URL field cannot be empty.", parent=dialog)
                return

            if not payload_str:
                payload = None
            elif payload_str.startswith("-----BEGIN"):
                payload = payload_str
            else:
                try:
                    payload = json.loads(payload_str)
                except Exception:
                    payload = payload_str

            # Construct roles dict
            roles_dict = {}
            for role, (r_var, w_var) in perm_vars.items():
                perms = []
                if r_var.get():
                    perms.append("Read")
                if w_var.get():
                    perms.append("Write")
                if perms:
                    roles_dict[role] = perms

            # Compute next serial number
            next_sr = max(api["sr"] for api in API_ENDPOINTS) + 1 if API_ENDPOINTS else 1

            new_api = {
                "sr": next_sr,
                "name": name,
                "url": url,
                "method": method,
                "roles": roles_dict,
                "payload": payload
            }

            API_ENDPOINTS.append(new_api)

            # Sync GUI Table
            self.populate_table()
            
            # Select the newly added row
            if self.tree.exists(str(next_sr)):
                self.tree.selection_set(str(next_sr))
                self.tree.see(str(next_sr))

            self.update_status(f"Added new API Endpoint: '{name}'")
            dialog.destroy()
            messagebox.showinfo("Success", f"Successfully added API Endpoint #{next_sr}: '{name}'")

        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Add URL/API", command=add_and_close, style="Accent.TButton").pack(side="right", padx=5)

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
                self.timer_lbl.config(text=f"Token Timer: {remaining}s remaining", foreground=self.COLOR_BLUE)
                self.root.after(1000, self.update_timer_loop)
            else:
                if getattr(self, "pipeline_running", False):
                    self.root.after(1000, self.update_timer_loop)
                else:
                    self.timer_lbl.config(text="Token Timer: EXPIRED (0s)", foreground=self.COLOR_RED)
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

    def append_to_output(self, text):
        self.output_text.config(state="normal")
        self.output_text.insert(tk.END, text)
        self.output_text.config(state="disabled")
        self.output_text.see(tk.END)

    def run_3_step_auth_logic(self, log_func=None):
        """
        Executes the 3-step authentication logic.
        log_func: a function that takes a string and writes it to logs.
        Returns the final token if successful, raises an Exception otherwise.
        """
        def log(msg):
            if log_func:
                log_func(msg)

        ap_username = self.ap_name_var.get().strip()
        password = self.password_var.get().strip()

        login_api = next((a for a in API_ENDPOINTS if a["sr"] == 1), None)
        login_url = login_api["url"] if login_api else "http://192.168.4.1/api/login"

        config_api = next((a for a in API_ENDPOINTS if a["sr"] == 28), None)
        config_url = config_api["url"] if config_api else "http://192.168.4.1/api/device/config/update"

        # Create/Prepare session headers
        self.http_session.headers.update({"Content-Type": "application/json"})

        # ==================== STEP 1 ====================
        log(f"\n[Step 1] POST to {login_url}")
        payload1 = {"username": ap_username, "password": password}
        log(f"Payload: {json.dumps(payload1, indent=2)}")

        resp = self.http_session.post(login_url, json=payload1, timeout=5)
        log(f"Response ({resp.status_code}): {resp.text}")

        if resp.status_code != 200:
            raise Exception(f"Step 1 AP Login failed with status code {resp.status_code}")

        token1 = ""
        try:
            data = resp.json()
            if isinstance(data, dict):
                token1 = data.get("Data", {}).get("token") or data.get("sessionToken") or data.get("session") or ""
        except Exception:
            pass
        if not token1:
            token1 = resp.cookies.get("session") or ""

        if not token1:
            raise Exception("Failed to extract session token from Step 1 response.")

        # ==================== STEP 2 ====================
        role_map = {
            "Viewer": "Viewer",
            "Operator": "Operator",
            "System Admin": "SystemAdmin",
            "Security Admin": "SecurityAdmin"
        }
        selected_role = self.role_var.get()
        dev_username = role_map.get(selected_role, "SecurityAdmin")

        if config_api and isinstance(config_api.get("payload"), dict):
            payload2 = copy.deepcopy(config_api["payload"])
            payload2["device_username"] = dev_username
            payload2["device_password"] = password
        else:
            payload2 = {
                "type": 2,
                "device_username": dev_username,
                "device_password": password
            }

        headers2 = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token1}",
            "Cookie": f"session={token1}"
        }

        log(f"\n[Step 2] POST to {config_url}")
        log(f"Headers: {json.dumps(headers2, indent=2)}")
        log(f"Payload: {json.dumps(payload2, indent=2)}")

        resp2 = self.http_session.post(config_url, json=payload2, headers=headers2, timeout=5)
        log(f"Response ({resp2.status_code}): {resp2.text}")

        if resp2.status_code != 200:
            raise Exception(f"Step 2 config update failed with status code {resp2.status_code}")

        # ==================== STEP 3 ====================
        log(f"\n[Step 3] POST to {login_url}")
        payload3 = {"username": dev_username, "password": password}
        log(f"Payload: {json.dumps(payload3, indent=2)}")

        resp3 = self.http_session.post(login_url, json=payload3, timeout=5)
        log(f"Response ({resp3.status_code}): {resp3.text}")

        if resp3.status_code != 200:
            raise Exception(f"Step 3 authentication as {dev_username} failed with status code {resp3.status_code}")

        token3 = ""
        try:
            data3 = resp3.json()
            if isinstance(data3, dict):
                token3 = data3.get("Data", {}).get("token") or data3.get("sessionToken") or data3.get("session") or ""
        except Exception:
            pass
        if not token3:
            token3 = resp3.cookies.get("session") or ""

        if not token3:
            raise Exception("Failed to extract final session token in Step 3.")

        # Update the final headers and cookies
        self.session_token_var.set(token3)
        self.http_session.cookies.clear()
        self.http_session.headers.update({
            "Authorization": f"Bearer {token3}",
            "Cookie": f"session={token3}"
        })

        return token3

    def fetch_session_token(self):
        """Implements the new 3-Step Session Login & Credential Pushing workflow"""
        self.update_status("Initiating 3-Step Login sequence...")
        
        # Clear output area for logs
        self.output_text.config(state="normal")
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, "=== STARTING 3-STEP AUTHENTICATION PIPELINE ===\n")
        self.output_text.config(state="disabled")

        def _thread():
            def log_to_gui(msg):
                self.root.after(0, lambda: self.append_to_output(msg + "\n"))

            try:
                token3 = self.run_3_step_auth_logic(log_to_gui)
                self.update_status("Session connection established!")
                self.root.after(0, self.start_1min_timer)
                self.root.after(0, lambda: messagebox.showinfo("Authentication Success", f"Session token successfully established: {token3}"))
            except Exception as ex:
                self.update_status(f"Pipeline error: {str(ex)}")
                self.root.after(0, lambda: messagebox.showerror("Pipeline Error", f"Exception during authentication sequence:\n{str(ex)}"))

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

        # Synchronize session headers with the GUI token
        token = self.session_token_var.get().strip()
        self.http_session.headers.update({
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}",
            "Cookie": f"session={token}"
        })

        WRITE_TO_READ_MAPPING = {3: 4, 5: 6, 7: 8, 10: 11, 13: 12, 17: 16, 24: 23, 26: 26}
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

            # Retrieve payload statically
            if api["sr"] == 1:
                active_payload = {
                    "username": self.ap_name_var.get().strip(),
                    "password": self.password_var.get().strip()
                }
            elif api["sr"] == 28:
                role_map = {
                    "Viewer": "Viewer",
                    "Operator": "Operator",
                    "System Admin": "SystemAdmin",
                    "Security Admin": "SecurityAdmin"
                }
                active_payload = copy.deepcopy(api["payload"])
                if isinstance(active_payload, dict):
                    active_payload["device_username"] = role_map.get(selected_role, "SecurityAdmin")
                    active_payload["device_password"] = self.password_var.get().strip()
            else:
                active_payload = api.get("payload")
            
            self.root.after(0, lambda a=api, p=active_payload: self.update_request_preview(a, p))

            body_output = ""
            write_status = "-"
            read_status = "-"

            # STEP 1: Execute Write (POST API)
            if (orig_write_acc == "YES" or has_perm_restart) and ("POST" in api["method"].upper() or "GET/POST" in api["method"].upper()):
                try:
                    if api["sr"] == 1:
                        body_output += "--- STEP 1: RUNNING 3-STEP AUTHENTICATION SEQUENCE ---\n"
                        def log_to_body(msg):
                            nonlocal body_output
                            body_output += msg + "\n"
                        token_res = self.run_3_step_auth_logic(log_to_body)
                        write_status = "200 OK"
                        body_output += f"\nSuccess! Session token established: {token_res}\n"
                        self.root.after(0, self.start_1min_timer)
                    else:
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
                        body_output += f"--- STEP 1: POST WRITE RESPONSE ---\nStatus: {write_status}\n{r.text}\n"
                except Exception as ex:
                    write_status = "Error/Offline"
                    body_output += f"--- STEP 1: POST ERROR ---\n{str(ex)}\n"
            else:
                write_status = "-----"

            # STEP 2: Strict Delay (1ms pause between Write and Read)
            time.sleep(0.001)

            # STEP 3: Execute Read (GET API) and Verify Payload Match
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
            
            res_entry = {
                "sr": api["sr"],
                "name": api["name"],
                "method": api["method"],
                "url": api["url"],
                "read_status": read_status,
                "write_status": write_status,
                "body": body_output
            }
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
            messagebox.showwarning("No Data", "No execution results to export.")
            return

        pdf_path = "API_Report.pdf"
        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=letter,
            rightMargin=36,
            leftMargin=36,
            topMargin=36,
            bottomMargin=36
        )

        styles = getSampleStyleSheet()

        # Define premium typography / colors
        color_primary = colors.HexColor("#7c3aed") # Deep Purple
        color_secondary = colors.HexColor("#2563eb") # Royal Blue
        color_text = colors.HexColor("#1f2937")
        color_bg_code = colors.HexColor("#f9fafb")
        color_border_code = colors.HexColor("#e5e7eb")

        title_style = ParagraphStyle(
            'DocTitle',
            parent=styles['Heading1'],
            fontName='Helvetica-Bold',
            fontSize=20,
            leading=24,
            textColor=color_primary,
            spaceAfter=15
        )

        section_style = ParagraphStyle(
            'SectionTitle',
            parent=styles['Heading2'],
            fontName='Helvetica-Bold',
            fontSize=14,
            leading=18,
            textColor=color_secondary,
            spaceBefore=12,
            spaceAfter=8
        )

        item_title_style = ParagraphStyle(
            'ItemTitle',
            parent=styles['Heading3'],
            fontName='Helvetica-Bold',
            fontSize=11,
            leading=14,
            textColor=color_primary,
            spaceBefore=10,
            spaceAfter=4
        )

        meta_style = ParagraphStyle(
            'MetaStyle',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=9,
            leading=12,
            textColor=color_text
        )

        code_style = ParagraphStyle(
            'CodeStyle',
            parent=styles['Normal'],
            fontName='Courier',
            fontSize=7,
            leading=9,
            textColor=colors.HexColor("#374151")
        )

        story = []

        # 1. Title
        story.append(Paragraph("RMS Device Suite - API Execution Report", title_style))

        # 2. Metadata Block
        meta_text = (
            f"<b>Date/Time:</b> {time.strftime('%Y-%m-%d %H:%M:%S')}<br/>"
            f"<b>Target AP:</b> {self.ap_name_var.get().strip()}<br/>"
            f"<b>Role:</b> {self.role_var.get()}<br/>"
            f"<b>Session Token:</b> {self.session_token_var.get().strip() or 'None'}"
        )
        story.append(Paragraph(meta_text, meta_style))
        story.append(Spacer(1, 15))

        # 3. Overview Table
        story.append(Paragraph("Execution Overview", section_style))
        table_data = [["Sr.", "API Name", "Method", "URL", "Read Status", "Write Status"]]
        for item in self.execution_results:
            name_p = Paragraph(html.escape(item["name"]), meta_style)
            url_p = Paragraph(html.escape(item.get("url", "-")), meta_style)
            table_data.append([
                str(item["sr"]),
                name_p,
                item.get("method", "-"),
                url_p,
                item["read_status"],
                item["write_status"]
            ])

        col_widths = [25, 120, 50, 185, 80, 80]
        overview_table = Table(table_data, colWidths=col_widths)
        overview_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#f3f4f6")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#1f2937")),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8.5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#d1d5db")),
        ]))
        story.append(overview_table)
        story.append(Spacer(1, 20))

        # 4. Detailed Logs Section
        story.append(PageBreak())
        story.append(Paragraph("Detailed Execution Logs", section_style))
        story.append(Spacer(1, 10))

        for item in self.execution_results:
            details_flowables = []
            details_flowables.append(Paragraph(f"Sr {item['sr']}: {html.escape(item['name'])}", item_title_style))

            info_text = (
                f"<b>URL:</b> {html.escape(item.get('url', '-'))} | "
                f"<b>Method:</b> {item.get('method', '-')} | "
                f"<b>Read Status:</b> {item['read_status']} | "
                f"<b>Write Status:</b> {item['write_status']}"
            )
            details_flowables.append(Paragraph(info_text, meta_style))
            details_flowables.append(Spacer(1, 5))

            escaped_body = html.escape(item["body"])
            escaped_body = escaped_body.replace("\n", "<br/>").replace(" ", "&nbsp;")
            body_p = Paragraph(escaped_body, code_style)

            code_table = Table([[body_p]], colWidths=[540])
            code_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), color_bg_code),
                ('GRID', (0, 0), (-1, -1), 0.5, color_border_code),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('LEFTPADDING', (0, 0), (-1, -1), 8),
                ('RIGHTPADDING', (0, 0), (-1, -1), 8),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ]))

            details_flowables.append(code_table)
            details_flowables.append(Spacer(1, 12))

            story.append(KeepTogether(details_flowables))

        try:
            doc.build(story)
            messagebox.showinfo("Export Successful", f"Detailed report exported successfully to {pdf_path}")
        except Exception as ex:
            messagebox.showerror("Export Error", f"Failed to build PDF report:\n{str(ex)}")

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
