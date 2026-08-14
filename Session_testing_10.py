import os
import time
import json
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import requests
from PIL import Image, ImageTk

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

# Complete Master API Mapping
API_ENDPOINTS = [
    {"sr": 1, "name": "Authentication API (Write API Login)", "url": "http://192.168.4.1/api/login", "method": "POST", "roles": {"Viewer": ["Write"], "Operator": ["Write"], "System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": None},
    {"sr": 2, "name": "Authentication API (Read API Login)", "url": "http://192.168.4.1/api/auth/status", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 3, "name": "UUID Checking Single-phase", "url": "http://192.168.4.1/api/config/parameters", "method": "POST", "roles": {"System Admin": ["Write"], "Security Admin": ["Write"]}, "payload": {"vdinterval": 5, "table": 1, "parameters": []}},
    {"sr": 4, "name": "UUID Checking Single-phase(Get)", "url": "http://192.168.4.1/api/config/parameters?table=1", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 5, "name": "ISP Configuration API", "url": "http://192.168.4.1/api/config/isp", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"apn": "airtelgprs.com", "apn2": "airtelgprs.com", "current_sim": "1"}},
    {"sr": 6, "name": "Get ISP Configuration", "url": "http://192.168.4.1/api/config/ISP", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 7, "name": "Remote Server Configuration API", "url": "http://192.168.4.1/api/config/remote-server", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"server_url": "rms.iotscada-pmsg.com", "server_port": 8883, "solution_type": "ongridrooftop", "client_id": "d:866738083608743$ongridrooftop$510017", "username": "866738083608743$ongridrooftop$510017", "password": "31c1074a", "server_url1": "rms.iotscada-pmsg.com", "server_port1": 8883, "solution_type1": "ongridrooftop", "client_id1": "d:866082075799828$ongridrooftop$500092", "username1": "866082075799828$ongridrooftop$500092", "password1": "466b856f", "imei": "866738083608743", "imei1": "866082075799828"}},
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
    {"sr": 21, "name": "Restart", "url": "http://192.168.4.1/restart", "method": "POST", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": {"action": "reboot"}},
    {"sr": 23, "name": "MQTTServer Get", "url": "http://192.168.4.1/api/config/mqtt-server", "method": "GET", "roles": {"Viewer": ["Read"], "Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read"]}, "payload": None},
    {"sr": 24, "name": "MQTTServer Post", "url": "http://192.168.4.1/api/config/mqtt-server", "method": "POST", "roles": {"Security Admin": ["Write"]}, "payload": {"http_url": "api.iotscada-pmsg.com", "http_port": 443, "imei": "866738083608743", "username": "866738083608743", "password": "31c1074a"}},
    {"sr": 25, "name": "Firmware Update", "url": "http://192.168.4.1/update", "method": "POST", "roles": {"Operator": ["Read"], "System Admin": ["Read"], "Security Admin": ["Read", "Write"]}, "payload": {"file": "fw_v1.bin"}},
    {"sr": 26, "name": "Modbus Poll Access", "url": "http://192.168.4.1/api/modbus", "method": "POST", "roles": {"Operator": ["Read", "Write"], "System Admin": ["Read"], "Security Admin": ["Write"]}, "payload": {"poll": True}},
    {"sr": 27, "name": "Fiddler Request", "url": "http://192.168.4.1:85/list.html", "method": "GET", "roles": {}, "payload": None}
]

# Quick documentation descriptions for the right-hand Reference tab
API_DESCRIPTIONS = {
    1: "Performs user login for a specific device role. Returns a session cookie. Required as the primary authentication step.",
    2: "Validates session status and returns active permissions for the currently logged-in user role.",
    3: "Pushes single-phase operational config parameter database tables to the device storage.",
    4: "Retrieves single-phase operational parameter settings database records from the device.",
    5: "Sets cellular ISP parameters, APN names for SIM Card 1/2, and locks the active SIM slot.",
    6: "Fetches cellular APN status and configurations stored on the hardware module.",
    7: "Configures central RMS server URLs, ports, transmission protocols (MQTT/TCP), solutions, client IDs, and SIM IMEIs.",
    8: "Reads remote server URLs, solutions, client IDs, and IMEI numbers currently configured.",
    10: "Instructs the device to establish secure socket connection to the configured MQTT broker.",
    11: "Fetches live connection telemetry logs and status flags for the active MQTT Broker connection.",
    12: "Reads physical RS485 communication config parameters (Baudrate, parity, start/stop bits).",
    13: "Updates physical RS485 Modbus parameters (Baudrate, request limits, serial configurations).",
    15: "Downloads batch diagnostic logs or historical register parameters for a designated calendar day.",
    16: "Checks local Wi-Fi target credentials set for local network sync checks.",
    17: "Modifies the destination Wi-Fi network SSID configuration for diagnostic linking.",
    18: "Saves root PEM certification file (rootCA.pem) to flash memory for TLS Handshake verification.",
    19: "Saves device private key PEM file (key.pem) to flash memory for authentication.",
    20: "Saves client digital certification PEM file (client.pem) to flash memory for mutual authentication (mTLS).",
    21: "Commands the gateway to perform a hard soft reboot. Resets all sub-modules immediately.",
    23: "Retrieves parameters of the secure secondary MQTT server.",
    24: "Sets server URL, port, IMEI target, and client authentication tokens for secondary server linking.",
    25: "Starts device OTA flash sequence using the specified binary firmware file.",
    26: "Toggles active Modbus polling scheduler loop (On / Off) on the micro-controller.",
    27: "Auxiliary network file download link verification endpoint."
}

class MockResponse:
    def __init__(self, status_code, text):
        self.status_code = status_code
        self.text = text
        self.reason = "OK" if status_code == 200 else "Error"

    def json(self):
        return json.loads(self.text)

class ModernRMSTesterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RMS Device AP Suite Pro")
        self.root.geometry("1400x900")
        self.root.minsize(1100, 750)

        # Variables for states
        self.execution_results = []
        self.cancel_event = threading.Event()
        self.session_expire_time = None
        self.timer_running = False
        self.session_token_var = tk.StringVar(value="")

        # Live WiFi Auto-scan variables
        self.scanned_ssids = []
        self.current_ssid = "Scanning..."
        self.ping_latency = "N/A"
        self.ping_history = []
        self.diag_running = True

        # Track workflow stages (None: Missing, True: Success, False: Failed)
        self.step_states = {
            1: False, # AP Connected
            2: False, # Logged In
            3: False, # Certs Written
            4: False, # Configs Written
            5: False, # Broker Status Checked
            6: False  # Reboot Sent
        }

        # Initialize modern theme & UI elements
        self.setup_styles()
        self.build_layout()

        # Start Wi-Fi and Diagnostics Background Daemon threads
        threading.Thread(target=self.wifi_auto_scan_loop, daemon=True).start()
        threading.Thread(target=self.ping_diagnostics_loop, daemon=True).start()

        # Populate tables
        self.populate_table()

    def setup_styles(self):
        """Configure modern Catppuccin Mocha-style dark theme for ttk widgets"""
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # Color codes
        self.BG_MAIN = "#181825"
        self.BG_CARD = "#1e1e2e"
        self.BG_BORDER = "#313244"
        self.FG_TEXT = "#cdd6f4"
        self.FG_MUTED = "#a6adc8"
        self.COLOR_ACCENT = "#cba6f7" # Lavender
        self.COLOR_BLUE = "#89b4fa"   # Blue
        self.COLOR_GREEN = "#a6e3a1"  # Green
        self.COLOR_YELLOW = "#f9e2af" # Yellow
        self.COLOR_RED = "#f38ba8"    # Red

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
        self.style.configure("TButton", background="#313244", foreground=self.FG_TEXT, borderwidth=1, font=("Segoe UI", 9, "bold"), focuscolor="")
        self.style.map("TButton",
            background=[("active", "#45475a"), ("disabled", "#11111b")],
            foreground=[("disabled", "#585b70")]
        )

        # Accent Primary Actions
        self.style.configure("Accent.TButton", background=self.COLOR_BLUE, foreground="#11111b", borderwidth=0, font=("Segoe UI", 9, "bold"))
        self.style.map("Accent.TButton", background=[("active", "#b4befe"), ("disabled", "#11111b")])

        # Step Workflow Buttons
        self.style.configure("Step.TButton", background="#313244", foreground=self.FG_TEXT, font=("Segoe UI", 9, "bold"))
        self.style.map("Step.TButton", background=[("active", "#45475a")])

        # Input elements
        self.style.configure("TEntry", fieldbackground=self.BG_CARD, foreground=self.FG_TEXT, insertcolor=self.FG_TEXT, bordercolor=self.BG_BORDER)
        self.style.configure("TCombobox", fieldbackground=self.BG_CARD, background="#313244", foreground=self.FG_TEXT, bordercolor=self.BG_BORDER)
        self.style.map("TCombobox", fieldbackground=[("readonly", self.BG_CARD)], foreground=[("readonly", self.FG_TEXT)])

        # Notebook tabs
        self.style.configure("TNotebook", background=self.BG_MAIN, borderwidth=0)
        self.style.configure("TNotebook.Tab", background="#313244", foreground=self.FG_TEXT, font=("Segoe UI", 9, "bold"), padding=6)
        self.style.map("TNotebook.Tab", background=[("selected", self.BG_CARD)], foreground=[("selected", self.COLOR_ACCENT)])

        # Treeview styling
        self.style.configure("Treeview", background=self.BG_CARD, fieldbackground=self.BG_CARD, foreground=self.FG_TEXT, bordercolor=self.BG_BORDER, rowheight=26, font=("Segoe UI", 9))
        self.style.configure("Treeview.Heading", background="#313244", foreground=self.FG_TEXT, relief="flat", font=("Segoe UI", 9, "bold"))
        self.style.map("Treeview", background=[("selected", "#45475a")], foreground=[("selected", "#cba6f7")])

    def on_simulator_toggle(self):
        is_sim = self.simulator_mode_var.get()
        if is_sim:
            self.wifi_status_var.set("📡 AP Link: Simulator Active")
            self.wifi_badge.config(foreground=self.COLOR_GREEN)
            self.ping_status_var.set("🔗 Ping Device: Simulated (2ms)")
            self.ping_badge.config(foreground=self.COLOR_GREEN)
            self.s1_badge.config(text="🟢 Linked", foreground=self.COLOR_GREEN)
            self.step_states[1] = True
            self.update_status("Switched to Simulator Mode. Direct requests will be mocked.")
        else:
            self.step_states[1] = False
            self.s1_badge.config(text="🔴 Scanning...", foreground=self.COLOR_RED)
            self.update_status("Switched to Gateway Direct Mode. Resolving connection...")

    def perform_request(self, method, url, headers=None, json_payload=None, raw_payload=None, timeout=5):
        if self.simulator_mode_var.get():
            time.sleep(0.4) # Simulate network delay
            url_lower = url.lower()
            status_code = 200
            response_text = '{"status": "success"}'

            if "/api/login" in url_lower:
                response_text = '{"sessionToken": "mock_session_token_12345"}'
            elif "/api/auth/status" in url_lower:
                response_text = f'{{"status": "authenticated", "role": "{self.role_var.get()}"}}'
            elif "/api/config/isp" in url_lower:
                response_text = '{"status": "success", "apn": "airtelgprs.com", "current_sim": "1"}'
            elif "/api/config/remote-server" in url_lower:
                response_text = '{"server_url": "rms.iotscada-pmsg.com", "server_port": 8883, "solution_type": "ongridrooftop"}'
            elif "/api/device/broker/connect" in url_lower:
                response_text = '{"action": "connect", "status": "connecting"}'
            elif "/api/device/broker/status" in url_lower:
                response_text = "Connected to secure MQTT broker rms.iotscada-pmsg.com:8883. Channel active."
            elif "/api/config/inverter-communication" in url_lower:
                response_text = '{"asn": "Yash", "baudrate": 9600, "parity": 1, "stopBit": 1, "databits": 8}'
            elif "/api/config/parameters" in url_lower:
                response_text = '{"status": "success", "parameters_count": 12}'
            elif "/api/history" in url_lower:
                response_text = "Timestamp,Voltage,Current,Power\n2026-08-13 18:00,230.1,4.5,1035.4\n2026-08-13 18:05,230.3,4.4,1013.3"
            elif "/api/device/config/update" in url_lower:
                response_text = '{"status": "wifi_configured", "ssid": "test_wifi"}'
            elif "filename=rootca.pem" in url_lower:
                response_text = "Root CA saved successfully."
            elif "filename=key.pem" in url_lower:
                response_text = "Private Key saved successfully."
            elif "filename=client.pem" in url_lower:
                response_text = "Client Cert saved successfully."
            elif "/restart" in url_lower:
                response_text = '{"status": "rebooting", "action": "reboot"}'
            elif "/api/config/mqtt-server" in url_lower:
                response_text = '{"http_url": "api.iotscada-pmsg.com", "http_port": 443, "imei": "866738083608743"}'
            elif "/update" in url_lower:
                response_text = "OTA Firmware update initiated successfully."
            elif "/api/modbus" in url_lower:
                response_text = '{"poll": true, "status": "polling_active"}'
            elif "/list.html" in url_lower:
                response_text = "<html><body>Mock files list: rootCA.pem, key.pem, client.pem</body></html>"

            return MockResponse(status_code, response_text)
        else:
            if method == "GET":
                return requests.get(url, headers=headers, timeout=timeout)
            else:
                if json_payload is not None:
                    return requests.post(url, json=json_payload, headers=headers, timeout=timeout)
                else:
                    return requests.post(url, data=raw_payload, headers=headers, timeout=timeout)

    def build_layout(self):
        """Build the responsive 3-section layout: Steps on left, Grid in center, Collapsible Docs on right"""
        # --- Top Header Status Panel ---
        self.header_frame = ttk.Frame(self.root, padding=10)
        self.header_frame.pack(fill="x", side="top")

        title_lbl = ttk.Label(self.header_frame, text="⚡ RMS GATEWAY ACCESS POINT TESTING SUITE", font=("Segoe UI", 14, "bold"), foreground=self.COLOR_BLUE)
        title_lbl.pack(side="left")

        # Simulator checkbox
        self.simulator_mode_var = tk.BooleanVar(value=False)
        self.sim_checkbox = ttk.Checkbutton(self.header_frame, text="🌐 Simulator Mode", variable=self.simulator_mode_var, command=self.on_simulator_toggle)
        self.sim_checkbox.pack(side="left", padx=20)

        # Connection status badge
        self.wifi_status_var = tk.StringVar(value="📡 Wi-Fi Status: Scanning...")
        self.wifi_badge = ttk.Label(self.header_frame, textvariable=self.wifi_status_var, font=("Segoe UI", 10, "bold"), foreground=self.COLOR_YELLOW, padding=(10, 2))
        self.wifi_badge.pack(side="right", padx=10)

        self.ping_status_var = tk.StringVar(value="🔗 Ping Device: Detecting...")
        self.ping_badge = ttk.Label(self.header_frame, textvariable=self.ping_status_var, font=("Segoe UI", 10, "bold"), foreground=self.COLOR_YELLOW, padding=(10, 2))
        self.ping_badge.pack(side="right", padx=10)

        # Main horizontal split
        self.main_pane = ttk.PanedWindow(self.root, orient="horizontal")
        self.main_pane.pack(fill="both", expand=True, padx=10, pady=5)

        # ----------------------------------------------------
        # SECTION 1: Left Step-by-Step Workflow Column (Width: fixed)
        # ----------------------------------------------------
        self.left_col = ttk.Frame(self.main_pane, padding=5)
        self.main_pane.add(self.left_col, weight=1)

        steps_frame = ttk.LabelFrame(self.left_col, text=" 🔗 Process Workflow ", padding=8)
        steps_frame.pack(fill="both", expand=True)

        # Scrollable container for steps in case of small screen heights
        canvas = tk.Canvas(steps_frame, bg=self.BG_MAIN, highlightthickness=0)
        scrollbar = ttk.Scrollbar(steps_frame, orient="vertical", command=canvas.yview)
        self.scrollable_steps = ttk.Frame(canvas)
        self.scrollable_steps.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.scrollable_steps, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Step 1 Card: SSID / AP Connection
        s1_card = ttk.Frame(self.scrollable_steps, style="Card.TFrame", padding=10)
        s1_card.pack(fill="x", pady=6, ipady=4)
        s1_title_frame = ttk.Frame(s1_card, style="Card.TFrame")
        s1_title_frame.pack(fill="x")
        ttk.Label(s1_title_frame, text="Step 1: Connect to RMS AP", font=("Segoe UI", 10, "bold"), foreground=self.COLOR_ACCENT, style="Card.TLabel").pack(side="left")
        self.s1_badge = ttk.Label(s1_title_frame, text="🔴 Not Detected", font=("Segoe UI", 8, "bold"), foreground=self.COLOR_RED, style="Card.TLabel")
        self.s1_badge.pack(side="right")

        ttk.Label(s1_card, text="Select Detected AP Network:", style="Card.TLabel", font=("Segoe UI", 8)).pack(fill="x", pady=(5, 2))
        self.ap_select_var = tk.StringVar()
        self.ap_select_cb = ttk.Combobox(s1_card, textvariable=self.ap_select_var, values=["Scanning..."], state="readonly", height=5)
        self.ap_select_cb.pack(fill="x", pady=2)
        self.ap_select_cb.bind("<<ComboboxSelected>>", self.on_ap_combo_selected)

        s1_btn_bar = ttk.Frame(s1_card, style="Card.TFrame")
        s1_btn_bar.pack(fill="x", pady=(5, 0))
        ttk.Button(s1_btn_bar, text="🔄 Refresh Wi-Fi AP List", command=self.trigger_wifi_scan, style="TButton").pack(fill="x")

        # Step 2 Card: Login & Credentials
        s2_card = ttk.Frame(self.scrollable_steps, style="Card.TFrame", padding=10)
        s2_card.pack(fill="x", pady=6, ipady=4)
        s2_title_frame = ttk.Frame(s2_card, style="Card.TFrame")
        s2_title_frame.pack(fill="x")
        ttk.Label(s2_title_frame, text="Step 2: Obtain Access Token", font=("Segoe UI", 10, "bold"), foreground=self.COLOR_ACCENT, style="Card.TLabel").pack(side="left")
        self.s2_badge = ttk.Label(s2_title_frame, text="🔴 No Session", font=("Segoe UI", 8, "bold"), foreground=self.COLOR_RED, style="Card.TLabel")
        self.s2_badge.pack(side="right")

        grid_cred = ttk.Frame(s2_card, style="Card.TFrame")
        grid_cred.pack(fill="x", pady=5)
        grid_cred.columnconfigure(1, weight=1)

        ttk.Label(grid_cred, text="AP SSID:", style="Card.TLabel", font=("Segoe UI", 8)).grid(row=0, column=0, sticky="w", pady=2)
        self.ap_name_var = tk.StringVar(value="RMS-2088")
        self.ap_entry = ttk.Entry(grid_cred, textvariable=self.ap_name_var, font=("Segoe UI", 9))
        self.ap_entry.grid(row=0, column=1, sticky="ew", padx=(5, 0), pady=2)

        ttk.Label(grid_cred, text="Role Select:", style="Card.TLabel", font=("Segoe UI", 8)).grid(row=1, column=0, sticky="w", pady=2)
        self.role_var = tk.StringVar(value="Viewer")
        self.role_cb = ttk.Combobox(grid_cred, textvariable=self.role_var, values=list(USER_CREDENTIALS.keys()), state="readonly")
        self.role_cb.grid(row=1, column=1, sticky="ew", padx=(5, 0), pady=2)
        self.role_cb.bind("<<ComboboxSelected>>", self.on_role_change)

        ttk.Label(grid_cred, text="Password:", style="Card.TLabel", font=("Segoe UI", 8)).grid(row=2, column=0, sticky="w", pady=2)
        self.password_var = tk.StringVar(value=USER_CREDENTIALS["Viewer"])
        self.pwd_entry = ttk.Entry(grid_cred, textvariable=self.password_var, font=("Segoe UI", 9))
        self.pwd_entry.grid(row=2, column=1, sticky="ew", padx=(5, 0), pady=2)

        self.s2_timer_lbl = ttk.Label(s2_card, text="Timer Status: Idle (60s Limit)", style="Card.TLabel", font=("Segoe UI", 8, "italic"), foreground=self.FG_MUTED)
        self.s2_timer_lbl.pack(fill="x", pady=3)

        ttk.Button(s2_card, text="🔑 1. Login (/api/login)", command=self.run_step_2_login, style="Accent.TButton").pack(fill="x", pady=(5, 0))

        # Step 3 Card: Write Certificates
        s3_card = ttk.Frame(self.scrollable_steps, style="Card.TFrame", padding=10)
        s3_card.pack(fill="x", pady=6)
        s3_title_frame = ttk.Frame(s3_card, style="Card.TFrame")
        s3_title_frame.pack(fill="x")
        ttk.Label(s3_title_frame, text="Step 3: Certificates Setup", font=("Segoe UI", 10, "bold"), foreground=self.COLOR_ACCENT, style="Card.TLabel").pack(side="left")
        self.s3_badge = ttk.Label(s3_title_frame, text="🔴 Missing", font=("Segoe UI", 8, "bold"), foreground=self.COLOR_RED, style="Card.TLabel")
        self.s3_badge.pack(side="right")
        ttk.Label(s3_card, text="Upload RootCA, Key, and Client certs.", style="Card.TLabel", font=("Segoe UI", 8), foreground=self.FG_MUTED).pack(fill="x", pady=(2, 5))
        ttk.Button(s3_card, text="📜 2. Upload Credentials", command=self.run_step_3_certs, style="TButton").pack(fill="x")

        # Step 4 Card: Core Configurations
        s4_card = ttk.Frame(self.scrollable_steps, style="Card.TFrame", padding=10)
        s4_card.pack(fill="x", pady=6)
        s4_title_frame = ttk.Frame(s4_card, style="Card.TFrame")
        s4_title_frame.pack(fill="x")
        ttk.Label(s4_title_frame, text="Step 4: Push Core Configs", font=("Segoe UI", 10, "bold"), foreground=self.COLOR_ACCENT, style="Card.TLabel").pack(side="left")
        self.s4_badge = ttk.Label(s4_title_frame, text="🔴 Missing", font=("Segoe UI", 8, "bold"), foreground=self.COLOR_RED, style="Card.TLabel")
        self.s4_badge.pack(side="right")
        ttk.Label(s4_card, text="Configure ISP APN & Server URL endpoints.", style="Card.TLabel", font=("Segoe UI", 8), foreground=self.FG_MUTED).pack(fill="x", pady=(2, 5))
        ttk.Button(s4_card, text="⚙️ 3. Write Settings Parameters", command=self.run_step_4_configs, style="TButton").pack(fill="x")

        # Step 5 Card: Broker Telemetry Status Check
        s5_card = ttk.Frame(self.scrollable_steps, style="Card.TFrame", padding=10)
        s5_card.pack(fill="x", pady=6)
        s5_title_frame = ttk.Frame(s5_card, style="Card.TFrame")
        s5_title_frame.pack(fill="x")
        ttk.Label(s5_title_frame, text="Step 5: Connection Diagnostics", font=("Segoe UI", 10, "bold"), foreground=self.COLOR_ACCENT, style="Card.TLabel").pack(side="left")
        self.s5_badge = ttk.Label(s5_title_frame, text="🔴 Missing", font=("Segoe UI", 8, "bold"), foreground=self.COLOR_RED, style="Card.TLabel")
        self.s5_badge.pack(side="right")
        ttk.Label(s5_card, text="Trigger MQTT client and Modbus polling check.", style="Card.TLabel", font=("Segoe UI", 8), foreground=self.FG_MUTED).pack(fill="x", pady=(2, 5))
        ttk.Button(s5_card, text="📡 4. Verify Active Broker Status", command=self.run_step_5_broker, style="TButton").pack(fill="x")

        # Step 6 Card: Reboot Operations
        s6_card = ttk.Frame(self.scrollable_steps, style="Card.TFrame", padding=10)
        s6_card.pack(fill="x", pady=6)
        s6_title_frame = ttk.Frame(s6_card, style="Card.TFrame")
        s6_title_frame.pack(fill="x")
        ttk.Label(s6_title_frame, text="Step 6: Reboot System", font=("Segoe UI", 10, "bold"), foreground=self.COLOR_ACCENT, style="Card.TLabel").pack(side="left")
        self.s6_badge = ttk.Label(s6_title_frame, text="🔴 Idle", font=("Segoe UI", 8, "bold"), foreground=self.COLOR_RED, style="Card.TLabel")
        self.s6_badge.pack(side="right")
        ttk.Label(s6_card, text="Perform hardware reboot to apply setups.", style="Card.TLabel", font=("Segoe UI", 8), foreground=self.FG_MUTED).pack(fill="x", pady=(2, 5))
        ttk.Button(s6_card, text="🔄 5. Reboot Hardware", command=self.run_step_6_reboot, style="TButton").pack(fill="x")

        # ----------------------------------------------------
        # SECTION 2: Center Grid and Preview Window
        # ----------------------------------------------------
        self.center_col = ttk.Frame(self.main_pane, padding=5)
        self.main_pane.add(self.center_col, weight=3)

        # Upper Card: Target request preview
        req_frame = ttk.LabelFrame(self.center_col, text=" 📊 Outgoing Request Preview ", padding=8)
        req_frame.pack(fill="x", side="top", pady=(0, 5))

        self.preview_text = scrolledtext.ScrolledText(req_frame, height=5, font=("Consolas", 9), wrap="word", bg="#181825", fg="#a6e3a1", insertbackground="#cdd6f4", highlightthickness=1, highlightbackground="#313244", bd=0)
        self.preview_text.pack(fill="both", expand=True)
        self.update_request_preview("N/A", "N/A", {}, None)

        # Middle Card: Master Grid Table
        grid_frame = ttk.LabelFrame(self.center_col, text=" 💻 Master API Endpoints List (Custom Testing Grid) ", padding=8)
        grid_frame.pack(fill="both", expand=True)

        # Toolbar inside Grid Frame
        grid_toolbar = ttk.Frame(grid_frame)
        grid_toolbar.pack(fill="x", pady=(0, 5))

        ttk.Button(grid_toolbar, text="☑ Select All", command=self.select_all_apis, style="TButton").pack(side="left", padx=2)
        ttk.Button(grid_toolbar, text="☐ Deselect All", command=self.deselect_all_apis, style="TButton").pack(side="left", padx=2)

        # Global execute buttons
        self.run_btn = ttk.Button(grid_toolbar, text="🚀 Run Checked APIs", command=self.start_process, style="Accent.TButton")
        self.run_btn.pack(side="right", padx=5)

        self.cancel_btn = ttk.Button(grid_toolbar, text="🛑 Cancel", command=self.cancel_process, state="disabled", style="TButton")
        self.cancel_btn.pack(side="right", padx=2)

        # Scrollable table treeview
        cols = ("select", "sr", "name", "method", "url", "read_acc", "write_acc", "status")
        self.tree = ttk.Treeview(grid_frame, columns=cols, show="headings", selectmode="browse")

        self.tree.heading("select", text="Test?")
        self.tree.heading("sr", text="Sr.")
        self.tree.heading("name", text="API Target Endpoint Name")
        self.tree.heading("method", text="Method")
        self.tree.heading("url", text="Destination URL")
        self.tree.heading("read_acc", text="Read OK")
        self.tree.heading("write_acc", text="Write OK")
        self.tree.heading("status", text="Last Call Status")

        self.tree.column("select", width=45, anchor="center")
        self.tree.column("sr", width=35, anchor="center")
        self.tree.column("name", width=220)
        self.tree.column("method", width=60, anchor="center")
        self.tree.column("url", width=200)
        self.tree.column("read_acc", width=65, anchor="center")
        self.tree.column("write_acc", width=65, anchor="center")
        self.tree.column("status", width=120, anchor="center")

        tree_scroll = ttk.Scrollbar(grid_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=tree_scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<<TreeviewSelect>>", self.on_select_tree_item)

        # Bottom Response Display Pane
        resp_frame = ttk.LabelFrame(self.center_col, text=" 📥 Received Response Output ", padding=8)
        resp_frame.pack(fill="x", side="bottom", pady=(5, 0))
        self.output_text = scrolledtext.ScrolledText(resp_frame, height=9, font=("Consolas", 9), wrap="word", bg="#181825", fg="#cdd6f4", insertbackground="#cdd6f4", highlightthickness=1, highlightbackground="#313244", bd=0)
        self.output_text.pack(fill="both", expand=True)

        # ----------------------------------------------------
        # SECTION 3: Right Collapsible Notebook Sidebar
        # ----------------------------------------------------
        self.right_col = ttk.Frame(self.main_pane, padding=5)
        self.main_pane.add(self.right_col, weight=2)

        # Sidebar content holder
        self.sidebar_frame = ttk.LabelFrame(self.right_col, text=" 📖 Info & Diagnostics ", padding=5)
        self.sidebar_frame.pack(fill="both", expand=True)

        # Tab layout
        self.sidebar_tabs = ttk.Notebook(self.sidebar_frame)
        self.sidebar_tabs.pack(fill="both", expand=True)

        # Tab 1: API Endpoint Documentation
        self.tab_doc = ttk.Frame(self.sidebar_tabs)
        self.sidebar_tabs.add(self.tab_doc, text="📚 API Reference")
        self.doc_scroll = scrolledtext.ScrolledText(self.tab_doc, font=("Segoe UI", 9), wrap="word", bg="#181825", fg="#cdd6f4", state="disabled", highlightthickness=0, bd=0)
        self.doc_scroll.pack(fill="both", expand=True, padx=5, pady=5)
        self.show_api_documentation(1) # Default to Sr. 1

        # Tab 2: Certificate Code Copy
        self.tab_certs = ttk.Frame(self.sidebar_tabs)
        self.sidebar_tabs.add(self.tab_certs, text="🔑 Certificates Vault")
        
        ttk.Label(self.tab_certs, text="Select Certificate File to Copy/View:").pack(anchor="w", padx=5, pady=(5, 2))
        self.cert_selector_var = tk.StringVar(value="Root CA Certificate")
        self.cert_selector_cb = ttk.Combobox(self.tab_certs, textvariable=self.cert_selector_var, values=["Root CA Certificate", "Client Certificate", "Private Key File"], state="readonly")
        self.cert_selector_cb.pack(fill="x", padx=5, pady=2)
        self.cert_selector_cb.bind("<<ComboboxSelected>>", self.on_cert_selector_change)

        self.cert_vault_text = scrolledtext.ScrolledText(self.tab_certs, font=("Consolas", 8), wrap="word", bg="#181825", fg="#a6adc8", highlightthickness=0, bd=0)
        self.cert_vault_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        btn_vault_bar = ttk.Frame(self.tab_certs)
        btn_vault_bar.pack(fill="x", pady=5)
        ttk.Button(btn_vault_bar, text="📋 Copy Cert Content to Clipboard", command=self.copy_cert_vault_to_clipboard, style="TButton").pack(fill="x", padx=5)
        self.on_cert_selector_change() # Init loading

        # Tab 3: Ping Live Connection Status
        self.tab_ping = ttk.Frame(self.sidebar_tabs)
        self.sidebar_tabs.add(self.tab_ping, text="📡 Ping Diagnostics")
        
        ttk.Label(self.tab_ping, text="Latency tracking to Device Gateway IP (192.168.4.1):", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=5, pady=(5, 2))
        self.ping_log = scrolledtext.ScrolledText(self.tab_ping, font=("Consolas", 9), wrap="word", bg="#181825", fg="#a6e3a1", state="disabled", highlightthickness=0, bd=0)
        self.ping_log.pack(fill="both", expand=True, padx=5, pady=5)

        # Tab 4: Role Permission Matrix Chart
        self.tab_chart = ttk.Frame(self.sidebar_tabs)
        self.sidebar_tabs.add(self.tab_chart, text="📷 Permission Matrix")
        
        self.chart_label = ttk.Label(self.tab_chart, background=self.BG_MAIN)
        self.chart_label.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.tab_chart.bind("<Configure>", self.resize_chart_image)
        self.original_image = None
        try:
            if os.path.exists("image.png"):
                self.original_image = Image.open("image.png")
            elif os.path.exists("a:\\Coding\\Python\\Api_check_with_gui\\image.png"):
                self.original_image = Image.open("a:\\Coding\\Python\\Api_check_with_gui\\image.png")
        except Exception as e:
            print("Failed to load matrix chart image:", e)

        # Toggle Sidenav Button in bottom frame
        self.sidebar_collapsed = False

        # --- Bottom Status and Export Panel ---
        self.bottom_frame = ttk.Frame(self.root, padding=8)
        self.bottom_frame.pack(fill="x", side="bottom")

        self.status_lbl = ttk.Label(self.bottom_frame, text="Status: Application Ready", font=("Segoe UI", 9, "bold"), foreground=self.COLOR_GREEN)
        self.status_lbl.pack(side="left")

        # Global Session Timer Display
        self.timer_lbl = ttk.Label(self.bottom_frame, text="Token Timer: Not Logged In", font=("Segoe UI", 9, "bold"), foreground=self.FG_MUTED)
        self.timer_lbl.pack(side="right", padx=10)

        # Collapsible button
        self.toggle_side_btn = ttk.Button(self.bottom_frame, text="📖 Hide Sidebar Panel", command=self.toggle_sidebar, style="TButton")
        self.toggle_side_btn.pack(side="right", padx=10)

        # Reports Export
        ttk.Button(self.bottom_frame, text="📄 Export PDF Report", command=self.export_pdf, style="TButton").pack(side="right", padx=5)
        ttk.Button(self.bottom_frame, text="📊 Export Excel Sheet", command=self.export_excel, style="TButton").pack(side="right", padx=5)

    def toggle_sidebar(self):
        """Collapses or expands the right-hand documentation/diagnostics panel"""
        if self.sidebar_collapsed:
            self.main_pane.add(self.right_col, weight=2)
            self.toggle_side_btn.config(text="📖 Hide Sidebar Panel")
            self.sidebar_collapsed = False
        else:
            self.main_pane.forget(self.right_col)
            self.toggle_side_btn.config(text="📖 Show Sidebar Panel")
            self.sidebar_collapsed = True

    # ----------------------------------------------------
    # Background Thread Daemon Loops
    # ----------------------------------------------------
    def wifi_auto_scan_loop(self):
        """Background process scanning Windows interfaces for RMS APs and current link SSID"""
        startupinfo = None
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0 # SW_HIDE

        while self.diag_running:
            if self.simulator_mode_var.get():
                connected_ssid = "RMS-Simulated-Gateway-AP"
                rms_devices = ["RMS-Simulated-Gateway-AP", "RMS-Test-AP-10", "RMS-Developer-AP"]
                self.root.after(0, lambda c=connected_ssid, d=rms_devices: self.update_wifi_ui(c, d))
                time.sleep(5)
                continue

            # 1. Fetch currently connected SSID
            connected_ssid = None
            try:
                out = subprocess.check_output(
                    "netsh wlan show interfaces", 
                    shell=True, 
                    text=True, 
                    errors='ignore', 
                    startupinfo=startupinfo
                )
                for line in out.split('\n'):
                    if "SSID" in line and "BSSID" not in line:
                        parts = line.split(":")
                        if len(parts) > 1:
                            connected_ssid = parts[1].strip()
            except Exception:
                pass

            # 2. Fetch all available Wi-Fi networks in range
            aps_in_range = []
            try:
                out = subprocess.check_output(
                    "netsh wlan show networks", 
                    shell=True, 
                    text=True, 
                    errors='ignore', 
                    startupinfo=startupinfo
                )
                for line in out.split('\n'):
                    if "SSID" in line:
                        parts = line.split(":")
                        if len(parts) > 1:
                            ssid = parts[1].strip()
                            if ssid and ssid not in aps_in_range:
                                aps_in_range.append(ssid)
            except Exception:
                pass

            # Filter for RMS specific devices (SSID starting with RMS-)
            rms_devices = [ssid for ssid in aps_in_range if ssid.upper().startswith("RMS-")]
            
            # Safety checks in case no device detected
            if not rms_devices:
                rms_devices = ["No RMS APs Found"]

            # Update UI on main thread safely
            self.root.after(0, lambda c=connected_ssid, d=rms_devices: self.update_wifi_ui(c, d))
            time.sleep(5)

    def update_wifi_ui(self, connected_ssid, rms_devices):
        """Update interface labels and combobox listings with scanned Wi-Fi information"""
        self.scanned_ssids = rms_devices
        self.ap_select_cb.config(values=self.scanned_ssids)

        if connected_ssid:
            self.current_ssid = connected_ssid
            is_rms = connected_ssid.upper().startswith("RMS-")
            self.step_states[1] = is_rms

            if is_rms:
                self.wifi_status_var.set(f"📡 AP Link: Connected ({connected_ssid})")
                self.wifi_badge.config(foreground=self.COLOR_GREEN)
                self.s1_badge.config(text="🟢 Linked", foreground=self.COLOR_GREEN)
                
                # Auto populate step 2 username if user hasn't typed anything else
                if self.ap_name_var.get() == "RMS-2088" or self.ap_name_var.get() == "":
                    self.ap_name_var.set(connected_ssid)
            else:
                self.wifi_status_var.set(f"📡 AP Link: Alternate SSID ({connected_ssid})")
                self.wifi_badge.config(foreground=self.COLOR_YELLOW)
                self.s1_badge.config(text="🟡 Wrong AP", foreground=self.COLOR_YELLOW)
        else:
            self.current_ssid = None
            self.step_states[1] = False
            self.wifi_status_var.set("📡 AP Link: Disconnected")
            self.wifi_badge.config(foreground=self.COLOR_RED)
            self.s1_badge.config(text="🔴 Disconnected", foreground=self.COLOR_RED)

    def trigger_wifi_scan(self):
        """Manually trigger immediate thread callback scan"""
        self.wifi_status_var.set("📡 AP Link: Scanning Networks...")
        self.wifi_badge.config(foreground=self.COLOR_YELLOW)
        threading.Thread(target=self.wifi_auto_scan_loop, daemon=True).start()

    def on_ap_combo_selected(self, event=None):
        """Pre-populate the authentication name block with dropdown selection"""
        selected_ap = self.ap_select_var.get()
        if selected_ap and selected_ap != "No RMS APs Found":
            self.ap_name_var.set(selected_ap)
            self.update_status(f"Selected Target Device AP: {selected_ap}")

    def ping_diagnostics_loop(self):
        """Daemon thread checking link integrity to standard AP IP (192.168.4.1)"""
        startupinfo = None
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0

        while self.diag_running:
            if self.simulator_mode_var.get():
                is_alive = True
                latency = f"{(int(time.time()) % 3) + 1}ms"
                log_str = f"[{time.strftime('%H:%M:%S')}] IP: 192.168.4.1 -> Reply Success (Time={latency})\n"
                self.ping_history.append(log_str)
                if len(self.ping_history) > 50:
                    self.ping_history.pop(0)
                badge_text = f"🔗 Ping Device: Simulated ({latency})"
                self.root.after(0, lambda b=badge_text, f=self.COLOR_GREEN: self.update_ping_ui(b, f))
                time.sleep(2.5)
                continue

            is_alive = False
            latency = "N/A"

            # Execute single fast ping command
            try:
                cmd = "ping -n 1 -w 800 192.168.4.1" if os.name == 'nt' else "ping -c 1 -W 1 192.168.4.1"
                out = subprocess.check_output(
                    cmd, 
                    shell=True, 
                    text=True, 
                    errors='ignore', 
                    startupinfo=startupinfo
                )
                
                is_alive = True
                # Parse latency value
                if "time=" in out:
                    parts = out.split("time=")
                    if len(parts) > 1:
                        latency = parts[1].split("ms")[0].strip() + "ms"
                elif "time<" in out:
                    latency = "<1ms"
            except Exception:
                is_alive = False

            # Update statistics tracking
            timestamp = time.strftime("%H:%M:%S")
            log_str = f"[{timestamp}] IP: 192.168.4.1 -> "
            if is_alive:
                log_str += f"Reply Success (Time={latency})\n"
                badge_text = f"🔗 Ping Device: Online ({latency})"
                badge_fg = self.COLOR_GREEN
            else:
                log_str += "Request Time Out / Destination Unreachable\n"
                badge_text = "🔗 Ping Device: Offline"
                badge_fg = self.COLOR_RED

            self.ping_history.append(log_str)
            if len(self.ping_history) > 50:
                self.ping_history.pop(0)

            self.root.after(0, lambda b=badge_text, f=badge_fg: self.update_ping_ui(b, f))
            time.sleep(2.5)

    def update_ping_ui(self, badge_text, badge_fg):
        """Update live ping status labels and tab logging panel"""
        self.ping_status_var.set(badge_text)
        self.ping_badge.config(foreground=badge_fg)

        self.ping_log.config(state="normal")
        self.ping_log.delete("1.0", tk.END)
        self.ping_log.insert(tk.END, "".join(reversed(self.ping_history)))
        self.ping_log.config(state="disabled")

    # ----------------------------------------------------
    # Gating & Verification Gates Core Logic
    # ----------------------------------------------------
    def check_prerequisites(self, step_number):
        """Ensure previous logical steps are complete before proceeding"""
        # Step 2: Login Check
        if step_number == 2:
            # S1 Warning gate (SSID Check)
            if not self.step_states[1]:
                ans = messagebox.askyesno(
                    "Prerequisite Warning",
                    f"Warning: Your PC is currently connected to SSID: {self.current_ssid or 'None'}.\n"
                    f"This does not match an RMS AP Gateway structure (starts with 'RMS-').\n\n"
                    "You may not be able to connect to 192.168.4.1.\n"
                    "Do you want to ignore this and attempt login anyway?"
                )
                return ans
            return True

        # Steps 3 to 6: Requires Session Token Authentication (Step 2)
        token = self.session_token_var.get().strip()
        if not token:
            messagebox.showerror(
                "Access Gate Blocked",
                "Prerequisite Failure: Authentication token is missing.\n\n"
                "Please execute Step 2 (Obtain Access Token) by clicking the 'Login' button first."
            )
            return False

        if self.session_expire_time and time.time() >= self.session_expire_time:
            messagebox.showerror(
                "Access Gate Blocked",
                "Prerequisite Failure: Authentication session has EXPIRED.\n\n"
                "Please fetch a new token in Step 2 to reactivate communications."
            )
            return False

        # Additional structural validations
        if step_number == 5:
            # Checking if configs and certs were done
            if not (self.step_states[3] and self.step_states[4]):
                # Warn user but let them proceed
                return messagebox.askyesno(
                    "Workflow Warning",
                    "Certificates (Step 3) or Settings parameters (Step 4) have not been written in this session.\n\n"
                    "Verifying broker links now may check stale parameters.\n"
                    "Do you wish to run the broker check anyway?"
                )
        return True

    # ----------------------------------------------------
    # Step Buttons Handler Callbacks
    # ----------------------------------------------------
    def run_step_2_login(self):
        """Process login and extract cookie session token"""
        if not self.check_prerequisites(2):
            return

        self.update_status("Authenticating Role Login (POST /api/login)...")
        ap_username = self.ap_name_var.get().strip()
        password = self.password_var.get().strip()
        
        login_url = "http://192.168.4.1/api/login"
        payload = {"username": ap_username, "password": password}
        headers = {"Content-Type": "application/json"}

        self.update_request_preview("POST", login_url, headers, payload)

        def _thread():
            try:
                resp = self.perform_request("POST", login_url, json_payload=payload, headers=headers, timeout=4)
                if resp.status_code == 200:
                    token = ""
                    try:
                        data = resp.json()
                        token = data.get("sessionToken") or data.get("session") or ""
                    except Exception:
                        pass

                    if not token:
                        token = resp.cookies.get("session") or "3fa9a12da510c17d"

                    self.root.after(0, lambda t=token: self.on_login_success(t))
                else:
                    self.root.after(0, lambda: self.on_login_failure(f"HTTP {resp.status_code} - Unauthorized"))
            except Exception as ex:
                self.root.after(0, lambda err=str(ex): self.on_login_failure(f"Offline: {err}"))

        threading.Thread(target=_thread, daemon=True).start()

    def on_login_success(self, token):
        self.session_token_var.set(token)
        self.step_states[2] = True
        self.s2_badge.config(text="🟢 Session Active", foreground=self.COLOR_GREEN)
        self.update_status("Login completed successfully!")
        
        # Start countdown
        self.start_1min_timer()

        # Update Grid statuses from pending/unauthorized to Ready
        self.populate_table()

    def on_login_failure(self, err_msg):
        self.session_token_var.set("")
        self.step_states[2] = False
        self.s2_badge.config(text="🔴 Login Failed", foreground=self.COLOR_RED)
        self.update_status(f"Login failed: {err_msg}")
        messagebox.showerror("Authentication Failure", f"Failed to login to Gateway:\n{err_msg}")

    def run_step_3_certs(self):
        """Upload ROOT CA, KEY, and CLIENT PEM certs files"""
        if not self.check_prerequisites(3):
            return

        token = self.session_token_var.get().strip()
        headers = {"Cookie": f"session={token}", "Content-Type": "text/plain"}
        self.update_status("Uploading certificates sequence starting...")

        def _thread():
            success = True
            errors = []
            
            # Root CA
            try:
                r1 = self.perform_request("POST", "http://192.168.4.1/write.html?filename=rootCA.pem", raw_payload=ROOT_CA_PEM, headers=headers, timeout=5)
                if r1.status_code != 200:
                    success = False
                    errors.append(f"RootCA: HTTP {r1.status_code}")
            except Exception as e:
                success = False
                errors.append(f"RootCA: {str(e)}")

            # Key
            try:
                r2 = self.perform_request("POST", "http://192.168.4.1/write.html?filename=key.pem", raw_payload=KEY_PEM, headers=headers, timeout=5)
                if r2.status_code != 200:
                    success = False
                    errors.append(f"Key: HTTP {r2.status_code}")
            except Exception as e:
                success = False
                errors.append(f"Key: {str(e)}")

            # Client Cert
            try:
                r3 = self.perform_request("POST", "http://192.168.4.1/write.html?filename=client.pem", raw_payload=CLIENT_PEM, headers=headers, timeout=5)
                if r3.status_code != 200:
                    success = False
                    errors.append(f"Client Cert: HTTP {r3.status_code}")
            except Exception as e:
                success = False
                errors.append(f"Client Cert: {str(e)}")

            if success:
                self.root.after(0, self.on_certs_success)
            else:
                self.root.after(0, lambda errs="; ".join(errors): self.on_certs_failure(errs))

        threading.Thread(target=_thread, daemon=True).start()

    def on_certs_success(self):
        self.step_states[3] = True
        self.s3_badge.config(text="🟢 Uploaded", foreground=self.COLOR_GREEN)
        self.update_status("Certificates sequence uploaded successfully!")
        
        # Log results to Grid
        self.update_grid_status(18, "200 OK")
        self.update_grid_status(19, "200 OK")
        self.update_grid_status(20, "200 OK")
        messagebox.showinfo("Success", "Security credentials uploaded successfully.")

    def on_certs_failure(self, errors):
        self.step_states[3] = False
        self.s3_badge.config(text="🔴 Upload Failed", foreground=self.COLOR_RED)
        self.update_status(f"Cert upload failed: {errors}")
        messagebox.showerror("Upload Failure", f"Failed uploading credentials:\n{errors}")

    def run_step_4_configs(self):
        """Configure cellular ISP values and MQTT server targets"""
        if not self.check_prerequisites(4):
            return

        token = self.session_token_var.get().strip()
        headers = {"Cookie": f"session={token}", "Content-Type": "application/json"}
        self.update_status("Pushing configurations (ISP, Server URLs, Modbus)...")

        # Configurations payloads
        isp_payload = {"apn": "airtelgprs.com", "apn2": "airtelgprs.com", "current_sim": "1"}
        server_payload = {
            "server_url": "rms.iotscada-pmsg.com", "server_port": 8883, "solution_type": "ongridrooftop",
            "client_id": f"d:{self.ap_name_var.get()}$ongridrooftop$510017",
            "username": f"{self.ap_name_var.get()}$ongridrooftop$510017", "password": "password_secret"
        }

        def _thread():
            success = True
            errors = []

            # 1. APN settings
            try:
                r1 = self.perform_request("POST", "http://192.168.4.1/api/config/isp", json_payload=isp_payload, headers=headers, timeout=4)
                if r1.status_code != 200:
                    success = False
                    errors.append(f"APN Settings: HTTP {r1.status_code}")
            except Exception as e:
                success = False
                errors.append(f"APN Settings: {str(e)}")

            # 2. Remote Server
            try:
                r2 = self.perform_request("POST", "http://192.168.4.1/api/config/remote-server", json_payload=server_payload, headers=headers, timeout=4)
                if r2.status_code != 200:
                    success = False
                    errors.append(f"Remote Server: HTTP {r2.status_code}")
            except Exception as e:
                success = False
                errors.append(f"Remote Server: {str(e)}")

            if success:
                self.root.after(0, self.on_configs_success)
            else:
                self.root.after(0, lambda err="; ".join(errors): self.on_configs_failure(err))

        threading.Thread(target=_thread, daemon=True).start()

    def on_configs_success(self):
        self.step_states[4] = True
        self.s4_badge.config(text="🟢 Configured", foreground=self.COLOR_GREEN)
        self.update_status("APN and target server configs written successfully!")
        self.update_grid_status(5, "200 OK")
        self.update_grid_status(7, "200 OK")
        messagebox.showinfo("Success", "Core system settings parameters pushed successfully.")

    def on_configs_failure(self, err):
        self.step_states[4] = False
        self.s4_badge.config(text="🔴 Write Failed", foreground=self.COLOR_RED)
        self.update_status(f"Config write failed: {err}")
        messagebox.showerror("Write Failure", f"Failed uploading parameters:\n{err}")

    def run_step_5_broker(self):
        """Trigger remote server broker connection and verify its telemetry status"""
        if not self.check_prerequisites(5):
            return

        token = self.session_token_var.get().strip()
        headers = {"Cookie": f"session={token}", "Content-Type": "application/json"}
        self.update_status("Triggering secure broker linking...")

        def _thread():
            success = True
            status_desc = "Unknown"
            
            # 1. Trigger linking call
            try:
                r1 = self.perform_request("POST", "http://192.168.4.1/api/device/broker/connect", json_payload={"action": "connect"}, headers=headers, timeout=4)
                if r1.status_code != 200:
                    success = False
            except Exception:
                success = False

            # Wait briefly for link sequence
            time.sleep(2)

            # 2. Check telemetry status
            try:
                r2 = self.perform_request("GET", "http://192.168.4.1/api/device/broker/status", headers=headers, timeout=4)
                if r2.status_code == 200:
                    status_desc = r2.text
                else:
                    success = False
            except Exception:
                success = False

            if success:
                self.root.after(0, lambda desc=status_desc: self.on_broker_success(desc))
            else:
                self.root.after(0, self.on_broker_failure)

        threading.Thread(target=_thread, daemon=True).start()

    def on_broker_success(self, status):
        self.step_states[5] = True
        self.s5_badge.config(text="🟢 Linked", foreground=self.COLOR_GREEN)
        self.update_status("Telemetry Broker connection triggered and checked!")
        self.update_grid_status(10, "200 OK")
        self.update_grid_status(11, "200 OK")
        
        # Display response
        self.output_text.config(state="normal")
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, f"=== MQTT BROKER TELEMETRY RESPONSE ===\n{status}")
        
        messagebox.showinfo("Broker Success", f"Secure link telemetry verification:\n{status[:200]}")

    def on_broker_failure(self):
        self.step_states[5] = False
        self.s5_badge.config(text="🔴 Offline", foreground=self.COLOR_RED)
        self.update_status("Failed to verify secure broker link.")
        messagebox.showerror("Broker Failure", "Failed checking active MQTT Broker status link.")

    def run_step_6_reboot(self):
        """Initiate device soft restart command"""
        if not self.check_prerequisites(6):
            return

        ans = messagebox.askyesno("Confirm Reboot", "Are you sure you want to reboot the gateway hardware now?")
        if not ans:
            return

        token = self.session_token_var.get().strip()
        headers = {"Cookie": f"session={token}", "Content-Type": "application/json"}
        self.update_status("Sending restart command (POST /restart)...")

        def _thread():
            try:
                resp = self.perform_request("POST", "http://192.168.4.1/restart", json_payload={"action": "reboot"}, headers=headers, timeout=4)
                if resp.status_code == 200:
                    self.root.after(0, self.on_reboot_success)
                else:
                    self.root.after(0, lambda: self.on_reboot_failure(f"HTTP {resp.status_code}"))
            except Exception as e:
                self.root.after(0, lambda err=str(e): self.on_reboot_failure(err))

        threading.Thread(target=_thread, daemon=True).start()

    def on_reboot_success(self):
        self.step_states[6] = True
        self.s6_badge.config(text="🟢 Reboot Sent", foreground=self.COLOR_GREEN)
        self.update_status("Reboot signal sent. Gateway restarting!")
        self.update_grid_status(21, "200 OK")
        
        # Clear token as session is wiped on reboot
        self.clear_session()
        messagebox.showinfo("Device Reboot", "Gateway rebooted successfully. Connection will drop temporarily.")

    def on_reboot_failure(self, err):
        self.step_states[6] = False
        self.s6_badge.config(text="🔴 Failed", foreground=self.COLOR_RED)
        self.update_status(f"Reboot failure: {err}")
        messagebox.showerror("Reboot Failed", f"Could not reboot device:\n{err}")

    # ----------------------------------------------------
    # Grid and Tables Operations (Selection, Load, and Click)
    # ----------------------------------------------------
    def populate_table(self):
        """Draw endpoint rows on Treeview Grid with active authorization statuses"""
        self.tree.delete(*self.tree.get_children())
        selected_role = self.role_var.get()
        token_active = bool(self.session_token_var.get().strip())

        for api in API_ENDPOINTS:
            permissions = api["roles"].get(selected_role, [])
            read_access = "YES" if "Read" in permissions else "-"
            write_access = "YES" if "Write" in permissions else "-"

            if permissions:
                status = "Ready" if token_active else "Pending Token"
                tick = "☑"
            else:
                status = "Unauthorized"
                tick = "☐"

            self.tree.insert("", "end", iid=str(api["sr"]), values=(
                tick, api["sr"], api["name"], api["method"], api["url"], read_access, write_access, status
            ))

    def on_tree_click(self, event):
        """Row check/uncheck clicking events handler"""
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == "#1":  # Column index for check marks
                item_id = self.tree.identify_row(event.y)
                if item_id:
                    vals = list(self.tree.item(item_id, "values"))
                    vals[0] = "☑" if vals[0] == "☐" else "☐"
                    self.tree.item(item_id, values=vals)

    def on_select_tree_item(self, event):
        """Selected row details output updating to preview panel and docs references tab"""
        selected_item = self.tree.selection()
        if not selected_item:
            return
        sr_id = int(selected_item[0])

        # 1. Update Endpoint Document Reference Tab
        self.show_api_documentation(sr_id)

        # 2. Update Request Preview Box
        for api in API_ENDPOINTS:
            if api["sr"] == sr_id:
                token = self.session_token_var.get().strip()
                headers = {"Cookie": f"session={token}"} if token else {}
                
                payload = api.get("payload")
                if api["name"] == "Authentication API (Write API Login)":
                    payload = {"username": self.ap_name_var.get().strip(), "password": self.password_var.get().strip()}

                self.update_request_preview(api["method"], api["url"], headers, payload)
                break

        # 3. Update Output Area if matching result already executed
        found_result = False
        for res in self.execution_results:
            if res["sr"] == sr_id:
                self.display_output(res)
                found_result = True
                break

        if not found_result:
            self.output_text.config(state="normal")
            self.output_text.delete("1.0", tk.END)
            self.output_text.insert(tk.END, f"API Sr. #{sr_id} selected.\nNo response received in this session yet. Press 'Run Checked APIs' to trigger.")

    def update_grid_status(self, sr_id, status_text):
        """Utility to quickly find and change status value inside table grid rows"""
        try:
            item_id = str(sr_id)
            if self.tree.exists(item_id):
                vals = list(self.tree.item(item_id, "values"))
                vals[7] = status_text
                self.tree.item(item_id, values=vals)
        except Exception:
            pass

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

    def on_role_change(self, event=None):
        role = self.role_var.get()
        self.password_var.set(USER_CREDENTIALS.get(role, ""))
        self.populate_table()
        self.update_status(f"Switched role focus to: {role}")

    # ----------------------------------------------------
    # Token Expiration 1-minute Session Timer
    # ----------------------------------------------------
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
                self.timer_lbl.config(text=f"Session Expires: {remaining}s", foreground=self.COLOR_GREEN)
                self.s2_timer_lbl.config(text=f"Session Active: {remaining}s remaining", foreground=self.COLOR_GREEN)
                self.root.after(1000, self.update_timer_loop)
            else:
                self.clear_session()
        else:
            self.timer_lbl.config(text="Session Timer: Off", foreground=self.FG_MUTED)

    def clear_session(self):
        self.session_expire_time = None
        self.timer_running = False
        self.session_token_var.set("")
        self.step_states[2] = False
        self.s2_badge.config(text="🔴 Expired", foreground=self.COLOR_RED)
        self.timer_lbl.config(text="Session Expired", foreground=self.COLOR_RED)
        self.s2_timer_lbl.config(text="Token Expired. Please authenticate again.", foreground=self.COLOR_RED)
        self.populate_table()
        self.update_status("Authentication session expired and cleared.")

    # ----------------------------------------------------
    # Dynamic Documentation Viewer Tab
    # ----------------------------------------------------
    def show_api_documentation(self, sr_id):
        """Generate dynamic description document text in Right sidebar reference page"""
        self.doc_scroll.config(state="normal")
        self.doc_scroll.delete("1.0", tk.END)

        # Lookup API details
        api = None
        for a in API_ENDPOINTS:
            if a["sr"] == sr_id:
                api = a
                break

        if not api:
            self.doc_scroll.insert(tk.END, "Select an API endpoint to view documentation details.")
            self.doc_scroll.config(state="disabled")
            return

        roles_text = ""
        for role, perms in api["roles"].items():
            roles_text += f" - {role}: {', '.join(perms)}\n"
        if not roles_text:
            roles_text = " - General public access (No session required)\n"

        doc_content = (
            f"ENDPOINT REF: Sr #{api['sr']}\n"
            f"====================================\n"
            f"NAME: {api['name']}\n\n"
            f"METHOD: {api['method']}\n"
            f"TARGET URL: {api['url']}\n\n"
            f"DESCRIPTION:\n"
            f"{API_DESCRIPTIONS.get(sr_id, 'No additional descriptions available.')}\n\n"
            f"AUTHORIZATION MATRIX PERMISSIONS:\n"
            f"{roles_text}\n"
            f"DEFAULT JSON BODY PAYLOAD SCHEMA:\n"
        )
        if api["payload"]:
            doc_content += json.dumps(api["payload"], indent=2)
        else:
            doc_content += "None (Empty Body / GET Param Query)"

        self.doc_scroll.insert(tk.END, doc_content)
        self.doc_scroll.config(state="disabled")

    # ----------------------------------------------------
    # Certificates Tab Handler
    # ----------------------------------------------------
    def on_cert_selector_change(self, event=None):
        """Update vault textarea based on combo box select"""
        choice = self.cert_selector_var.get()
        self.cert_vault_text.delete("1.0", tk.END)
        if choice == "Root CA Certificate":
            self.cert_vault_text.insert(tk.END, ROOT_CA_PEM)
        elif choice == "Client Certificate":
            self.cert_vault_text.insert(tk.END, CLIENT_PEM)
        else:
            self.cert_vault_text.insert(tk.END, KEY_PEM)

    def copy_cert_vault_to_clipboard(self):
        content = self.cert_vault_text.get("1.0", tk.END).strip()
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.update_status(f"Copied {self.cert_selector_var.get()} to system clipboard!")

    def resize_chart_image(self, event=None):
        if not self.original_image:
            return
        
        # Get frame dimensions
        width = event.width if event else self.tab_chart.winfo_width()
        height = event.height if event else self.tab_chart.winfo_height()
        
        if width < 50 or height < 50:
            return
        
        # Calculate aspect ratio
        img_w, img_h = self.original_image.size
        aspect = img_w / img_h
        
        # Resize to fit width while keeping aspect ratio
        new_w = width - 10
        new_h = int(new_w / aspect)
        
        # If the height is too big for the frame, scale down to fit height instead
        if new_h > height - 10:
            new_h = height - 10
            new_w = int(new_h * aspect)
            
        if new_w <= 0 or new_h <= 0:
            return
            
        try:
            resized = self.original_image.resize((new_w, new_h), Image.Resampling.LANCZOS)
            self.tk_image = ImageTk.PhotoImage(resized)
            self.chart_label.config(image=self.tk_image)
        except Exception:
            pass

    # ----------------------------------------------------
    # Outgoing Request Preview Box Update
    # ----------------------------------------------------
    def update_request_preview(self, method, url, headers, payload):
        self.preview_text.config(state="normal")
        self.preview_text.delete("1.0", tk.END)
        
        preview_str = f"🚀 [{method}] {url}\n"
        preview_str += f"HEADERS: {json.dumps(headers, indent=2)}\n"
        
        if payload is not None:
            if isinstance(payload, dict):
                preview_str += f"JSON BODY:\n{json.dumps(payload, indent=2)}"
            else:
                preview_str += f"RAW TEXT BODY:\n{str(payload)}"
        else:
            preview_str += "BODY: Empty/None"

        self.preview_text.insert(tk.END, preview_str)
        self.preview_text.config(state="normal")

    def display_output(self, res):
        self.output_text.config(state="normal")
        self.output_text.delete("1.0", tk.END)
        output_data = (
            f"Device SSID   : {self.ap_name_var.get()}\n"
            f"API Target    : {res['name']}\n"
            f"Request URL   : {res['url']}\n"
            f"Method Action : {res['method']}\n"
            f"Role Privileges: Read={res['read_acc']} | Write={res['write_acc']}\n"
            f"Response Status: {res['status']}\n\n"
            f"================ RESPONSE BODY ================\n"
            f"{res['body']}"
        )
        self.output_text.insert(tk.END, output_data)
        self.output_text.see(tk.END)

    def update_status(self, text):
        self.root.after(0, lambda: self.status_lbl.config(text=f"Status: {text}"))

    # ----------------------------------------------------
    # Custom Testing Pipeline (Run Selected APIs Checked in Grid)
    # ----------------------------------------------------
    def start_process(self):
        """Global batch testing runner for grid-checked endpoints"""
        token = self.session_token_var.get().strip()
        if not token:
            messagebox.showwarning("Missing Credentials", "Prerequisite missing: obtain an active Session Token (Step 2) first.")
            return

        if self.session_expire_time and time.time() >= self.session_expire_time:
            messagebox.showwarning("Session Expired", "Authentication timer has expired. Re-login (Step 2) to start testing.")
            return

        selected_sr_list = []
        for item_id in self.tree.get_children():
            vals = self.tree.item(item_id, "values")
            if vals[0] == "☑":
                selected_sr_list.append(int(vals[1]))

        if not selected_sr_list:
            messagebox.showwarning("No Selection", "Please check at least one API checkbox in the grid to run test.")
            return

        self.cancel_event.clear()
        self.run_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.execution_results.clear()

        self.update_status(f"Starting testing pipeline for {len(selected_sr_list)} checked APIs...")
        threading.Thread(target=self.run_execution_pipeline, args=(selected_sr_list,), daemon=True).start()

    def cancel_process(self):
        self.cancel_event.set()
        self.update_status("Cancelling background testing execution...")
        self.cancel_btn.config(state="disabled")

    def run_execution_pipeline(self, selected_sr_list):
        role = self.role_var.get()
        apis_to_run = [a for a in API_ENDPOINTS if a["sr"] in selected_sr_list]

        for idx, api in enumerate(apis_to_run):
            if self.cancel_event.is_set():
                break

            # Confirm token is valid before sending request
            if self.session_expire_time and time.time() >= self.session_expire_time:
                self.update_status("Wired connection token expired during tests!")
                break

            token = self.session_token_var.get().strip()
            sr_id = str(api["sr"])

            permissions = api["roles"].get(role, [])
            read_acc = "YES" if "Read" in permissions else "-"
            write_acc = "YES" if "Write" in permissions else "-"

            self.update_status(f"Executing API {idx+1}/{len(apis_to_run)}: {api['name']}...")
            self.tree.item(sr_id, values=("☑", api["sr"], api["name"], api["method"], api["url"], read_acc, write_acc, "Calling..."))

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
                    r = self.perform_request("GET", api["url"], headers=headers, timeout=5)
                else:
                    if isinstance(payload, dict):
                        r = self.perform_request("POST", api["url"], json_payload=payload, headers=headers, timeout=5)
                    else:
                        r = self.perform_request("POST", api["url"], raw_payload=payload, headers=headers, timeout=5)

                status_str = f"{r.status_code} {r.reason}"
                body_output = r.text
            except Exception as ex:
                status_str = "Offline / Link Error"
                body_output = f"Connection failure detail: {str(ex)}"

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

            # Instantly display output on screen
            self.root.after(0, lambda res=res_entry: self.display_output(res))

            # Trigger step workflow states if core API was tested
            if api["sr"] in [18, 19, 20] and status_str.startswith("200"):
                self.step_states[3] = True
                self.s3_badge.config(text="🟢 Uploaded", foreground=self.COLOR_GREEN)
            elif api["sr"] in [5, 7] and status_str.startswith("200"):
                self.step_states[4] = True
                self.s4_badge.config(text="🟢 Configured", foreground=self.COLOR_GREEN)
            elif api["sr"] == 11 and status_str.startswith("200"):
                self.step_states[5] = True
                self.s5_badge.config(text="🟢 Linked", foreground=self.COLOR_GREEN)

            # Delay spacing between requests
            if idx < len(apis_to_run) - 1:
                for remaining in range(5, 0, -1):
                    if self.cancel_event.is_set() or (self.session_expire_time and time.time() >= self.session_expire_time):
                        break
                    self.update_status(f"Waiting delay... ({remaining}s remaining before next API)")
                    time.sleep(1)

        if not self.cancel_event.is_set() and (not self.session_expire_time or time.time() < self.session_expire_time):
            self.update_status("Batch testing pipeline completed successfully!")

        self.root.after(0, lambda: self.run_btn.config(state="normal"))
        self.root.after(0, lambda: self.cancel_btn.config(state="disabled"))

    # ----------------------------------------------------
    # PDF & Excel Report Exports
    # ----------------------------------------------------
    def export_excel(self):
        if not self.execution_results:
            messagebox.showinfo("Export Warning", "No results generated in this session to export. Run some tests first.")
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "API Verification Results"

        ws.append(["Gateway AP SSID Name", self.ap_name_var.get()])
        ws.append(["Active User Role", self.role_var.get()])
        ws.append(["Session Access Token", self.session_token_var.get()])
        ws.append(["Diagnostic Date", time.strftime("%Y-%m-%d %H:%M:%S")])
        ws.append([])
        ws.append(["Sr.", "Target API Name", "Method", "URL Link", "Read OK", "Write OK", "HTTP Status Response", "Short Response Preview"])

        for item in self.execution_results:
            ws.append([item["sr"], item["name"], item["method"], item["url"], item["read_acc"], item["write_acc"], item["status"], str(item["body"])[:300]])

        filename = f"{self.ap_name_var.get()}_API_Report.xlsx"
        try:
            wb.save(filename)
            self.update_status(f"Excel report saved: {filename}")
            messagebox.showinfo("Export Complete", f"Excel report worksheet saved as:\n{filename}")
        except Exception as e:
            messagebox.showerror("Export Failure", f"Failed saving Excel file:\n{str(e)}")

    def export_pdf(self):
        if not self.execution_results:
            messagebox.showinfo("Export Warning", "No results generated in this session to export. Run some tests first.")
            return

        filename = f"{self.ap_name_var.get()}_API_Report.pdf"
        try:
            doc = SimpleDocTemplate(filename, pagesize=letter)
            styles = getSampleStyleSheet()
            story = []

            # Style Customization
            title_style = ParagraphStyle("TitleStyle", parent=styles["Heading1"], fontSize=16, leading=20, textColor=colors.HexColor("#89b4fa"))
            story.append(Paragraph("RMS Device API Test Summary Report", title_style))
            story.append(Spacer(1, 10))

            info_text = f"<b>Device SSID Network:</b> {self.ap_name_var.get()}<br/>" \
                        f"<b>Active Operative Role:</b> {self.role_var.get()}<br/>" \
                        f"<b>Session Access Token:</b> {self.session_token_var.get() or 'None'}<br/>" \
                        f"<b>Report Timestamp:</b> {time.strftime('%Y-%m-%d %H:%M:%S')}"
            story.append(Paragraph(info_text, styles["Normal"]))
            story.append(Spacer(1, 15))

            table_data = [["Sr.", "API Name", "Method", "Read", "Write", "HTTP Status Code"]]
            for item in self.execution_results:
                table_data.append([str(item["sr"]), item["name"], item["method"], item["read_acc"], item["write_acc"], item["status"]])

            t = Table(table_data, colWidths=[30, 220, 50, 50, 50, 100])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#313244")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#cdd6f4")),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#45475a")),
            ]))

            story.append(t)
            doc.build(story)
            self.update_status(f"PDF report saved: {filename}")
            messagebox.showinfo("Export Complete", f"PDF report document saved as:\n{filename}")
        except Exception as e:
            messagebox.showerror("Export Failure", f"Failed saving PDF file:\n{str(e)}")

    def destroy(self):
        """Cleanup daemon execution loops safely before exiting"""
        self.diag_running = False
        self.root.destroy()

if __name__ == "__main__":
    self_root = tk.Tk()
    # Apply standard DPI scaling if available on Windows
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = ModernRMSTesterApp(self_root)
    # Handle window close event cleanly
    self_root.protocol("WM_DELETE_WINDOW", app.destroy)
    self_root.mainloop()
