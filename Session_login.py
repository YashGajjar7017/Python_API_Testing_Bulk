import tkinter as tk
from tkinter import ttk, messagebox
import requests
import threading
import time

# Configuration & Role Mapping
LOGIN_URL = "http://192.168.4.1/api/login"
PASSWORDS = {
    "viewer": "viewer_001",
    "systemAdmin": "sysadmin_001",
    "operator": "operator_001",
    "securityAdmin": "secadmin_001"
}

class SessionAuthApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RMS Session Auth Manager")
        self.root.geometry("480x420")
        self.root.resizable(False, False)
        
        self.timer_running = False
        self.remaining_seconds = 60
        self.active_session_token = None

        self._build_ui()

    def _build_ui(self):
        # Header / Title
        title_label = ttk.Label(self.root, text="RMS Authentication System", font=("Helvetica", 14, "bold"))
        title_label.pack(pady=10)

        # Input Frame
        frame = ttk.Frame(self.root, padding=15)
        frame.pack(fill="x")

        # Username Input
        ttk.Label(frame, text="Username:").grid(row=0, column=0, sticky="w", pady=5)
        self.username_entry = ttk.Entry(frame, width=30)
        self.username_entry.insert(0, "RMS-")
        self.username_entry.grid(row=0, column=1, pady=5)

        # Role Dropdown Selection
        ttk.Label(frame, text="Select Role:").grid(row=1, column=0, sticky="w", pady=5)
        self.role_var = tk.StringVar(value="viewer")
        self.role_dropdown = ttk.Combobox(
            frame, 
            textvariable=self.role_var, 
            values=list(PASSWORDS.keys()), 
            state="readonly", 
            width=28
        )
        self.role_dropdown.grid(row=1, column=1, pady=5)

        # Login Button
        self.login_btn = ttk.Button(self.root, text="Login & Get Session", command=self.handle_login)
        self.login_btn.pack(pady=10)

        # Divider
        ttk.Separator(self.root, orient="horizontal").pack(fill="x", padx=15, pady=5)

        # Session Display Area (Cookie Information)
        session_frame = ttk.LabelFrame(self.root, text=" Session Details ", padding=10)
        session_frame.pack(fill="x", padx=15, pady=5)

        self.session_info_label = ttk.Label(
            session_frame, 
            text="No active session.", 
            font=("Courier", 10), 
            wraplength=420, 
            justify="left"
        )
        self.session_info_label.pack(fill="x", pady=5)

        # Timer Display
        self.timer_label = ttk.Label(self.root, text="Timer: Idle", font=("Helvetica", 11, "italic"), foreground="gray")
        self.timer_label.pack(pady=10)

    def handle_login(self):
        username = self.username_entry.get().strip()
        role = self.role_var.get()
        password = PASSWORDS.get(role)

        if not username:
            messagebox.showerror("Error", "Please enter a valid username.")
            return

        payload = {
            "username": username,
            "password": password
        }

        try:
            # Send HTTP POST request to login endpoint
            response = requests.post(LOGIN_URL, json=payload, timeout=5)
            
            # Extract session cookie from response
            session_value = response.cookies.get("session")

            # Fallback check if session is sent in JSON body instead of Cookie header
            if not session_value and response.headers.get("Content-Type") == "application/json":
                session_value = response.json().get("session")

            if session_value:
                self.active_session_token = session_value
                
                # Copy session token to clipboard automatically
                self.root.clipboard_clear()
                self.root.clipboard_append(session_value)
                
                # Format output matching the cookie table parameters
                display_text = (
                    f"Name:     session\n"
                    f"Value:    {session_value}\n"
                    f"Domain:   192.168.4.1\n"
                    f"Path:     /\n"
                    f"Expires:  Session\n"
                    f"HttpOnly: true\n"
                    f"Secure:   false"
                )
                self.session_info_label.config(text=display_text, foreground="black")
                messagebox.showinfo("Success", "Session token copied to clipboard!")

                # Start 1-minute expiration timer
                self.start_timer(60)
            else:
                messagebox.showwarning("Warning", "Login requested successfully, but no 'session' cookie was received.")

        except requests.exceptions.RequestException as e:
            messagebox.showerror("Connection Error", f"Failed to reach target server:\n{e}")

    def start_timer(self, seconds):
        self.remaining_seconds = seconds
        self.timer_running = True
        self.login_btn.config(state="disabled")
        self._update_timer_loop()

    def _update_timer_loop(self):
        if self.timer_running and self.remaining_seconds > 0:
            self.timer_label.config(
                text=f"Session Expires In: {self.remaining_seconds}s", 
                foreground="red"
            )
            self.remaining_seconds -= 1
            self.root.after(1000, self._update_timer_loop)
        elif self.remaining_seconds <= 0:
            self.clear_session()

    def clear_session(self):
        self.timer_running = False
        self.active_session_token = None
        self.session_info_label.config(text="Session expired and removed.", foreground="gray")
        self.timer_label.config(text="Timer: Expired", foreground="gray")
        self.login_btn.config(state="normal")
        
        # Clear system clipboard if it contains the expired token
        try:
            self.root.clipboard_clear()
        except tk.TclError:
            pass
            
        messagebox.showwarning("Session Expired", "1-minute limit reached. Session token has been cleared.")

if __name__ == "__main__":
    root = tk.Tk()
    app = SessionAuthApp(root)
    root.mainloop()