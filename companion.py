import os
import sys
import logging
import threading
import time
import webbrowser
import urllib.request
import urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
import json
import traceback
import keyring
import winreg
import customtkinter as ctk
from PIL import Image, ImageDraw
import pystray
from supabase import create_client, Client

# --- CONFIGURATION ---
APP_VERSION = "1.0.3"
APP_NAME = "LMU Paddock Companion"
AUTH_URL = "https://lmupaddock.com/app-auth" 
SUPPORT_URL = "https://lmupaddock.com/support"  
UPDATE_URL = "https://lmupaddock.com/version.json"
PORT = 42069
APPDATA_DIR = os.path.join(os.getenv('APPDATA'), 'LMUPaddock')
SETTINGS_FILE = os.path.join(APPDATA_DIR, 'settings.json')
LOCK_FILE = os.path.join(APPDATA_DIR, 'companion.lock')
SUPABASE_URL = "https://qumgjoricnkzpopzevwl.supabase.co"
SUPABASE_KEY = "sb_publishable_MymWCvMB8AYaPNDGLBbRoA_bWWxcbL3"

if not os.path.exists(APPDATA_DIR):
    os.makedirs(APPDATA_DIR)

# --- LOGGING SETUP ---
log_file = os.path.join(APPDATA_DIR, 'companion_debug.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - [%(levelname)s] - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

# --- HELPER FUNCTIONS ---
def check_single_instance():
    if os.path.exists(LOCK_FILE):
        try:
            os.remove(LOCK_FILE)
        except OSError:
            return False
    try:
        f = open(LOCK_FILE, 'w')
        f.write(str(os.getpid()))
        return f
    except Exception:
        return False

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def clean_string(val):
    if isinstance(val, bytes):
        return val.decode('utf-8', errors='ignore').replace('\x00', '').strip()
    return str(val).replace('\x00', '').strip()

def format_laptime(seconds):
    if seconds <= 0.0: return "No time"
    seconds = round(seconds, 3)
    minutes = int(seconds // 60)
    remainder = seconds % 60
    return f"{minutes}:{remainder:06.3f}"

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                return json.load(f)
        except: pass
    return {"minimize_to_tray": True, "has_seen_welcome": False}

def save_settings(settings):
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f)
    except Exception as e:
        logging.error(f"Failed to save settings: {e}")

def is_autostart_enabled():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
        winreg.QueryValueEx(key, APP_NAME)
        winreg.CloseKey(key)
        return True
    except WindowsError:
        return False

def toggle_autostart(enable):
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
        if enable:
            if getattr(sys, 'frozen', False):
                app_path = f'"{sys.executable}"'
            else:
                app_path = f'"{sys.executable}" "{os.path.abspath(__file__)}"'
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, app_path)
        else:
            try: winreg.DeleteValue(key, APP_NAME)
            except FileNotFoundError: pass
        winreg.CloseKey(key)
    except Exception as e:
        logging.error(f"Failed to toggle auto-start: {e}")

# --- TELEMETRY COLLECTOR ENGINE ---
class TelemetryCollector:
    def __init__(self, ui_callback, auth_update_callback, access_token, refresh_token):
        self.ui_callback = ui_callback
        self.auth_update_callback = auth_update_callback
        self.is_running = False
        self.thread = None
        self.supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
        self.auth_subscription = None
        try:
            self.supabase.auth.set_session(access_token, refresh_token)
            self.auth_subscription = self.supabase.auth.on_auth_state_change(self._handle_auth_change)
            
            user_res = self.supabase.auth.get_user()
            self.user_id = user_res.user.id
        except Exception as e:
            self.ui_callback(f"[ERROR] Auth verification failed: {e}")
            self.user_id = None

    def _handle_auth_change(self, event, session):
        if event == "TOKEN_REFRESHED" and session:
            self.auth_update_callback(session.access_token, session.refresh_token)
            self.ui_callback("[SYS] Security tokens rotated and secured.")

    def fetch_garage_setup(self):
        url = "http://localhost:6397/rest/garage/getPlayerGarageData"
        setup = {"bb": 50.0, "tc": 0, "abs": 0, "tc_cut": 0, "tc_slip": 0, "wing": 0}
        try:
            req = urllib.request.urlopen(url, timeout=0.8)
            data = json.loads(req.read())
            for k, v in data.items():
                if not isinstance(v, dict): continue
                if k == "VM_BRAKE_BALANCE":
                    s_val = str(v.get("stringValue", ""))
                    if ":" in s_val:
                        try: setup["bb"] = round(float(s_val.split(":")[0]), 1)
                        except: pass
                elif k == "VM_TRACTIONCONTROLMAP": setup["tc"] = v.get("value", 0)
                elif k == "VM_ANTILOCKBRAKESYSTEMMAP": setup["abs"] = v.get("value", 0)
                elif k == "VM_TRACTIONCONTROLPOWERCUTMAP": setup["tc_cut"] = v.get("value", 0)
                elif k == "VM_TRACTIONCONTROLSLIPANGLEMAP": setup["tc_slip"] = v.get("value", 0)
                elif k == "VM_REAR_WING": setup["wing"] = v.get("value", 0)
            return setup
        except: 
            return None

    def fetch_garage_summary(self):
        """ Fetch both the real car model and the specific track layout name """
        url = "http://localhost:6397/rest/garage/summary"
        summary = {"car": None, "track": None}
        try:
            req = urllib.request.urlopen(url, timeout=0.8)
            data = json.loads(req.read())
            
            # Fetch Car
            tree_path = data.get("car", {}).get("displayProperties", {}).get("fullTreePath", "")
            if tree_path: 
                summary["car"] = tree_path.split(",")[-1].strip()
                
            # Fetch Track Layout Name
            track_name = data.get("track", {}).get("displayProperties", {}).get("name", "")
            if track_name: 
                summary["track"] = track_name.strip()
                
        except: pass
        return summary

    def start(self):
        self.is_running = True
        self.thread = threading.Thread(target=self._run_loop, daemon=True)
        self.thread.start()
        self.ui_callback("Telemetry API Engine active. Waiting for LMU connection...")

    def stop(self):
        self.is_running = False
        if self.auth_subscription:
            try: self.auth_subscription.unsubscribe()
            except: pass
        if self.thread: self.thread.join(timeout=2)
        self.ui_callback("Telemetry API Engine stopped.")

    def _run_loop(self):
        standings_url = "http://localhost:6397/rest/watch/standings"
        session_url = "http://localhost:6397/rest/watch/sessionInfo" 
        
        last_saved_raw_time = -1.0 
        last_car_name = ""
        track_name = "Unknown Track"
        base_setup = {"bb": 50.0, "tc": 0, "abs": 0, "tc_cut": 0, "tc_slip": 0, "wing": 0}
        connected_to_game = False

        while self.is_running:
            try:
                # Basic session fallback for track name
                try:
                    sess_req = urllib.request.urlopen(session_url, timeout=0.5)
                    sess_data = json.loads(sess_req.read())
                    if isinstance(sess_data, dict) and 'trackName' in sess_data:
                        track_name = clean_string(sess_data['trackName'])
                except: pass

                new_setup = self.fetch_garage_setup()
                if new_setup and new_setup != base_setup:
                    base_setup = new_setup
                    self.ui_callback(f"[SYS] Setup Synced: BB {base_setup['bb']} | TC {base_setup['tc']} | ABS {base_setup['abs']}")

                req = urllib.request.urlopen(standings_url, timeout=1.0)
                standings = json.loads(req.read())
                
                if not connected_to_game:
                    self.ui_callback("[SYS] Connected to LMU REST API.")
                    connected_to_game = True

                vehicles = standings if isinstance(standings, list) else standings.get('vehicles', [])
                
                player_data = None
                for v in vehicles:
                    if v.get('player') == True or v.get('player') == 1:
                        player_data = v
                        break

                if player_data:
                    # Fetch detailed telemetry (Specific Track Layout & Car Model)
                    garage_summary = self.fetch_garage_summary()
                    
                    if garage_summary["track"]:
                        track_name = garage_summary["track"]  # Overwrite shortName with specific layout name
                    
                    # --- Extract & Normalize Car Class ---
                    raw_class = player_data.get('carClass', 'Unknown Class')
                    
                    # Normalisation mapping for Le Mans Ultimate API
                    class_map = {
                        "Hyper": "Hypercar",
                        "LMP_ELMS" : "LMP2",
                        "LMP3" : "LMP3",
                        "GT3": "LMGT3",
                        "GTE" : "GTE",
                    }
                    car_class = class_map.get(raw_class, raw_class)
                    
                    if garage_summary["car"]: 
                        current_car = garage_summary["car"]
                    else:
                        current_car = player_data.get('vehicleName', 'Unknown Car')
                        
                    driver_name = player_data.get('driverName', 'Unknown Driver')
                    current_last_lap = float(player_data.get('lastLapTime', -1.0))

                    if current_car != last_car_name and current_car != 'Unknown Car':
                        last_car_name = current_car
                        last_saved_raw_time = current_last_lap 
                        self.ui_callback(f"[TRACKING] Class: {car_class} | Vehicle: {current_car}")
                        self.ui_callback(f"[SYS] Driver Recognized: {driver_name}")

                    if current_last_lap > 0 and current_last_lap != last_saved_raw_time:
                        lap_time_str = format_laptime(current_last_lap)
                        
                        payload = {
                            "user_id": self.user_id, "track": track_name, "car": current_car, 
                            "car_class": car_class,
                            "lap_time": lap_time_str, "raw_time": current_last_lap, 
                            "abs": base_setup["abs"], "brake_bias": base_setup["bb"], 
                            "tc_onboard": base_setup["tc"], "tc_power_cut": base_setup["tc_cut"], 
                            "tc_slip_angle": base_setup["tc_slip"], "rear_wing": base_setup["wing"]
                        }
                        try:
                            self.supabase.table("laps").insert(payload).execute()
                            self.ui_callback(f"[SUCCESS] Lap Recorded: {lap_time_str} on {track_name}")
                        except Exception as e:
                            err_msg = str(e)
                            if "JWT expired" in err_msg or "PGRST303" in err_msg:
                                self.ui_callback("[SYS] Security token expired. Requesting a fresh one...")
                                try:
                                    self.supabase.auth.refresh_session()
                                    self.supabase.table("laps").insert(payload).execute()
                                    self.ui_callback(f"[SUCCESS] Lap Recorded (After Refresh): {lap_time_str} on {track_name}")
                                except Exception as refresh_err:
                                    self.ui_callback(f"[ERROR] Token refresh failed: {refresh_err}. Please restart connection.")
                            else:
                                self.ui_callback(f"[ERROR] DB Upload Failed: {e}")
                                
                            last_saved_raw_time = current_last_lap

                time.sleep(0.5)

            except Exception as e:
                if connected_to_game:
                    self.ui_callback("Awaiting Game Connection...")
                    connected_to_game = False
                time.sleep(2)

# --- LOCAL AUTH SERVER ---
class AuthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed_path = urllib.parse.urlparse(self.path)
        if parsed_path.path == '/callback':
            query_params = urllib.parse.parse_qs(parsed_path.query)
            access_token = query_params.get('access_token', [None])[0]
            refresh_token = query_params.get('refresh_token', [None])[0]

            if access_token and refresh_token:
                keyring.set_password(APP_NAME, "access_token", access_token)
                keyring.set_password(APP_NAME, "refresh_token", refresh_token)
                self.send_response(200)
                self.send_header('Content-type', 'text/html')
                self.end_headers()
                html = "<html><body style='background-color:#131315;color:#4AE176;font-family:monospace;text-align:center;padding-top:100px;'><h1 style='text-transform:uppercase;'>Handshake Complete</h1><p>You can close this tab and return to the app.</p><script>setTimeout(function(){window.close();},3000);</script></body></html>"
                self.wfile.write(html.encode('utf-8'))
                self.server.app_reference.on_auth_success(access_token, refresh_token)
            else:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b"Auth failed: Missing tokens.")
    def log_message(self, format, *args): pass


# --- UI APPLICATION ---
class PaddockCompanionApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.auth_server = None
        self.collector = None
        self.tray_icon = None
        self.settings = load_settings()
        self.update_download_url = None

        self.title(f"{APP_NAME} - v{APP_VERSION}")
        self.geometry("600x500")
        self.resizable(False, False)
        ctk.set_appearance_mode("dark")
        try:
            self.iconbitmap(get_resource_path("logo.ico"))
        except Exception as e:
            logging.error(f"Could not load window icon: {e}")

        self.protocol('WM_DELETE_WINDOW', self.on_closing)

        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # --- HEADER (With Update Button) ---
        self.header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.header_frame.pack(fill="x", pady=(0, 20))

        self.header_label = ctk.CTkLabel(self.header_frame, text="TELEMETRY COMPANION", font=("Space Grotesk", 18, "bold", "italic"), text_color="#E8173A")
        self.header_label.pack(side="left")

        self.update_btn = ctk.CTkButton(self.header_frame, text="UPDATE AVAILABLE", font=("Space Grotesk", 10, "bold"), fg_color="#4AE176", text_color="black", hover_color="#36b85a", height=24, command=self.open_update_link)
        
        # Start periodic update checker
        threading.Thread(target=self.update_checker_loop, daemon=True).start()

        self.auth_frame = ctk.CTkFrame(self.main_frame, fg_color="#1B1B1D", border_width=2, border_color="#353437", corner_radius=0)
        self.auth_frame.pack(fill="x", pady=(0, 20), ipadx=10, ipady=10)

        self.console_label = ctk.CTkLabel(self.main_frame, text="LIVE DATA FEED", font=("Space Grotesk", 10, "bold"), text_color="#a1a1aa")
        self.console_label.pack(anchor="w")

        self.console = ctk.CTkTextbox(self.main_frame, width=560, height=180, font=("Consolas", 12), text_color="#4AE176", fg_color="#0e0e10", corner_radius=0, border_width=1, border_color="#353437")
        self.console.pack(fill="both", expand=True, pady=(5, 0))
        self.console.configure(state="disabled")

        # --- FOOTER ---
        self.footer_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.footer_frame.pack(fill="x", pady=(15, 0))

        self.checkbox_frame = ctk.CTkFrame(self.footer_frame, fg_color="transparent")
        self.checkbox_frame.pack(side="left")

        self.autostart_var = ctk.BooleanVar(value=is_autostart_enabled())
        self.autostart_checkbox = ctk.CTkCheckBox(self.checkbox_frame, text="Start with Windows", font=("Space Grotesk", 11), fg_color="#E8173A", hover_color="#bf002a", variable=self.autostart_var, command=self.on_autostart_toggle)
        self.autostart_checkbox.pack(anchor="w", pady=(0, 5))

        self.minimize_var = ctk.BooleanVar(value=self.settings.get("minimize_to_tray", True))
        self.minimize_checkbox = ctk.CTkCheckBox(self.checkbox_frame, text="Minimize to Tray on Close", font=("Space Grotesk", 11), fg_color="#E8173A", hover_color="#bf002a", variable=self.minimize_var, command=self.on_minimize_toggle)
        self.minimize_checkbox.pack(anchor="w")

        self.support_btn = ctk.CTkButton(self.footer_frame, text="Support / Help", font=("Space Grotesk", 11, "underline"), text_color="#a1a1aa", fg_color="transparent", hover_color="#1B1B1D", width=100, command=lambda: webbrowser.open(SUPPORT_URL))
        self.support_btn.pack(side="right", anchor="s")

        self.log_to_console("System initialized.")
        self.check_authentication()
        
        # Check if we should show the welcome popup
        self.after(500, self.show_welcome_popup)

    def show_welcome_popup(self):
        """Displays a one-time welcome and instructions popup."""
        if not self.settings.get("has_seen_welcome", False):
            popup = ctk.CTkToplevel(self)
            popup.title("Welcome to LMU Paddock!")
            popup.geometry("450x350")
            popup.resizable(False, False)
            popup.attributes("-topmost", True)
            popup.configure(fg_color="#1B1B1D")
            
            popup.grab_set()

            # Header
            header_lbl = ctk.CTkLabel(popup, text="WELCOME TO THE GRID!", font=("Space Grotesk", 18, "bold", "italic"), text_color="#E8173A")
            header_lbl.pack(pady=(20, 10))

            # Main instructions
            intro_text = (
                "To get started, click the 'LOGIN WITH BROWSER' button on the main window "
                "to securely connect your LMU Paddock account."
            )
            intro_lbl = ctk.CTkLabel(popup, text=intro_text, font=("Space Grotesk", 12), justify="center", wraplength=380, text_color="white")
            intro_lbl.pack(padx=20, pady=(0, 15))

            # Important Warning Container
            warning_frame = ctk.CTkFrame(popup, fg_color="#201F21", border_width=1, border_color="#353437", corner_radius=5)
            warning_frame.pack(padx=20, pady=(0, 20), fill="x")

            warning_title = ctk.CTkLabel(warning_frame, text="⚠️ IMPORTANT: CAR SETUPS", font=("Space Grotesk", 11, "bold"), text_color="#FFD700")
            warning_title.pack(pady=(10, 5))

            warning_text = (
                "Due to current game limitations, your setup (Brake Bias, TC, ABS, etc.) "
                "is recorded EXACTLY as it is when you drive out of the garage. \n\n"
                "Live adjustments made on-track are not yet captured."
            )
            warning_lbl = ctk.CTkLabel(warning_frame, text=warning_text, font=("Space Grotesk", 11), justify="center", wraplength=340, text_color="#a1a1aa")
            warning_lbl.pack(padx=10, pady=(0, 10))

            def close_popup():
                self.settings["has_seen_welcome"] = True
                save_settings(self.settings)
                popup.grab_release()
                popup.destroy()

            popup.protocol("WM_DELETE_WINDOW", close_popup)
            
            btn = ctk.CTkButton(popup, text="GOT IT, LET'S RACE", font=("Space Grotesk", 12, "bold", "italic"), fg_color="#E8173A", hover_color="#bf002a", command=close_popup)
            btn.pack(pady=(0, 20))

    def update_checker_loop(self):
        """ Runs in background, checks for updates every hour """
        while True:
            has_update = self.check_for_updates()
            if has_update:
                break # Stop loop once we notify the user
            time.sleep(3600) # Check every 1 hour

    def check_for_updates(self):
        try:
            req = urllib.request.urlopen(UPDATE_URL, timeout=3.0)
            data = json.loads(req.read())
            latest_version = data.get("version", "1.0.3")
            self.update_download_url = data.get("url", SUPPORT_URL)

            if latest_version > APP_VERSION:
                self.after(0, self.show_update_button, latest_version)
                if self.tray_icon:
                    try:
                        self.tray_icon.notify(
                            f"Version {latest_version} is available. Click to download.",
                            title="LMU Paddock Update"
                        )
                    except: pass
                return True
        except json.JSONDecodeError:
            logging.warning("Update check skipped: version.json not found or invalid.")
        except Exception as e:
            logging.error(f"Update check failed: {e}")
        return False

    def show_update_button(self, new_version):
        self.update_btn.configure(text=f"UPDATE v{new_version} AVAILABLE")
        self.update_btn.pack(side="right")
        self.log_to_console(f"Notice: A new version of the Companion (v{new_version}) is available!")

    def open_update_link(self):
        if self.update_download_url:
            webbrowser.open(self.update_download_url)

    def on_autostart_toggle(self):
        enabled = self.autostart_var.get()
        toggle_autostart(enabled)
        self.log_to_console(f"Auto-start {'enabled' if enabled else 'disabled'}.")

    def on_minimize_toggle(self):
        self.settings["minimize_to_tray"] = self.minimize_var.get()
        save_settings(self.settings)
        self.log_to_console(f"Minimize to Tray {'enabled' if self.minimize_var.get() else 'disabled'}.")

    def on_closing(self):
        if self.minimize_var.get():
            self.hide_window()
        else:
            self.quit_app()

    def hide_window(self):
            self.withdraw() 
            
            if not self.tray_icon:
                menu = pystray.Menu(
                    pystray.MenuItem('Open Dashboard', self.show_window, default=True),
                    pystray.MenuItem('Exit Companion', self.quit_app)
                )
                try:
                    icon_path = get_resource_path("logo.ico")
                    if os.path.exists(icon_path):
                        icon_image = Image.open(icon_path)
                    else:
                        icon_image = Image.new('RGB', (64, 64), color=(232, 23, 58))
                    
                    self.tray_icon = pystray.Icon("LMU_Paddock", icon_image, f"{APP_NAME} (Running)", menu)
                    
                    self.tray_icon.run_detached() 
                except Exception as e:
                    logging.error(f"Could not setup tray icon: {e}")
                    self.deiconify() 

            self.log_to_console("Monitoring continued in background.")

    def show_window(self, icon=None, item=None):
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None
        self.after(0, self.deiconify)

    def quit_app(self, icon=None, item=None):
        if self.tray_icon:
            self.tray_icon.stop()
        if self.collector:
            self.collector.stop()
        if hasattr(self, 'lock_file'):
            self.lock_file.close()
            try: os.remove(LOCK_FILE)
            except: pass
        self.destroy()
        sys.exit(0)

    def log_to_console(self, message):
        self.after(0, self._safe_log_update, message)

    def _safe_log_update(self, message):
        self.console.configure(state="normal")
        self.console.insert("end", f"> {message}\n")
        self.console.see("end")
        self.console.configure(state="disabled")
        logging.info(message)

    def update_keyring_tokens(self, acc, ref):
        keyring.set_password(APP_NAME, "access_token", acc)
        keyring.set_password(APP_NAME, "refresh_token", ref)

    def check_authentication(self):
        access_token = keyring.get_password(APP_NAME, "access_token")
        refresh_token = keyring.get_password(APP_NAME, "refresh_token")
        if access_token and refresh_token:
            self.show_connected_ui()
            self.start_telemetry(access_token, refresh_token)
        else:
            self.show_login_ui()
            self.log_to_console("Awaiting authorization...")

    def show_login_ui(self):
        for widget in self.auth_frame.winfo_children(): widget.destroy()
        label = ctk.CTkLabel(self.auth_frame, text="Status: DISCONNECTED", font=("Space Grotesk", 12, "bold"), text_color="#a1a1aa")
        label.pack(side="left", padx=15, pady=10)
        login_btn = ctk.CTkButton(self.auth_frame, text="LOGIN WITH BROWSER", font=("Space Grotesk", 12, "bold"), fg_color="#E8173A", hover_color="#bf002a", corner_radius=0, command=self.start_auth_flow)
        login_btn.pack(side="right", padx=15, pady=10)

    def show_connected_ui(self):
        for widget in self.auth_frame.winfo_children(): widget.destroy()
        label = ctk.CTkLabel(self.auth_frame, text="Status: SECURE UPLINK ESTABLISHED", font=("Space Grotesk", 12, "bold"), text_color="#4AE176")
        label.pack(side="left", padx=15, pady=10)
        logout_btn = ctk.CTkButton(self.auth_frame, text="DISCONNECT", font=("Space Grotesk", 12, "bold"), fg_color="transparent", border_width=1, border_color="#5d3f3e", hover_color="#353437", corner_radius=0, command=self.logout)
        logout_btn.pack(side="right", padx=15, pady=10)

    def start_auth_flow(self):
        self.log_to_console("Initiating handshake protocol...")
        threading.Thread(target=self.run_local_server, daemon=True).start()
        webbrowser.open(AUTH_URL)

    def run_local_server(self):
        try:
            self.auth_server = HTTPServer(('127.0.0.1', PORT), AuthHandler)
            self.auth_server.app_reference = self 
            self.log_to_console(f"Listening for callback on port {PORT}...")
            self.auth_server.handle_request()
        except Exception as e:
            self.log_to_console(f"Server error: {e}")

    def on_auth_success(self, access_token, refresh_token):
        self.log_to_console("Tokens received and secured.")
        self.after(100, self.show_connected_ui)
        self.start_telemetry(access_token, refresh_token)

    def start_telemetry(self, access_token, refresh_token):
        if self.collector and self.collector.is_running: return
        self.log_to_console("Authenticating Telemetry Engine...")
        self.collector = TelemetryCollector(self.log_to_console, self.update_keyring_tokens, access_token, refresh_token)
        
        if not self.collector.user_id:
            self.log_to_console("Session expired or invalid. Please log in again.")
            self.logout()
            return
            
        self.collector.start()

    def logout(self):
        if self.collector: self.collector.stop()
        try:
            keyring.delete_password(APP_NAME, "access_token")
            keyring.delete_password(APP_NAME, "refresh_token")
        except: pass
        self.log_to_console("Credentials wiped. Telemetry stopped.")
        self.show_login_ui()

if __name__ == "__main__":
    lock = check_single_instance()
    if not lock:
        import tkinter.messagebox as mb
        mb.showwarning("App Already Running", "LMU Paddock Companion is already running.")
        sys.exit(0)
    
    app = PaddockCompanionApp()
    app.lock_file = lock 
    app.mainloop()