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
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw
import pystray
import xml.etree.ElementTree as ET
from supabase import create_client, Client

# --- CONFIGURATION ---
APP_VERSION = "1.1.0"
APP_NAME = "LMU Paddock Companion"
AUTH_URL = "https://lmupaddock.com/app-auth" 
SUPPORT_URL = "https://www.lmupaddock.com/support"  
PROFILE_URL = "https://lmupaddock.com/profile"
WEBSITE_BASE_URL = "https://lmupaddock.com"
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
        try: os.remove(LOCK_FILE)
        except OSError: return False
    try:
        f = open(LOCK_FILE, 'w')
        f.write(str(os.getpid()))
        return f
    except Exception: return False

def get_resource_path(relative_path):
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def clean_string(val):
    if isinstance(val, bytes): return val.decode('utf-8', errors='ignore').replace('\x00', '').strip()
    return str(val).replace('\x00', '').strip()

def format_laptime(seconds):
    if seconds <= 0.0: return "No time"
    seconds = round(seconds, 3)
    minutes = int(seconds // 60)
    remainder = seconds % 60
    return f"{minutes}:{remainder:06.3f}"

def load_settings():
    defaults = {
        "minimize_to_tray": True,
        "has_seen_welcome": False,
        "results_dir": "",
        "processed_xmls": [],
        "driver_name_override": "",   # leave blank to auto-detect from UserData/player
        "last_seen_release_notes": "",
        "last_seen_companion_news_id": "",
        "show_console": False,
        "show_settings": False,
    }
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                settings = json.load(f)
                for k, v in defaults.items():
                    if k not in settings: settings[k] = v
                return settings
        except: pass
    return defaults

def save_settings(settings):
    try:
        with open(SETTINGS_FILE, 'w') as f: json.dump(settings, f)
    except Exception as e: logging.error(f"Failed to save settings: {e}")

def is_autostart_enabled():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
        winreg.QueryValueEx(key, APP_NAME)
        winreg.CloseKey(key)
        return True
    except WindowsError: return False

def toggle_autostart(enable):
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
        if enable:
            base_cmd = f'"{sys.executable}"' if getattr(sys, 'frozen', False) else f'"{sys.executable}" "{os.path.abspath(__file__)}"'
            app_path = f'{base_cmd} --minimized'
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, app_path)
        else:
            try: winreg.DeleteValue(key, APP_NAME)
            except FileNotFoundError: pass
        winreg.CloseKey(key)
    except Exception as e: logging.error(f"Failed to toggle auto-start: {e}")

# --- TELEMETRY COLLECTOR ENGINE ---
class TelemetryCollector:
    def __init__(self, ui_callback, auth_update_callback, access_token, refresh_token, get_settings_cb, save_settings_cb):
        self.ui_callback = ui_callback
        self.auth_update_callback = auth_update_callback
        self.get_settings = get_settings_cb
        self.save_settings = save_settings_cb
        
        self.is_running = False
        self.live_thread = None
        self.xml_thread = None
        self.supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
        self.auth_subscription = None
        
        self.live_weather_cache = {"ambientTemp": 20.0, "trackTemp": 25.0, "raining": False}

        try:
            self.auth_subscription = self.supabase.auth.on_auth_state_change(self._handle_auth_change)
            self.supabase.auth.set_session(access_token, refresh_token)
            
            current_session = self.supabase.auth.get_session()
            if current_session and (current_session.access_token != access_token or current_session.refresh_token != refresh_token):
                self.auth_update_callback(current_session.access_token, current_session.refresh_token)
            
            user_res = self.supabase.auth.get_user()
            self.user_id = user_res.user.id
            self._report_app_version()
        except Exception as e:
            self.ui_callback(f"[ERROR] Auth verification failed: {e}")
            self.user_id = None

    def _report_app_version(self):
        """Update the user's profile row with the currently running companion version."""
        if not self.user_id:
            return
        try:
            self.supabase.table("profiles").update(
                {"app_version": APP_VERSION}
            ).eq("id", self.user_id).execute()
        except Exception as e:
            try:
                if "JWT expired" in str(e) or "401" in str(e):
                    self.supabase.auth.refresh_session()
                    self.supabase.table("profiles").update(
                        {"app_version": APP_VERSION}
                    ).eq("id", self.user_id).execute()
                else:
                    self.ui_callback(f"[WARN] Could not report app version: {e}")
            except Exception as e2:
                self.ui_callback(f"[WARN] Could not report app version: {e2}")

    def _handle_auth_change(self, event, session):
        if event == "TOKEN_REFRESHED" and session:
            self.auth_update_callback(session.access_token, session.refresh_token)
            self.ui_callback("[SYS] Security tokens rotated and secured.")

    # --- 1. LIVE REST API ENGINE ---
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
        except: return None

    def fetch_garage_summary(self):
        url = "http://localhost:6397/rest/garage/summary"
        summary = {"track": None}
        try:
            req = urllib.request.urlopen(url, timeout=0.8)
            data = json.loads(req.read())
            track_name = data.get("track", {}).get("displayProperties", {}).get("name", "")
            if track_name: summary["track"] = track_name.strip()
        except: pass
        return summary

    def _live_loop(self):
        standings_url = "http://localhost:6397/rest/watch/standings"
        session_url = "http://localhost:6397/rest/watch/sessionInfo" 
        
        last_saved_raw_time = -1.0 
        last_car_name = ""
        track_name = "Unknown Track"
        track_venue = None
        track_length_m = None
        base_setup = {"bb": 50.0, "tc": 0, "abs": 0, "tc_cut": 0, "tc_slip": 0, "wing": 0}
        connected_to_game = False

        while self.is_running:
            try:
                try:
                    sess_req = urllib.request.urlopen(session_url, timeout=0.5)
                    sess_data = json.loads(sess_req.read())
                    if isinstance(sess_data, dict) and 'trackName' in sess_data:
                        track_name = clean_string(sess_data['trackName'])
                        # Canonical venue + length so the DB trigger can resolve track_id
                        venue_raw = sess_data.get("trackVenueName") or sess_data.get("trackVenue")
                        if venue_raw:
                            track_venue = clean_string(venue_raw)
                        length_raw = sess_data.get("lapDistance") or sess_data.get("trackLength")
                        if length_raw is not None:
                            try: track_length_m = float(length_raw)
                            except: pass
                        self.live_weather_cache["ambientTemp"] = float(sess_data.get("ambientTemp", 20.0))
                        self.live_weather_cache["trackTemp"] = float(sess_data.get("trackTemp", 25.0))
                        self.live_weather_cache["raining"] = bool(sess_data.get("raining", False))
                except: pass

                new_setup = self.fetch_garage_setup()
                if new_setup and new_setup != base_setup:
                    base_setup = new_setup

                req = urllib.request.urlopen(standings_url, timeout=1.0)
                standings = json.loads(req.read())
                
                if not connected_to_game:
                    self.ui_callback("[SYS] Connected to LMU REST API (Live Layer).")
                    connected_to_game = True

                vehicles = standings if isinstance(standings, list) else standings.get('vehicles', [])
                
                player_data = None
                for v in vehicles:
                    if v.get('player') == True or v.get('player') == 1:
                        player_data = v
                        break

                if player_data:
                    garage_summary = self.fetch_garage_summary()
                    if garage_summary["track"]: track_name = garage_summary["track"]
                    
                    raw_class = player_data.get('carClass', 'Unknown Class')
                    class_map = {"Hyper": "Hypercar", "LMP_ELMS" : "LMP2", "LMP3" : "LMP3", "GT3": "LMGT3", "GTE" : "GTE"}
                    car_class = class_map.get(raw_class, raw_class)
                    current_car = clean_string(player_data.get('vehicleName', 'Unknown Car'))
                    current_last_lap = float(player_data.get('lastLapTime', -1.0))

                    if current_car != last_car_name and current_car != 'Unknown Car':
                        last_car_name = current_car
                        last_saved_raw_time = current_last_lap 
                        self.ui_callback(f"[LIVE] Class: {car_class} | Vehicle: {current_car}")

                    if current_last_lap > 0 and current_last_lap != last_saved_raw_time:
                        lap_time_str = format_laptime(current_last_lap)
                        last_saved_raw_time = current_last_lap
                        
                        payload = {
                            "user_id": self.user_id, "track": track_name, "car": current_car, 
                            "car_class": car_class, "lap_time": lap_time_str, "raw_time": current_last_lap, 
                            "abs": base_setup["abs"], "brake_bias": base_setup["bb"], 
                            "tc_onboard": base_setup["tc"], "tc_power_cut": base_setup["tc_cut"], 
                            "tc_slip_angle": base_setup["tc_slip"], "rear_wing": base_setup["wing"],
                            "track_venue": track_venue,
                            "track_length_m": track_length_m,
                        }
                        try:
                            self.supabase.table("laps").insert(payload).execute()
                            self.ui_callback(f"[LIVE] Fast Lap Synced: {lap_time_str}")
                        except Exception as e:
                            if "JWT expired" in str(e) or "401" in str(e):
                                try:
                                    self.supabase.auth.refresh_session()
                                    self.supabase.table("laps").insert(payload).execute()
                                except: self.is_running = False 
                time.sleep(0.5)

            except Exception as e:
                if connected_to_game:
                    self.ui_callback("Awaiting Live Game Connection...")
                    connected_to_game = False
                time.sleep(2)

    # --- 2. ANALYTICAL XML ENGINE ---
    def _xml_watcher_loop(self):
        self.ui_callback("[SYS] Analytical XML Engine Started. Watching for new sessions...")
        while self.is_running:
            try:
                settings = self.get_settings()
                results_dir = settings.get("results_dir", "")
                processed = settings.get("processed_xmls", [])

                if results_dir and os.path.exists(results_dir):
                    for filename in os.listdir(results_dir):
                        if filename.endswith(".xml") and filename not in processed:
                            filepath = os.path.join(results_dir, filename)
                            
                            time.sleep(2) 
                            status, msg = self._parse_and_upload_xml(filepath)
                            
                            if status == "SUCCESS" or status == "SKIP":
                                processed.append(filename)
                                settings["processed_xmls"] = processed
                                self.save_settings(settings)
                                if status == "SUCCESS":
                                    self.ui_callback(f"[XML] Session Auto-Uploaded: {filename}")
                            elif status == "ERROR":
                                self.ui_callback(f"[CRITICAL ERROR] Auto-upload paused: {msg}")
                                self.ui_callback(f"Need help? Visit: {SUPPORT_URL}")
                                time.sleep(60) 
                                
            except Exception as e:
                logging.error(f"XML Watcher Error: {e}")
            
            time.sleep(5)

    def sync_historical_data(self, on_complete_callback=None):
        def run_sync():
            try:
                settings = self.get_settings()
                results_dir = settings.get("results_dir", "")
                processed = settings.get("processed_xmls", [])
                
                if not results_dir or not os.path.exists(results_dir):
                    self.ui_callback("[XML] No valid Results directory selected.")
                    return

                all_xmls = [f for f in os.listdir(results_dir) if f.endswith(".xml")]
                unprocessed = [f for f in all_xmls if f not in processed]
                
                if not unprocessed:
                    self.ui_callback("[XML] All historical data is already synced or was previously skipped.")
                    return
                    
                self.ui_callback(f"[XML] Found {len(unprocessed)} files to check. Syncing securely...")
                
                success_count = 0
                error_count = 0
                for idx, filename in enumerate(unprocessed):
                    if not self.is_running: break 
                    
                    filepath = os.path.join(results_dir, filename)
                    time.sleep(0.5) 
                    
                    status, err_msg = self._parse_and_upload_xml(filepath)
                    
                    if status == "SUCCESS":
                        processed.append(filename)
                        success_count += 1
                        self.ui_callback(f"[SUCCESS] Uploaded: {filename}")
                    
                    elif status == "SKIP":
                        processed.append(filename)
                        self.ui_callback(f"[SKIP] {filename} -> {err_msg}")
                        
                    elif status == "ERROR":
                        # Log but keep going so a single transient error doesn't kill the run
                        error_count += 1
                        self.ui_callback(f"[ERROR] {filename} -> {err_msg}")
                        # Brief backoff before continuing
                        time.sleep(1.0)
                
                settings["processed_xmls"] = processed
                self.save_settings(settings)
                
                if success_count > 0 or error_count > 0:
                    summary = f"[XML] Sync complete. Uploaded {success_count} new session(s)"
                    if error_count > 0:
                        summary += f", {error_count} error(s) (will retry next sync)"
                    self.ui_callback(summary + ".")
                else:
                    self.ui_callback("[XML] Done checking files. No new valid sessions were found to upload.")
            except Exception as e:
                logging.error(f"Historical Sync Error: {e}")
                self.ui_callback(f"[XML ERROR] Sync process crashed: {e}")
            finally:
                if on_complete_callback:
                    on_complete_callback()

        threading.Thread(target=run_sync, daemon=True).start()

    def _resolve_player_name(self):
        """
        Find the local LMU driver's display name (e.g. "Jeffrey Hermann"),
        which is what appears in <Driver><Name> in result XMLs.

        Order of precedence:
          1) Manual override in companion settings ("driver_name_override")
          2) UserData/player/<PlayerFile>.JSON -> DRIVER."Player Name"
        """
        try:
            settings = self.get_settings()
            override = (settings.get("driver_name_override") or "").strip()
            if override:
                return override

            results_dir = settings.get("results_dir") or ""
            if not results_dir:
                return None

            # results_dir is .../UserData/Log/Results -> walk up two levels for UserData
            userdata_dir = os.path.abspath(os.path.join(results_dir, os.pardir, os.pardir))
            player_dir = os.path.join(userdata_dir, "player")
            if not os.path.isdir(player_dir):
                return None

            # Prefer Settings.JSON (matches <PlayerFile>Settings</PlayerFile> in XML),
            # otherwise pick the first *.JSON that has a DRIVER section.
            candidates = ["Settings.JSON"] + [
                f for f in os.listdir(player_dir)
                if f.lower().endswith(".json") and f.lower() != "settings.json"
            ]
            for fname in candidates:
                fpath = os.path.join(player_dir, fname)
                if not os.path.isfile(fpath): continue
                try:
                    with open(fpath, "r", encoding="utf-8-sig") as f:
                        data = json.load(f)
                    drv = data.get("DRIVER") or {}
                    name = (drv.get("Player Name") or drv.get("Player Nick") or "").strip()
                    if name:
                        return name
                except Exception:
                    continue
        except Exception as e:
            logging.error(f"Player name resolution failed: {e}")
        return None

    def _parse_and_upload_xml(self, filepath):
        """
        Parse a Le Mans Ultimate result XML and upload it to Supabase.

        Returns (status, message) where status is 'SUCCESS' | 'SKIP' | 'ERROR'.
        """
        SESSION_ELEMS = ("Race", "Qualify", "Warmup", "TestDay")
        try:
            tree = ET.parse(filepath)
            root = tree.getroot()

            results = root.find("RaceResults")
            if results is None:
                return "SKIP", "Not a valid RaceResults XML."

            # ---- pick the session block (Race / Qualify / Practice* / Warmup / TestDay) ----
            session_node = None
            session_type = None
            for child in results:
                tag = child.tag
                if tag in SESSION_ELEMS or tag.startswith("Practice"):
                    session_node = child
                    session_type = "Practice" if tag.startswith("Practice") else tag
                    break

            if session_node is None:
                return "SKIP", "No race/qualify/practice block in XML."

            # ---- session metadata ----
            setting     = (results.findtext("Setting") or "").strip()
            server_name = (results.findtext("ServerName") or "").strip()
            track_venue = (results.findtext("TrackVenue") or "").strip()
            track_course= (results.findtext("TrackCourse") or "").strip() or "Unknown Track"
            track_event = (results.findtext("TrackEvent") or "").strip()
            game_version= (results.findtext("GameVersion") or "").strip()

            try:    track_length_m = float(results.findtext("TrackLength") or 0) or None
            except: track_length_m = None
            try:    race_laps = int(results.findtext("RaceLaps") or 0)
            except: race_laps = None
            try:    race_minutes = int(results.findtext("RaceTime") or 0)
            except: race_minutes = None

            try:
                most_laps = int(session_node.findtext("MostLapsCompleted") or 0) or None
            except:
                most_laps = None
            try:
                fas = int(session_node.findtext("FormationAndStart") or 0)
            except:
                fas = None

            # session_started_at: <DateTime> in <Race>/<Qualify> is unix epoch seconds
            session_started_at = None
            dt_text = session_node.findtext("DateTime") or results.findtext("DateTime")
            if dt_text:
                try:
                    from datetime import datetime, timezone as _tz
                    session_started_at = datetime.fromtimestamp(int(dt_text), tz=_tz.utc).isoformat()
                except Exception:
                    session_started_at = None

            # ---- locate the local player's <Driver> node ----
            player_name = self._resolve_player_name()
            if not player_name:
                return "SKIP", "Could not resolve your LMU driver name (set it in companion settings)."

            target = player_name.casefold()
            player_node = None
            for drv in session_node.findall("Driver"):
                if (drv.findtext("Name") or "").strip().casefold() == target:
                    player_node = drv
                    break

            if player_node is None:
                return "SKIP", f"You ({player_name}) were not driving in this session."

            # ---- player result block ----
            car_name   = (player_node.findtext("VehName") or "Unknown Car").strip()
            car_class  = (player_node.findtext("CarClass") or "Unknown Class").strip()
            car_number = (player_node.findtext("CarNumber") or "").strip()
            team_name  = (player_node.findtext("TeamName") or "").strip() or None
            try:    grid_pos = int(player_node.findtext("GridPos") or 0) or None
            except: grid_pos = None
            try:    finish_pos = int(player_node.findtext("Position") or 0) or None
            except: finish_pos = None
            try:    class_grid_pos = int(player_node.findtext("ClassGridPos") or 0) or None
            except: class_grid_pos = None
            try:    class_finish_pos = int(player_node.findtext("ClassPosition") or 0) or None
            except: class_finish_pos = None
            try:    best_lap_time = float(player_node.findtext("BestLapTime") or -1)
            except: best_lap_time = None
            if best_lap_time is not None and best_lap_time <= 0: best_lap_time = None
            try:    finish_time = float(player_node.findtext("FinishTime") or -1)
            except: finish_time = None
            if finish_time is not None and finish_time <= 0: finish_time = None
            try:    total_laps = int(player_node.findtext("Laps") or 0)
            except: total_laps = 0
            try:    pit_stops = int(player_node.findtext("Pitstops") or 0)
            except: pit_stops = 0
            finish_status   = (player_node.findtext("FinishStatus") or "").strip() or None
            control_and_aids = (player_node.findtext("ControlAndAids") or "").strip() or None

            # ---- per-lap parsing ----
            class_map = {"Hyper": "Hypercar", "LMP_ELMS": "LMP2", "LMP3": "LMP3",
                         "GT3": "LMGT3", "GTE": "GTE", "LMP2_ELMS": "LMP2"}
            car_class_norm = class_map.get(car_class, car_class)

            laps_rows = []
            valid_times = []
            top_speed_kmh = 0.0
            total_fuel_used = 0.0
            used_wet = False

            def _ff(node, attr, default=None):
                v = node.get(attr)
                if v is None or v == "": return default
                try: return float(v)
                except: return default

            for lap in player_node.findall("Lap"):
                try:    lap_num = int(lap.get("num", "0"))
                except: continue
                try:    position = int(lap.get("p", "0")) or None
                except: position = None

                lap_time_text = (lap.text or "").strip()
                if lap_time_text == "--.----" or not lap_time_text:
                    lap_time = None
                    is_valid = False
                else:
                    try:
                        lap_time = float(lap_time_text)
                        is_valid = lap_time > 0
                    except:
                        lap_time = None
                        is_valid = False

                s1 = _ff(lap, "s1"); s2 = _ff(lap, "s2"); s3 = _ff(lap, "s3")
                ts = _ff(lap, "topspeed", 0.0) or 0.0
                if ts > top_speed_kmh: top_speed_kmh = ts
                fu = _ff(lap, "fuelUsed", 0.0) or 0.0
                if fu > 0: total_fuel_used += fu

                fcomp_raw = lap.get("fcompound", "") or ""
                rcomp_raw = lap.get("rcompound", "") or ""
                fcomp = fcomp_raw.split(",", 1)[1] if "," in fcomp_raw else (fcomp_raw or None)
                rcomp = rcomp_raw.split(",", 1)[1] if "," in rcomp_raw else (rcomp_raw or None)
                if fcomp and "wet" in fcomp.lower(): used_wet = True
                if rcomp and "wet" in rcomp.lower(): used_wet = True

                is_pit = lap.get("pit") == "1"

                if is_valid and lap_time and lap_time > 0:
                    valid_times.append(lap_time)

                laps_rows.append({
                    "lap_num": lap_num,
                    "position": position,
                    "lap_time": lap_time,
                    "is_valid": is_valid,
                    "s1": s1, "s2": s2, "s3": s3,
                    "top_speed": ts or None,
                    "fuel": _ff(lap, "fuel"),
                    "fuel_used": fu or None,
                    "tire_wear_fl": _ff(lap, "twfl"),
                    "tire_wear_fr": _ff(lap, "twfr"),
                    "tire_wear_rl": _ff(lap, "twrl"),
                    "tire_wear_rr": _ff(lap, "twrr"),
                    "compound_front": fcomp,
                    "compound_rear":  rcomp,
                    "is_pit": is_pit,
                })

            if not laps_rows:
                return "SKIP", "No lap data."

            # ---- aggregates ----
            clean_count = len(valid_times)
            invalid_count = len(laps_rows) - clean_count

            avg_time = sum(valid_times) / clean_count if clean_count else None
            median_time = None
            stdev = None
            if clean_count:
                sorted_t = sorted(valid_times)
                mid = clean_count // 2
                median_time = sorted_t[mid] if clean_count % 2 == 1 else (sorted_t[mid - 1] + sorted_t[mid]) / 2.0
                if clean_count > 1:
                    mean = avg_time
                    stdev = (sum((t - mean) ** 2 for t in valid_times) / (clean_count - 1)) ** 0.5

            # ---- incident counts (player-only) from <Stream> ----
            driver_contacts = 0
            object_contacts = 0
            tl_warnings = 0
            stream = session_node.find("Stream")
            if stream is not None:
                # Match by player name appearing in the text. Robust to "Name(id)" suffix.
                lower_target = target
                for inc in stream.findall("Incident"):
                    txt = (inc.text or "").lower()
                    if lower_target not in txt: continue
                    # "reported contact (...) with another vehicle" => driver contact
                    # "reported contact (...) with Immovable/Sign/Wall/Object" => object contact
                    if "another vehicle" in txt:
                        driver_contacts += 1
                    else:
                        object_contacts += 1
                for tl in stream.findall("TrackLimits"):
                    if (tl.get("Driver") or "").casefold() == target:
                        if (tl.text or "").strip().lower() == "warning":
                            tl_warnings += 1

            # rough rain inference: live cache OR any wet compound used
            is_raining = bool(self.live_weather_cache.get("raining")) or used_wet

            # ---- file hash for stronger dedup ----
            try:
                import hashlib
                with open(filepath, "rb") as f:
                    source_hash = hashlib.sha256(f.read()).hexdigest()
            except Exception:
                source_hash = None

            # ---- pre-check: if this file is already uploaded, skip the POST entirely ----
            if source_hash and self.user_id:
                try:
                    existing = (
                        self.supabase.table("race_sessions")
                        .select("id")
                        .eq("user_id", self.user_id)
                        .eq("source_hash", source_hash)
                        .limit(1)
                        .execute()
                    )
                    if existing and getattr(existing, "data", None):
                        return "SKIP", "Already uploaded (matching source hash)."
                except Exception:
                    # If the dedup check fails (network/RLS), fall through and let the
                    # insert path decide. We'll still catch a conflict below.
                    pass

            session_payload = {
                "user_id": self.user_id,
                "source_file": os.path.basename(filepath),
                "source_hash": source_hash,
                "session_type": session_type,
                "setting": setting or None,
                "server_name": server_name or None,
                "track_venue": track_venue or None,
                "track_course": track_course,
                "track_event": track_event or None,
                "track_length_m": track_length_m,
                "game_version": game_version or None,
                "race_laps": race_laps,
                "race_minutes": race_minutes,
                "formation_and_start": fas,
                "most_laps_completed": most_laps,
                "player_name": player_name,
                "car_name": car_name,
                "car_class": car_class_norm,
                "car_number": car_number or None,
                "team_name": team_name,
                "grid_pos": grid_pos,
                "finish_pos": finish_pos,
                "class_grid_pos": class_grid_pos,
                "class_finish_pos": class_finish_pos,
                "best_lap_time": best_lap_time,
                "finish_time": finish_time,
                "total_laps": total_laps,
                "pit_stops": pit_stops,
                "finish_status": finish_status,
                "control_and_aids": control_and_aids,
                "clean_lap_count": clean_count,
                "invalid_lap_count": invalid_count,
                "avg_lap_time": avg_time,
                "median_lap_time": median_time,
                "consistency_stdev": stdev,
                "top_speed_kmh": top_speed_kmh or None,
                "total_fuel_used": total_fuel_used or None,
                "driver_contacts": driver_contacts,
                "object_contacts": object_contacts,
                "track_limit_warnings": tl_warnings,
                "is_raining": is_raining,
                "session_started_at": session_started_at,
            }

            # ---- upload ----
            try:
                resp = None
                last_err = None
                for attempt in range(3):
                    try:
                        resp = self.supabase.table("race_sessions").insert(session_payload).execute()
                        break
                    except Exception as up_err:
                        last_err = up_err
                        msg = str(up_err).lower()
                        # WinError 10035 / WSAEWOULDBLOCK and similar transient socket issues
                        if "10035" in msg or "wouldblock" in msg or "timeout" in msg or "temporarily" in msg:
                            time.sleep(1.5 * (attempt + 1))
                            continue
                        raise
                if resp is None:
                    raise last_err or RuntimeError("Insert failed after retries")
                if not resp.data:
                    # PostgREST can return an empty body on conflict (e.g. when an
                    # ON CONFLICT DO NOTHING trigger fires, or when RLS hides the
                    # returned row). Treat as already-uploaded so we stop retrying.
                    return "SKIP", "Already uploaded (no row returned)."
                session_id = resp.data[0]["id"]

                lap_payload = []
                for r in laps_rows:
                    r2 = dict(r)
                    r2["session_id"] = session_id
                    r2["user_id"] = self.user_id
                    lap_payload.append(r2)

                # chunk in case sessions have 200+ laps (e.g. 24h races)
                CHUNK = 200
                for i in range(0, len(lap_payload), CHUNK):
                    batch = lap_payload[i:i+CHUNK]
                    for attempt in range(3):
                        try:
                            self.supabase.table("race_session_laps").insert(batch).execute()
                            break
                        except Exception as lap_err:
                            msg = str(lap_err).lower()
                            if "10035" in msg or "wouldblock" in msg or "timeout" in msg or "temporarily" in msg:
                                if attempt < 2:
                                    time.sleep(1.5 * (attempt + 1))
                                    continue
                            raise

                return "SUCCESS", ""
            except Exception as db_err:
                err_str = str(db_err)
                err_low = err_str.lower()
                # Treat any duplicate / unique-violation / 409 conflict as SKIP
                # so the watcher stops re-trying the same already-uploaded file.
                code = getattr(db_err, "code", "") or ""
                status_code = getattr(db_err, "status_code", None)
                if (
                    "23505" in err_str
                    or "duplicate key" in err_low
                    or "already exists" in err_low
                    or "unique constraint" in err_low
                    or "conflict" in err_low
                    or "\"409\"" in err_str
                    or status_code == 409
                    or code == "23505"
                ):
                    return "SKIP", "Already uploaded (duplicate file)."
                if "PGRST205" in err_str or ("relation" in err_low and "does not exist" in err_low):
                    return "ERROR", "Database tables 'race_sessions'/'race_session_laps' do not exist. Run the migration."
                return "ERROR", err_str

        except Exception as e:
            logging.error(f"Failed to parse XML {filepath}: {e}")
            return "SKIP", "Corrupt or unreadable XML."

    def start(self):
        self.is_running = True
        self.live_thread = threading.Thread(target=self._live_loop, daemon=True)
        self.xml_thread = threading.Thread(target=self._xml_watcher_loop, daemon=True)
        self.live_thread.start()
        self.xml_thread.start()

    def stop(self):
        self.is_running = False
        if self.auth_subscription:
            try: self.auth_subscription.unsubscribe()
            except: pass
        if self.live_thread: self.live_thread.join(timeout=2)
        if self.xml_thread: self.xml_thread.join(timeout=2)
        self.ui_callback("Telemetry API Engine stopped.")

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
        self.geometry("600x310")
        self.resizable(False, False)
        ctk.set_appearance_mode("dark")
        try: self.iconbitmap(get_resource_path("logo.ico"))
        except Exception: pass

        self.protocol('WM_DELETE_WINDOW', self.on_closing)

        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # --- HEADER ---
        self.header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.header_frame.pack(fill="x", pady=(0, 20))

        self.header_label = ctk.CTkLabel(self.header_frame, text="LMU PADDOCK COMPANION", font=("Space Grotesk", 18, "bold", "italic"), text_color="#E8173A")
        self.header_label.pack(side="left")

        # Settings cog button (right side of the header).
        self.settings_btn = ctk.CTkButton(
            self.header_frame, text="\u2699  SETTINGS", font=("Space Grotesk", 11, "bold"),
            fg_color="transparent", border_width=1, border_color="#353437",
            text_color="#a1a1aa", hover_color="#1B1B1D", width=110, height=28,
            command=self.toggle_settings_window,
        )
        self.settings_btn.pack(side="right")

        self.update_btn = ctk.CTkButton(self.header_frame, text="UPDATE AVAILABLE", font=("Space Grotesk", 10, "bold"), fg_color="#4AE176", text_color="black", hover_color="#36b85a", height=24, command=self.open_update_link)
        threading.Thread(target=self.update_checker_loop, daemon=True).start()

        # --- AUTH FRAME (always visible) ---
        self.auth_frame = ctk.CTkFrame(self.main_frame, fg_color="#1B1B1D", border_width=2, border_color="#353437", corner_radius=0)
        self.auth_frame.pack(fill="x", pady=(0, 20), ipadx=10, ipady=10)

        # --- FOOTER ---
        self.footer_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.footer_frame.pack(fill="x", side="bottom", pady=(15, 0))

        self.checkbox_frame = ctk.CTkFrame(self.footer_frame, fg_color="transparent")
        self.checkbox_frame.pack(side="left")

        self.autostart_var = ctk.BooleanVar(value=is_autostart_enabled())
        self.autostart_checkbox = ctk.CTkCheckBox(self.checkbox_frame, text="Start with Windows", font=("Space Grotesk", 11), fg_color="#E8173A", hover_color="#bf002a", variable=self.autostart_var, command=self.on_autostart_toggle)
        self.autostart_checkbox.pack(anchor="w", pady=(0, 5))

        self.minimize_var = ctk.BooleanVar(value=self.settings.get("minimize_to_tray", True))
        self.minimize_checkbox = ctk.CTkCheckBox(self.checkbox_frame, text="Minimize to Tray", font=("Space Grotesk", 11), fg_color="#E8173A", hover_color="#bf002a", variable=self.minimize_var, command=self.on_minimize_toggle)
        self.minimize_checkbox.pack(anchor="w")

        # Real solid-button footer actions.
        self.support_btn = ctk.CTkButton(
            self.footer_frame, text="SUPPORT / HELP",
            font=("Space Grotesk", 10, "bold"),
            fg_color="#353437", hover_color="#505050",
            text_color="white", width=120, height=30,
            command=lambda: webbrowser.open(SUPPORT_URL),
        )
        self.support_btn.pack(side="right", padx=(6, 0))

        self.profile_btn = ctk.CTkButton(
            self.footer_frame, text="MY PROFILE",
            font=("Space Grotesk", 10, "bold"),
            fg_color="#E8173A", hover_color="#bf002a",
            text_color="white", width=110, height=30,
            command=lambda: webbrowser.open(PROFILE_URL),
        )
        self.profile_btn.pack(side="right")

        self.live_feed_btn = ctk.CTkButton(
            self.footer_frame, text="VIEW LOG",
            font=("Space Grotesk", 10, "bold"),
            fg_color="transparent", border_width=1, border_color="#353437",
            text_color="#a1a1aa", hover_color="#1B1B1D",
            width=100, height=30,
            command=self.toggle_console_window,
        )
        self.live_feed_btn.pack(side="right", padx=(0, 6))

        # --- BACKING STATE FOR POPUPS ---
        # These StringVars are owned by the main app so their values persist
        # even when the settings popup is closed and re-opened.
        self.path_var = ctk.StringVar(value=self.settings.get("results_dir", "Select UserData\\Log\\Results folder..."))
        self.driver_name_var = ctk.StringVar(value=self.settings.get("driver_name_override", ""))

        # Live-feed log buffer (drained into the popup when first shown).
        from collections import deque as _deque
        self._log_buffer = _deque(maxlen=500)
        self.console = None              # textbox lives inside the popup
        self.settings_window = None
        self.console_window = None

        # --- INITIAL VISIBILITY ---
        # First-run UX: if no Results folder is configured, open the settings
        # popup automatically so the user actually sees the picker.
        has_results_dir = bool(self.settings.get("results_dir"))
        if not has_results_dir:
            self.after(200, self._open_settings_window)

        self.log_to_console("System initialized.")
        self.check_authentication()
        self.after(500, self.show_welcome_popup)
        self.after(800, self.show_release_notes_popup)

    # --- SETTINGS POPUP -----------------------------------------------------
    def toggle_settings_window(self):
        if self.settings_window and self.settings_window.winfo_exists():
            try: self.settings_window.destroy()
            except Exception: pass
            self.settings_window = None
        else:
            self._open_settings_window()

    def _open_settings_window(self):
        if self.settings_window and self.settings_window.winfo_exists():
            try:
                self.settings_window.deiconify()
                self.settings_window.lift()
                self.settings_window.focus_force()
            except Exception: pass
            return

        win = ctk.CTkToplevel(self)
        win.title("Settings - LMU Paddock Companion")
        win.geometry("660x520")
        win.resizable(False, False)
        win.configure(fg_color="#131315")
        try:
            win.after(250, lambda: win.iconbitmap(get_resource_path("logo.ico")))
        except Exception:
            pass

        def on_close():
            try: win.destroy()
            except Exception: pass
            self.settings_window = None
        win.protocol("WM_DELETE_WINDOW", on_close)

        body = ctk.CTkFrame(win, fg_color="#1B1B1D", border_width=1, border_color="#353437", corner_radius=0)
        body.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(
            body, text="LMU RESULTS FOLDER (FOR ENDURANCE DATA)",
            font=("Space Grotesk", 10, "bold"), text_color="#a1a1aa",
        ).pack(anchor="w", padx=10, pady=(10, 0))

        path_row = ctk.CTkFrame(body, fg_color="transparent")
        path_row.pack(fill="x", padx=10, pady=5)
        self.path_entry = ctk.CTkEntry(path_row, textvariable=self.path_var, state="disabled", fg_color="#0e0e10", border_color="#353437")
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(path_row, text="BROWSE", width=80, fg_color="#353437", hover_color="#505050", command=self.select_directory).pack(side="right")

        # Folder-finder mini-guide (font sized to match BROWSE button).
        guide_frame = ctk.CTkFrame(body, fg_color="#0e0e10", border_width=1, border_color="#353437", corner_radius=0)
        guide_frame.pack(fill="x", padx=10, pady=(2, 8))
        ctk.CTkLabel(
            guide_frame,
            text="HOW TO FIND YOUR RESULTS FOLDER",
            font=("Space Grotesk", 11, "bold"),
            text_color="#a1a1aa",
        ).pack(anchor="w", padx=8, pady=(8, 4))
        ctk.CTkLabel(
            guide_frame,
            text=(
                "Steam (default):\n"
                "    C:\\Program Files (x86)\\Steam\\steamapps\\common\\Le Mans Ultimate\\UserData\\Log\\Results\n\n"
                "Tip: in Steam, right-click Le Mans Ultimate \u2192 Manage \u2192 Browse local files,\n"
                "then open  UserData \\ Log \\ Results."
            ),
            font=("Consolas", 11),
            text_color="#a1a1aa",
            justify="left",
            anchor="w",
        ).pack(anchor="w", padx=8, pady=(0, 8))

        # --- DRIVER NAME OVERRIDE ---
        detected = self._detect_player_name_from_files()
        label_text = "LMU DRIVER NAME (auto-detected; override only if wrong)"
        if detected:
            label_text = f"LMU DRIVER NAME \u2014 auto-detected: {detected}"
        ctk.CTkLabel(body, text=label_text,
                     font=("Space Grotesk", 10, "bold"), text_color="#a1a1aa").pack(anchor="w", padx=10, pady=(8, 0))
        driver_row = ctk.CTkFrame(body, fg_color="transparent")
        driver_row.pack(fill="x", padx=10, pady=5)
        placeholder = f"Detected: {detected} (leave blank to use this)" if detected else "e.g. Jeffrey Hermann (leave blank to auto-detect)"
        self.driver_entry = ctk.CTkEntry(driver_row, textvariable=self.driver_name_var,
                                         placeholder_text=placeholder,
                                         fg_color="#0e0e10", border_color="#353437")
        self.driver_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(driver_row, text="SAVE", width=80, fg_color="#353437", hover_color="#505050",
                      command=self.save_driver_name).pack(side="right")
        if detected:
            ctk.CTkLabel(body,
                         text=f"If '{detected}' is not your in-game driver name, type the correct one above and SAVE.",
                         font=("Space Grotesk", 11), text_color="#6b7280").pack(anchor="w", padx=10)

        btn_row = ctk.CTkFrame(body, fg_color="transparent")
        btn_row.pack(fill="x", padx=10, pady=(12, 12))
        self.sync_btn = ctk.CTkButton(btn_row, text="SYNC HISTORICAL DATA", font=("Space Grotesk", 10, "bold"), fg_color="transparent", border_width=1, border_color="#4AE176", text_color="#4AE176", hover_color="#1a3b23", command=self.trigger_historic_sync)
        self.sync_btn.pack(side="left")
        self.reset_btn = ctk.CTkButton(btn_row, text="RESET CACHE", font=("Space Grotesk", 10, "bold"), fg_color="transparent", border_width=1, border_color="#E8173A", text_color="#E8173A", hover_color="#3b1016", width=100, command=self.reset_cache)
        self.reset_btn.pack(side="right")

        ctk.CTkButton(
            win, text="CLOSE",
            font=("Space Grotesk", 10, "bold"),
            fg_color="#353437", hover_color="#505050",
            text_color="white", height=30,
            command=on_close,
        ).pack(fill="x", padx=16, pady=(0, 16))

        self.settings_window = win
        try:
            win.lift()
            win.focus_force()
        except Exception: pass

    # --- LIVE FEED POPUP ----------------------------------------------------
    def toggle_console_window(self):
        if self.console_window and self.console_window.winfo_exists():
            try: self.console_window.destroy()
            except Exception: pass
            self.console_window = None
            self.console = None
            self.settings["show_console"] = False
            save_settings(self.settings)
        else:
            self._open_console_window()

    def _open_console_window(self):
        if self.console_window and self.console_window.winfo_exists():
            try:
                self.console_window.deiconify()
                self.console_window.lift()
                self.console_window.focus_force()
            except Exception: pass
            return

        win = ctk.CTkToplevel(self)
        win.title("Log - LMU Paddock Companion")
        win.geometry("640x360")
        win.minsize(420, 220)
        win.resizable(True, True)
        win.configure(fg_color="#131315")
        try:
            win.after(250, lambda: win.iconbitmap(get_resource_path("logo.ico")))
        except Exception:
            pass

        def on_close():
            try: win.destroy()
            except Exception: pass
            self.console_window = None
            self.console = None
            self.settings["show_console"] = False
            save_settings(self.settings)
        win.protocol("WM_DELETE_WINDOW", on_close)

        header = ctk.CTkFrame(win, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(12, 4))
        ctk.CTkLabel(
            header, text="LOG",
            font=("Space Grotesk", 11, "bold"), text_color="#a1a1aa",
        ).pack(side="left")
        ctk.CTkButton(
            header, text="CLEAR",
            font=("Space Grotesk", 9, "bold"),
            fg_color="transparent", border_width=1, border_color="#353437",
            text_color="#a1a1aa", hover_color="#1B1B1D",
            width=64, height=22,
            command=self._clear_console,
        ).pack(side="right")

        self.console = ctk.CTkTextbox(
            win, font=("Consolas", 12),
            text_color="#4AE176", fg_color="#0e0e10",
            corner_radius=0, border_width=1, border_color="#353437",
        )
        self.console.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        # Replay buffered log lines.
        self.console.configure(state="normal")
        for line in list(self._log_buffer):
            self.console.insert("end", line)
        self.console.see("end")
        self.console.configure(state="disabled")

        self.console_window = win
        self.settings["show_console"] = True
        save_settings(self.settings)
        try:
            win.lift()
            win.focus_force()
        except Exception: pass

    def _clear_console(self):
        self._log_buffer.clear()
        if self.console is not None:
            try:
                self.console.configure(state="normal")
                self.console.delete("1.0", "end")
                self.console.configure(state="disabled")
            except Exception:
                pass

    def select_directory(self):
        selected_dir = filedialog.askdirectory(title="Select LMU Results Folder")
        if selected_dir:
            self.settings["results_dir"] = selected_dir
            self.path_var.set(selected_dir)
            save_settings(self.settings)
            self.log_to_console(f"Results folder updated: {selected_dir}")

    def save_driver_name(self):
        new_name = (self.driver_name_var.get() or "").strip()
        self.settings["driver_name_override"] = new_name
        save_settings(self.settings)
        if new_name:
            self.log_to_console(f"[SYS] Driver name override set to '{new_name}'.")
        else:
            self.log_to_console("[SYS] Driver name override cleared (auto-detect from UserData/player).")

    def _detect_player_name_from_files(self):
        """File-based detection only (ignores override) so the UI can show what
        WOULD be auto-detected, letting the user verify before relying on it."""
        try:
            results_dir = self.settings.get("results_dir") or ""
            if not results_dir:
                return None
            userdata_dir = os.path.abspath(os.path.join(results_dir, os.pardir, os.pardir))
            player_dir = os.path.join(userdata_dir, "player")
            if not os.path.isdir(player_dir):
                return None
            candidates = ["Settings.JSON"] + [
                f for f in os.listdir(player_dir)
                if f.lower().endswith(".json") and f.lower() != "settings.json"
            ]
            for fname in candidates:
                fpath = os.path.join(player_dir, fname)
                if not os.path.isfile(fpath):
                    continue
                try:
                    with open(fpath, "r", encoding="utf-8-sig") as f:
                        data = json.load(f)
                    drv = data.get("DRIVER") or {}
                    name = (drv.get("Player Name") or drv.get("Player Nick") or "").strip()
                    if name:
                        return name
                except Exception:
                    continue
        except Exception:
            pass
        return None

    def reset_cache(self):
        confirmed = messagebox.askyesno(
            "Reset Sync Cache?",
            (
                "This will erase the companion's record of which Result files have "
                "already been uploaded.\n\n"
                "The next sync will re-check EVERY result file in your LMU Results "
                "folder and re-upload any session that isn't already in the cloud. "
                "Depending on your LMU online race history this can take a while.\n\n"
                "Continue?"
            ),
            icon="warning",
            parent=self,
        )
        if not confirmed:
            return
        self.settings["processed_xmls"] = []
        save_settings(self.settings)
        self.log_to_console("[SYS] Sync memory erased! The next sync will check EVERY file again.")

    def trigger_historic_sync(self):
        if not self.settings.get("results_dir"):
            self.log_to_console("[ERROR] Please select the Results folder first!")
            return
        
        if self.collector:
            self.sync_btn.configure(state="disabled", text="SYNCING...")
            self.collector.sync_historical_data(on_complete_callback=self.on_historic_sync_complete)
        else:
            self.log_to_console("[ERROR] You must be logged in to sync data.")

    def on_historic_sync_complete(self):
        self.after(0, lambda: self.sync_btn.configure(state="normal", text="SYNC HISTORICAL DATA"))

    def show_welcome_popup(self):
        if not self.settings.get("has_seen_welcome", False):
            popup = ctk.CTkToplevel(self)
            popup.title("Welcome to LMU Paddock!")
            popup.geometry("450x350")
            popup.resizable(False, False)
            popup.attributes("-topmost", True)
            popup.configure(fg_color="#1B1B1D")
            popup.grab_set()
            try:
                popup.after(250, lambda: popup.iconbitmap(get_resource_path("logo.ico")))
            except Exception:
                pass

            header_lbl = ctk.CTkLabel(popup, text="WELCOME TO THE GRID!", font=("Space Grotesk", 18, "bold", "italic"), text_color="#E8173A")
            header_lbl.pack(pady=(20, 10))

            intro_lbl = ctk.CTkLabel(popup, text="To get started, log in and select your LMU 'Results' folder to enable Endurance Analytics.", font=("Space Grotesk", 12), justify="center", wraplength=380, text_color="white")
            intro_lbl.pack(padx=20, pady=(0, 15))

            def close_popup():
                self.settings["has_seen_welcome"] = True
                save_settings(self.settings)
                popup.grab_release()
                popup.destroy()

            popup.protocol("WM_DELETE_WINDOW", close_popup)
            ctk.CTkButton(popup, text="GOT IT, LET'S RACE", font=("Space Grotesk", 12, "bold", "italic"), fg_color="#E8173A", hover_color="#bf002a", command=close_popup).pack(pady=(0, 20))

    def show_release_notes_popup(self):
        # Don't show on the very first run -- the welcome popup covers that.
        if not self.settings.get("has_seen_welcome", False):
            return
        # Fetch the latest companion-flagged news post off the UI thread.
        threading.Thread(target=self._load_companion_news, daemon=True).start()

    def _load_companion_news(self):
        """
        Pull the most recent news_posts row flagged with notify_companion = true.
        Skipped if it's the same post the user already dismissed.
        """
        try:
            client = create_client(SUPABASE_URL, SUPABASE_KEY)
            resp = (
                client.table("news_posts")
                .select("id, title, summary, category, created_at")
                .eq("notify_companion", True)
                .order("created_at", desc=True)
                .limit(1)
                .execute()
            )
            rows = getattr(resp, "data", None) or []
            if not rows:
                return
            post = rows[0]
        except Exception as e:
            logging.info(f"Skipping news popup (fetch failed): {e}")
            return

        post_id = str(post.get("id") or "")
        if not post_id:
            return
        if self.settings.get("last_seen_companion_news_id", "") == post_id:
            return

        # Hop back onto the Tk thread to actually render the popup.
        try:
            self.after(0, lambda: self._render_news_popup(post))
        except Exception:
            pass

    def _render_news_popup(self, post):
        post_id = str(post.get("id") or "")
        title = (post.get("title") or "Update").strip() or "Update"
        summary = (post.get("summary") or "").strip()
        category = (post.get("category") or "Update").strip() or "Update"
        created_at = (post.get("created_at") or "").strip()
        try:
            # 2026-04-26T12:00:00+00:00 -> 26 Apr 2026
            from datetime import datetime as _dt
            iso = created_at.replace("Z", "+00:00")
            date_str = _dt.fromisoformat(iso).strftime("%d %b %Y").upper()
        except Exception:
            date_str = ""

        article_url = f"{WEBSITE_BASE_URL}/news/{post_id}"

        popup = ctk.CTkToplevel(self)
        popup.title(f"What's New \u2014 {title}")
        popup.geometry("520x460")
        popup.resizable(False, False)
        popup.attributes("-topmost", True)
        popup.configure(fg_color="#1B1B1D")
        popup.grab_set()
        try:
            popup.after(250, lambda: popup.iconbitmap(get_resource_path("logo.ico")))
        except Exception:
            pass

        header_lbl = ctk.CTkLabel(
            popup,
            text="WHAT'S NEW",
            font=("Space Grotesk", 18, "bold", "italic"),
            text_color="#E8173A",
        )
        header_lbl.pack(pady=(20, 4))

        meta_text = f"{category}" + (f"  \u2022  {date_str}" if date_str else "")
        meta_lbl = ctk.CTkLabel(
            popup,
            text=meta_text,
            font=("Space Grotesk", 11, "italic"),
            text_color="#a1a1aa",
        )
        meta_lbl.pack(pady=(0, 12))

        body_frame = ctk.CTkFrame(popup, fg_color="#131315", border_width=1, border_color="#353437", corner_radius=0)
        body_frame.pack(fill="both", expand=True, padx=20, pady=(0, 15))

        title_lbl = ctk.CTkLabel(
            body_frame,
            text=title,
            font=("Space Grotesk", 14, "bold"),
            text_color="#4AE176",
            anchor="w",
            justify="left",
            wraplength=440,
        )
        title_lbl.pack(fill="x", padx=15, pady=(15, 6))

        summary_box = ctk.CTkTextbox(
            body_frame,
            font=("Space Grotesk", 11),
            text_color="white",
            fg_color="#131315",
            border_width=0,
            wrap="word",
        )
        summary_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        summary_box.insert("1.0", summary or "Open the article on the website for full details.")
        summary_box.configure(state="disabled")

        def mark_seen():
            self.settings["last_seen_companion_news_id"] = post_id
            save_settings(self.settings)

        def close_popup():
            mark_seen()
            popup.grab_release()
            popup.destroy()

        def open_article():
            mark_seen()
            try: webbrowser.open(article_url)
            except Exception: pass
            popup.grab_release()
            popup.destroy()

        popup.protocol("WM_DELETE_WINDOW", close_popup)

        btn_row = ctk.CTkFrame(popup, fg_color="transparent")
        btn_row.pack(pady=(0, 20))
        ctk.CTkButton(
            btn_row,
            text="READ FULL ARTICLE",
            font=("Space Grotesk", 12, "bold", "italic"),
            fg_color="#E8173A",
            hover_color="#bf002a",
            width=180,
            command=open_article,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btn_row,
            text="DISMISS",
            font=("Space Grotesk", 12, "bold", "italic"),
            fg_color="transparent",
            border_width=1,
            border_color="#353437",
            hover_color="#1B1B1D",
            text_color="#a1a1aa",
            width=120,
            command=close_popup,
        ).pack(side="left")

    def update_checker_loop(self):
        while True:
            if self.check_for_updates(): break
            time.sleep(3600)

    def check_for_updates(self):
        try:
            req = urllib.request.urlopen(UPDATE_URL, timeout=3.0)
            data = json.loads(req.read())
            latest_version = data.get("version", "1.1.0")
            self.update_download_url = data.get("url", SUPPORT_URL)

            if latest_version > APP_VERSION:
                self.after(0, self.show_update_button, latest_version)
                return True
        except: pass
        return False

    def show_update_button(self, new_version):
        self.update_btn.configure(text=f"UPDATE v{new_version} AVAILABLE")
        self.update_btn.pack(side="right")
        self.log_to_console(f"Notice: A new version of the Companion (v{new_version}) is available!")

    def open_update_link(self):
        if self.update_download_url: webbrowser.open(self.update_download_url)

    def on_autostart_toggle(self):
        toggle_autostart(self.autostart_var.get())

    def on_minimize_toggle(self):
        self.settings["minimize_to_tray"] = self.minimize_var.get()
        save_settings(self.settings)

    def on_closing(self):
        if self.minimize_var.get(): self.hide_window()
        else: self.quit_app()

    def hide_window(self):
        self.withdraw() 
        if not self.tray_icon:
            menu = pystray.Menu(
                pystray.MenuItem('Open Dashboard', self.show_window, default=True),
                pystray.MenuItem('Exit Companion', self.quit_app)
            )
            try:
                icon_path = get_resource_path("logo.ico")
                icon_image = Image.open(icon_path) if os.path.exists(icon_path) else Image.new('RGB', (64, 64), color=(232, 23, 58))
                self.tray_icon = pystray.Icon("LMU_Paddock", icon_image, f"{APP_NAME} (Running)", menu)
                self.tray_icon.run_detached() 
            except Exception as e:
                logging.error(f"Could not setup tray icon: {e}")
                self.deiconify() 
        self.log_to_console("Monitoring continued in background.")

    def show_window(self, icon=None, item=None):
        if self.tray_icon:
            try: self.tray_icon.stop()
            except Exception: pass
            self.tray_icon = None
        try:
            self.after(0, self.deiconify)
        except Exception:
            try: self.deiconify()
            except Exception: pass

    def quit_app(self, icon=None, item=None):
        # pystray invokes menu callbacks from its own worker thread. Touching Tk
        # widgets from a non-Tk thread crashes the app with "invalid command name"
        # errors as <Configure> events fire on freshly-destroyed canvases. Marshal
        # the actual shutdown onto the Tk main loop.
        try:
            self.after(0, self._do_quit)
        except Exception:
            # Tk may already be torn down; fall back to direct shutdown.
            self._do_quit()

    def _do_quit(self):
        self._shutting_down = True
        # Stop tray icon first so it stops dispatching menu callbacks.
        if self.tray_icon:
            try: self.tray_icon.stop()
            except Exception: pass
            self.tray_icon = None
        # Stop the telemetry collector (signals threads, joins with timeout).
        if self.collector:
            try: self.collector.stop()
            except Exception: pass
        # Release the single-instance lock.
        if hasattr(self, 'lock_file'):
            try: self.lock_file.close()
            except Exception: pass
            try: os.remove(LOCK_FILE)
            except Exception: pass
        # Tear down Tk cleanly.
        try: self.quit()
        except Exception: pass
        try: self.destroy()
        except Exception: pass
        # Daemon background threads (httpx, watcher) may be mid-request; force exit
        # so we don't hang waiting for them.
        os._exit(0)

    def log_to_console(self, message):
        if getattr(self, "_shutting_down", False):
            try: logging.info(message)
            except Exception: pass
            return
        try:
            self.after(0, self._safe_log_update, message)
        except Exception:
            try: logging.info(message)
            except Exception: pass

    def _safe_log_update(self, message):
        timestamp = time.strftime("%H:%M:%S")
        line = f"[{timestamp}] > {message}\n"
        try:
            self._log_buffer.append(line)
        except Exception:
            pass
        if self.console is not None:
            try:
                self.console.configure(state="normal")
                self.console.insert("end", line)
                self.console.see("end")
                self.console.configure(state="disabled")
            except Exception:
                pass
        logging.info(message)

    def get_settings_for_collector(self):
        return self.settings

    def save_settings_from_collector(self, new_settings):
        self.settings = new_settings
        save_settings(self.settings)

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
        ctk.CTkLabel(self.auth_frame, text="Status: DISCONNECTED", font=("Space Grotesk", 12, "bold"), text_color="#a1a1aa").pack(side="left", padx=15, pady=10)
        ctk.CTkButton(self.auth_frame, text="LOGIN WITH BROWSER", font=("Space Grotesk", 12, "bold"), fg_color="#E8173A", hover_color="#bf002a", corner_radius=0, command=self.start_auth_flow).pack(side="right", padx=15, pady=10)

    def show_connected_ui(self):
        for widget in self.auth_frame.winfo_children(): widget.destroy()
        ctk.CTkLabel(self.auth_frame, text="Status: SECURE UPLINK ESTABLISHED", font=("Space Grotesk", 12, "bold"), text_color="#4AE176").pack(side="left", padx=15, pady=10)
        ctk.CTkButton(self.auth_frame, text="DISCONNECT", font=("Space Grotesk", 12, "bold"), fg_color="transparent", border_width=1, border_color="#5d3f3e", hover_color="#353437", corner_radius=0, command=self.logout).pack(side="right", padx=15, pady=10)

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
        except Exception as e: self.log_to_console(f"Server error: {e}")

    def on_auth_success(self, access_token, refresh_token):
        self.log_to_console("Tokens received and secured.")
        self.after(100, self.show_connected_ui)
        self.start_telemetry(access_token, refresh_token)

    def start_telemetry(self, access_token, refresh_token):
        if self.collector and self.collector.is_running: return
        self.log_to_console("Authenticating Telemetry Engine...")
        self.collector = TelemetryCollector(self.log_to_console, self.update_keyring_tokens, access_token, refresh_token, self.get_settings_for_collector, self.save_settings_from_collector)
        
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
    # Refresh autostart entry so existing installs pick up the --minimized flag.
    if is_autostart_enabled():
        toggle_autostart(True)
    if "--minimized" in sys.argv and app.settings.get("minimize_to_tray", True):
        app.after(0, app.hide_window)
    app.mainloop()