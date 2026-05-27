import os
import threading
import time
from datetime import datetime, timedelta
import configparser
import requests
import tkinter as tk
from tkinter import filedialog

import customtkinter as ctk

CONFIG_FILE = "config.ini"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("Bot Export Scada&Jms")
        self.geometry("1100x700")
        self.minsize(1000, 650)

        self.scheduler_running = False
        self.job_running = False
        self.stop_requested = False

        self.create_ui()
        self.load_config()

    def get_text(self, widget):
        return widget.get().strip()

    def set_text(self, widget, value):
        widget.delete(0, tk.END)
        widget.insert(0, value)

    def save_config(self):
        cfg = configparser.ConfigParser()
        cfg["API"] = {
            "dws_url": self.get_text(self.dws_url),
            "dws_token": self.get_text(self.dws_token),
            "jms_token": self.get_text(self.jms_token),
        }
        cfg["PATH"] = {"path": self.get_text(self.path)}
        cfg["FILE"] = {
            "name_dws": self.get_text(self.name_dws) or "DWS9-11.xlsx",
            "name_auto": self.get_text(self.name_auto) or "DWSXAUTOPDA.xlsx",
            "name_dwspda": self.get_text(self.name_dwspda) or "DWSPDA.xlsx",
            "name_realtime_db": self.get_text(self.name_realtime_db) or "RealtimeDB.xlsx",
        }
        cfg["TIME"] = {"run_minute": self.delay.get() or "5"}
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            cfg.write(f)

    def load_config(self):
        cfg = configparser.ConfigParser()
        if not os.path.exists(CONFIG_FILE):
            cfg["API"] = {"dws_url": "", "dws_token": "", "jms_token": ""}
            cfg["PATH"] = {"path": ""}
            cfg["FILE"] = {
                "name_dws": "DWS9-11.xlsx",
                "name_auto": "DWSXAUTOPDA.xlsx",
                "name_dwspda": "DWSPDA.xlsx",
                "name_realtime_db": "RealtimeDB.xlsx",
            }
            cfg["TIME"] = {"run_minute": "5"}
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                cfg.write(f)

        cfg.read(CONFIG_FILE, encoding="utf-8")
        a = cfg["API"] if cfg.has_section("API") else {}
        self.set_text(self.dws_url, a.get("dws_url", ""))
        self.set_text(self.dws_token, a.get("dws_token", ""))
        self.set_text(self.jms_token, a.get("jms_token", ""))

        p = cfg["PATH"] if cfg.has_section("PATH") else {}
        self.set_text(self.path, p.get("path", ""))

        f = cfg["FILE"] if cfg.has_section("FILE") else {}
        self.set_text(self.name_dws, f.get("name_dws", "DWS9-11.xlsx"))
        self.set_text(self.name_auto, f.get("name_auto", "DWSXAUTOPDA.xlsx"))
        self.set_text(self.name_dwspda, f.get("name_dwspda", "DWSPDA.xlsx"))
        self.set_text(self.name_realtime_db, f.get("name_realtime_db", "RealtimeDB.xlsx"))

        t = cfg["TIME"] if cfg.has_section("TIME") else {}
        minute = t.get("run_minute", "5")
        if minute in self.minutes:
            self.delay.set(minute)

    def create_ui(self):
        container = ctk.CTkFrame(self, corner_radius=0, fg_color="#0f172a")
        container.pack(fill="both", expand=True)

        header = ctk.CTkFrame(container, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(16, 10))
        ctk.CTkLabel(header, text="Export Dashboard", font=ctk.CTkFont(size=28, weight="bold")).pack(side="left")
        self.status = ctk.CTkLabel(header, text="Status: Idle", fg_color="#1e293b", corner_radius=14, padx=14, pady=6)
        self.status.pack(side="right")

        self.tabview = ctk.CTkTabview(container, corner_radius=16)
        self.tabview.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.tab_home = self.tabview.add("Home")
        self.tab_setting = self.tabview.add("Setting")

        self.build_home()
        self.build_setting()

    def build_home(self):
        frame = ctk.CTkFrame(self.tab_home, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        self.minutes = [str(i) for i in range(60)]
        hours = [f"{i:02d}:00" for i in range(24)]

        panel = ctk.CTkFrame(frame, corner_radius=12)
        panel.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(panel, text="Run minute").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.delay = ctk.CTkComboBox(panel, values=self.minutes, width=90)
        self.delay.set("5")
        self.delay.grid(row=0, column=1, padx=6, pady=10, sticky="w")

        ctk.CTkLabel(panel, text="Start date (YYYY-MM-DD)").grid(row=1, column=0, padx=10, pady=8, sticky="w")
        self.start_date = ctk.CTkEntry(panel, width=150)
        self.start_date.grid(row=1, column=1, padx=6, pady=8)
        self.set_text(self.start_date, datetime.now().strftime("%Y-%m-%d"))

        self.start_hour = ctk.CTkComboBox(panel, values=hours, width=100)
        self.start_hour.set("13:00")
        self.start_hour.grid(row=1, column=2, padx=6, pady=8)

        ctk.CTkLabel(panel, text="End date (YYYY-MM-DD)").grid(row=1, column=3, padx=10, pady=8, sticky="w")
        self.end_date = ctk.CTkEntry(panel, width=150)
        self.end_date.grid(row=1, column=4, padx=6, pady=8)
        self.set_text(self.end_date, (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d"))

        self.end_hour = ctk.CTkComboBox(panel, values=hours, width=100)
        self.end_hour.set("23:00")
        self.end_hour.grid(row=1, column=5, padx=6, pady=8)

        self.btn_start = ctk.CTkButton(panel, text="Start", width=100, command=self.toggle_start)
        self.btn_start.grid(row=1, column=6, padx=8)
        self.btn_run = ctk.CTkButton(panel, text="Run Now", width=100, fg_color="#0ea5e9", hover_color="#0284c7", command=self.toggle_run)
        self.btn_run.grid(row=1, column=7, padx=8)

        self.log_text = ctk.CTkTextbox(frame, corner_radius=12)
        self.log_text.pack(fill="both", expand=True)

    def build_setting(self):
        root = ctk.CTkFrame(self.tab_setting, fg_color="transparent")
        root.pack(fill="both", expand=True, padx=12, pady=12)

        self.dws_url = self._form_input(root, "DWS API URL", 0)
        self.dws_token = self._form_input(root, "DWS Token", 1)
        self.jms_token = self._form_input(root, "JMS AuthToken", 2)

        ctk.CTkLabel(root, text="Save Path").grid(row=3, column=0, sticky="w", padx=10, pady=8)
        self.path = ctk.CTkEntry(root)
        self.path.grid(row=3, column=1, sticky="ew", padx=10, pady=8)
        self.btn_browse = ctk.CTkButton(root, text="Browse", width=100, command=self.browse)
        self.btn_browse.grid(row=3, column=2, padx=10, pady=8)

        self.name_dws = self._form_input(root, "DWS File", 4)
        self.name_auto = self._form_input(root, "AUTOPDA File", 5)
        self.name_dwspda = self._form_input(root, "DWSPDA File", 6)
        self.name_realtime_db = self._form_input(root, "Realtime DB File", 7)
        root.grid_columnconfigure(1, weight=1)

    def _form_input(self, parent, label, row):
        ctk.CTkLabel(parent, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=8)
        entry = ctk.CTkEntry(parent)
        entry.grid(row=row, column=1, columnspan=2, sticky="ew", padx=10, pady=8)
        return entry

    def toggle_start(self):
        if not self.scheduler_running:
            self.start(); self.btn_start.configure(text="Stop")
        else:
            self.stop(); self.btn_start.configure(text="Start")

    def toggle_run(self):
        if self.job_running:
            self.stop_requested = True
            self.status.configure(text="Status: Stopping...")
            self.btn_run.configure(text="Run Now")
            return
        self.stop_requested = False
        threading.Thread(target=self.execute_pipeline, daemon=True).start()
        self.btn_run.configure(text="Stop")

    def log(self, msg, tag="SYS"):
        now = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"{now} [{tag}] {msg}\n")
        self.log_text.see("end")

    def browse(self):
        p = filedialog.askdirectory(title="Select Folder")
        if p:
            self.set_text(self.path, p)

    def run_dws(self):
        self.log("Start DWS request...", "DWS9-11")
        start, end = self.get_time_range()
        try:
            url = self.get_text(self.dws_url)
            token = self.get_text(self.dws_token)
            data = {"startTime": start, "endTime": end, "barcodeList": [], "businessTimeType": "", "dbCode": "2", "curPage": 1, "pageSize": 20}
            headers = {"Content-Type": "application/json;charset=UTF-8", "token": token, "Origin": "http://10.30.32.10", "Referer": "http://10.30.32.10/", "User-Agent": "Mozilla/5.0"}
            res = requests.post(url, json=data, headers=headers, timeout=(5, 10))
            self.log(f"STATUS: {res.status_code}", "DWS9-11")
            self.log(f"SIZE: {len(res.content)}", "DWS9-11")
            if res.status_code == 200 and len(res.content) > 100:
                os.makedirs(self.get_text(self.path), exist_ok=True)
                path = os.path.join(self.get_text(self.path), self.get_text(self.name_dws))
                with open(path, "wb") as f:
                    f.write(res.content)
                self.log(f"Downloaded: {self.get_text(self.name_dws)}")
            else:
                self.log(f"Download failed: {res.text[:200]}")
                self.status.configure(text="Status: Error")
        except Exception as e:
            self.log(f"ERROR: {e}")
            self.status.configure(text="Status: Error")
        finally:
            self.log("DWS finished", "SYS")

    def run_jms_auto(self):
        self._export_jms("https://jmsgw.jtexpress.co.th/operatingplatform", {"Content-Type": "application/json;charset=UTF-8", "authtoken": self.get_text(self.jms_token)}, *self.get_time_range(), "建包扫描", self.get_text(self.name_auto))

    def run_jms_pda(self):
        self._export_jms("https://jmsgw.jtexpress.co.th/operatingplatform", {"Content-Type": "application/json;charset=UTF-8", "authtoken": self.get_text(self.jms_token)}, *self.get_time_range(), "卸车扫描", self.get_text(self.name_dwspda))

    def run_realtime_db(self):
        self.log("Generate Realtime DB", "REALTIME")
        headers = {"Content-Type": "application/json;charset=UTF-8", "authtoken": self.get_text(self.jms_token), "origin": "https://jms.jtexpress.co.th", "referer": "https://jms.jtexpress.co.th/", "routename": "TrackRealTimeMonitoringDB", "lang": "TH", "langtype": "TH"}
        base = "https://jmsgw.jtexpress.co.th"
        res = requests.post(f"{base}/businessindicator/bigdataReport/pageExcelByTask/trail_monitor_detail_doris", json={"orderSourceCode": ["JMS"], "scanCode": "999004", "scanFranCode": "555090", "scanAgentCode": "555090", "countryId": "1", "modelName": "ควบคุมติดตามแบบเรียลไทม์DB(รายละเอียด)"}, headers=headers, timeout=(5, 10))
        if "login" in res.text.lower():
            raise Exception("Realtime DB token expired")
        self.log("Create export success", "REALTIME")
        if not self.sleep_with_stop(60):
            return
        list_payload = {"current": 1, "size": 20, "startTime": datetime.now().strftime("%Y-%m-%d 00:00:00"), "endTime": datetime.now().strftime("%Y-%m-%d 23:59:59"), "countryId": "1", "networkCode": "999004", "userId": 13144940}
        download_url = None
        for _ in range(30):
            if self.stop_requested:
                return
            res = requests.post(f"{base}/businessindicator/bigdataReport/report/file/list", json=list_payload, headers=headers, timeout=(5, 10))
            if "login" in res.text.lower():
                raise Exception("Realtime DB token expired")
            records = res.json().get("data", {}).get("list", [])
            records.sort(key=lambda x: x.get("createTime", ""), reverse=True)
            for record in records:
                if record.get("status") == 2 and record.get("downUrl") and record.get("business", "") == "trail_monitor_detail_doris":
                    download_url = record.get("downUrl")
                    self.log("Found download url", "REALTIME")
                    break
            if download_url:
                break
            self.log("Waiting file generate...", "REALTIME")
            if not self.sleep_with_stop(10):
                return
        if not download_url:
            raise Exception("Realtime DB download url not found")
        file = requests.get(download_url, timeout=(5, 30))
        if len(file.content) < 1000:
            raise Exception("Realtime DB file invalid")
        os.makedirs(self.get_text(self.path), exist_ok=True)
        save_path = os.path.join(self.get_text(self.path), self.get_text(self.name_realtime_db))
        with open(save_path, "wb") as f:
            f.write(file.content)
        self.log(f"Downloaded {self.get_text(self.name_realtime_db)}", "REALTIME")

    def _export_jms(self, base, headers, start, end, scanType, filename):
        self.log(f"Generate {filename}", "JMS")
        run_time = datetime.now()
        res = requests.post(f"{base}/scanningContrast/asyncDownExcel", json={"startTimeStr": start, "endTimeStr": end, "scanNetworkCode": "999004", "scanType": scanType, "excelType": "downExcelAll"}, headers=headers, timeout=(5, 10))
        if "login" in res.text.lower():
            raise Exception("Token expired")
        if not self.sleep_with_stop(50):
            return
        for _ in range(30):
            if self.stop_requested:
                self.log("Stopped by user"); return
            if not self.sleep_with_stop(10):
                return
            res = requests.post(f"{base}/downLoadCenter/downLoadInfoList", json={"current": 1, "size": 20}, headers=headers, timeout=(5, 10))
            if "login" in res.text.lower():
                raise Exception("Token expired")
            records = res.json().get("data", {}).get("records", [])
            records.sort(key=lambda x: x.get("downTime", ""), reverse=True)
            for r in records:
                try:
                    down_time = datetime.strptime(r.get("downTime"), "%Y-%m-%d %H:%M:%S")
                except Exception:
                    continue
                if scanType in r.get("queryJson", "") and r.get("finishOrNot") == "1" and r.get("downUrl") and down_time >= run_time - timedelta(hours=1):
                    sign = requests.post(f"{base}/downLoadCenter/getDownloadSignedUrl", json=r, headers=headers, timeout=(5, 10)).json()
                    url = sign.get("data")
                    if not url:
                        continue
                    file = requests.get(url if isinstance(url, str) else url[0], timeout=(5, 10))
                    if "html" in file.headers.get("Content-Type", "").lower() or "login" in file.text.lower():
                        raise Exception("Token expired")
                    if len(file.content) < 100_000:
                        self.log("File not ready, retry...", "JMS")
                        continue
                    os.makedirs(self.get_text(self.path), exist_ok=True)
                    path = os.path.join(self.get_text(self.path), filename)
                    with open(path, "wb") as f:
                        f.write(file.content)
                    self.log(f"Downloaded {filename}", "JMS")
                    return
        raise Exception(f"{filename} fail")

    def get_time_range(self):
        start = datetime.strptime(f"{self.get_text(self.start_date)} {self.start_hour.get()}", "%Y-%m-%d %H:%M")
        end = datetime.strptime(f"{self.get_text(self.end_date)} {self.end_hour.get()}", "%Y-%m-%d %H:%M")
        if start > end:
            start, end = end, start
        return (start.strftime("%Y-%m-%d %H:%M:%S"), end.strftime("%Y-%m-%d %H:%M:%S"))

    def sleep_with_stop(self, seconds):
        for _ in range(seconds):
            if self.stop_requested:
                return False
            time.sleep(1)
        return True

    def main_loop(self):
        self.log("Main loop started", "SYS")
        run_minute = int(self.delay.get())
        if not hasattr(self, "next_run") or self.next_run is None:
            now = datetime.now()
            self.next_run = now.replace(minute=run_minute, second=0, microsecond=0)
            if now > self.next_run + timedelta(seconds=30):
                self.next_run += timedelta(hours=1)
        while self.scheduler_running:
            self.log(f"Next run at {self.next_run.strftime('%H:%M:%S')}", "SYS")
            while datetime.now() < self.next_run:
                if not self.scheduler_running:
                    return
                time.sleep(1)
            if not self.job_running:
                threading.Thread(target=self.execute_pipeline, daemon=True).start()
            self.next_run += timedelta(hours=1)

    def execute_pipeline(self):
        if self.job_running:
            self.log("Skip: job running", "SYS")
            return
        self.job_running = True
        self.save_config()
        self.set_ui(False)
        self.status.configure(text="Status: Running")
        try:
            if self.stop_requested: return
            self.run_dws()
            if self.stop_requested or not self.sleep_with_stop(5): raise Exception("Stopped")
            self.run_jms_auto()
            if self.stop_requested or not self.sleep_with_stop(5): raise Exception("Stopped")
            self.run_jms_pda()
            if self.stop_requested or not self.sleep_with_stop(5): raise Exception("Stopped")
            self.run_realtime_db()
            self.status.configure(text="Status: Success")
        except Exception as e:
            self.log(f"ERROR: {e}", "SYS")
            self.status.configure(text="Status: Error")
        finally:
            self.job_running = False
            self.stop_requested = False
            self.after(0, lambda: self.btn_run.configure(text="Run Now"))
            if not self.scheduler_running:
                self.set_ui(True)

    def start(self):
        if self.scheduler_running:
            return
        self.stop_requested = False
        self.scheduler_running = True
        self.next_run = None
        self.set_ui(False)
        threading.Thread(target=self.main_loop, daemon=True).start()
        self.log("Auto scheduler started")

    def stop(self):
        self.scheduler_running = False
        self.stop_requested = True
        self.status.configure(text="Status: Stopped")
        self.set_ui(True)
        self.btn_start.configure(text="Start")
        self.btn_run.configure(text="Run Now")

    def set_ui(self, enable):
        state = "normal" if enable else "disabled"
        for w in [self.dws_url, self.dws_token, self.jms_token, self.path, self.start_date, self.end_date, self.start_hour, self.end_hour, self.delay, self.name_dws, self.name_auto, self.name_dwspda, self.name_realtime_db, self.btn_browse]:
            w.configure(state=state)
        self.btn_run.configure(state="normal")
        self.btn_start.configure(state="normal")

    def validate(self):
        return bool(self.get_text(self.dws_url) and self.get_text(self.dws_token) and self.get_text(self.jms_token) and self.get_text(self.path))


if __name__ == "__main__":
    app = App()
    app.mainloop()
