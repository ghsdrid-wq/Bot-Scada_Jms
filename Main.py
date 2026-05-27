import tkinter as tk
from tkinter import filedialog
import threading, time
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import requests
import os
import configparser
import customtkinter as ctk

CONFIG_FILE = "config.ini"

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Bot Export Scada&Jms")
        self.geometry("900x660")
        self.resizable(False, False)
        self.scheduler_running = False
        self.job_running = False
        self.stop_requested = False

        self.create_ui()
        self.load_config()
        
    def save_config(self):
        cfg = configparser.ConfigParser()

        cfg["API"] = {
            "dws_url": self.dws_url.get(),
            "dws_token": self.dws_token.get(),
            "jms_token": self.jms_token.get(),
        }

        cfg["PATH"] = {
            "path": self.path.get(),
        }

        cfg["FILE"] = {
            "name_dws": self.name_dws.get() or "DWS9-11.xlsx",
            "name_auto": self.name_auto.get() or "DWSXAUTOPDA.xlsx",
            "name_dwspda": self.name_dwspda.get() or "DWSPDA.xlsx",
            "name_realtime_db": self.name_realtime_db.get() or "RealtimeDB.xlsx",
        }

        cfg["TIME"] = {
            "run_minute": self.delay.get() or "5"
        }

        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            cfg.write(f)
        
    def load_config(self):
        cfg = configparser.ConfigParser()

        if not os.path.exists(CONFIG_FILE):
            cfg["API"] = {
                "dws_url": "",
                "dws_token": "",
                "jms_token": "",
            }

            cfg["PATH"] = {
                "path": ""
            }

            cfg["FILE"] = {
                "name_dws": "DWS9-11.xlsx",
                "name_auto": "DWSXAUTOPDA.xlsx",
                "name_dwspda": "DWSPDA.xlsx",
                "name_realtime_db": "RealtimeDB.xlsx",
            }

            cfg["TIME"] = {
                "run_minute": "5"
            }

            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                cfg.write(f)
        # โหลดไฟล์

        cfg.read(CONFIG_FILE, encoding="utf-8")

        a = cfg["API"] if cfg.has_section("API") else {}
        # set ค่า (ลบก่อน insert กันซ้ำ)
        self.dws_url.delete(0, tk.END)
        self.dws_url.insert(0, a.get("dws_url", ""))
        self.dws_token.delete(0, tk.END)
        self.dws_token.insert(0, a.get("dws_token", ""))
        self.jms_token.delete(0, tk.END)
        self.jms_token.insert(0, a.get("jms_token", ""))

        p = cfg["PATH"] if cfg.has_section("PATH") else {}
        self.path.delete(0, tk.END)
        self.path.insert(0, p.get("path", ""))

        f = cfg["FILE"] if cfg.has_section("FILE") else {}
        self.name_dws.delete(0, tk.END)
        self.name_dws.insert(0, f.get("name_dws", "DWS9-11.xlsx"))
        self.name_auto.delete(0, tk.END)
        self.name_auto.insert(0, f.get("name_auto", "DWSXAUTOPDA.xlsx"))
        self.name_dwspda.delete(0, tk.END)
        self.name_dwspda.insert(0, f.get("name_dwspda", "DWSPDA.xlsx"))
        self.name_realtime_db.delete(0, tk.END)
        self.name_realtime_db.insert(0,f.get("name_realtime_db", "RealtimeDB.xlsx"))
        
        t = cfg["TIME"] if cfg.has_section("TIME") else {}
        self.delay.set(t.get("run_minute", "5"))

    # ===== UI =====
    def create_ui(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=12, pady=12)

        title = ctk.CTkLabel(
            container,
            text="Bot Export Scada & JMS",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.pack(anchor="w", pady=(0, 8))

        notebook = ctk.CTkTabview(
            container,
            width=860,
            height=500,
            anchor="w"   # <- เพิ่ม
        )
        notebook.pack(fill="both", expand=True)

        self.tab_home = notebook.add("Home")
        self.tab_setting = notebook.add("Setting")

        self.build_home()
        self.build_setting()

    # ===== HOME =====
    def build_home(self):
        f = self.tab_home

        # ===== ROW 1 : STATUS =====
        self.status = ctk.CTkLabel(f, text="Status: Idle", font=ctk.CTkFont(size=14, weight="bold"))
        self.status.pack(anchor="w", padx=16, pady=(14, 8))

        # ===== FRAME =====
        frame = ctk.CTkFrame(f)
        frame.pack(fill="x", padx=16, pady=8)

        # ===== DATA =====
        hours = [f"{i:02d}:00" for i in range(24)]
        minutes = [str(i) for i in range(60)]

        # ===== ROW 2 : RUN MINUTE =====
        ctk.CTkLabel(frame, text="Run minute").grid(row=0, column=0, sticky="w", padx=8, pady=8)

        self.delay = ctk.CTkComboBox(frame, values=minutes, width=80, state="readonly")
        self.delay.grid(row=0, column=1, padx=8, pady=8, sticky="w")
        self.delay.set("5")

        # ===== ROW : START → END + BUTTON =====
        ctk.CTkLabel(frame, text="Start").grid(row=1, column=0, sticky="w", padx=8, pady=8)

        self.start_date = DateEntry(
            frame,
            width=12,
            date_pattern="yyyy-mm-dd",
            state="readonly",
            font=("Segoe UI", 13),
        )
        self.start_date.grid(row=1, column=1)

        self.start_hour = ctk.CTkComboBox(frame, values=hours, width=90, state="readonly")
        self.start_hour.grid(row=1, column=2, padx=8)
        self.start_hour.set("13:00")

        ctk.CTkLabel(frame, text="→").grid(row=1, column=3, padx=8)

        ctk.CTkLabel(frame, text="End").grid(row=1, column=4, padx=8)

        self.end_date = DateEntry(
            frame,
            width=12,
            date_pattern="yyyy-mm-dd",
            state="readonly",
            font=("Segoe UI", 13),
        )
        self.end_date.grid(row=1, column=5)

        # default date
        today = datetime.now().date()
        self.start_date.set_date(today)
        self.end_date.set_date(today + timedelta(days=1))
        
        self.end_hour = ctk.CTkComboBox(frame, values=hours, width=90, state="readonly")
        self.end_hour.grid(row=1, column=6, padx=8)
        self.end_hour.set("23:00")

        # ===== BUTTON =====
        self.btn_start = ctk.CTkButton(frame, text="Start", command=self.toggle_start, width=100)
        self.btn_start.grid(row=1, column=7, padx=10)

        self.btn_run = ctk.CTkButton(frame, text="Run Now", command=self.toggle_run, width=110)
        self.btn_run.grid(row=1, column=8)

        # ===== LOG FRAME =====
        log_frame = ctk.CTkFrame(f)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(8, 16))

        self.log_text = ctk.CTkTextbox(log_frame, height=300)
        self.log_text.pack(fill="both", expand=True, padx=8, pady=8)

    # ===== SETTING =====
    def build_setting(self):
        f = self.tab_setting

        # ===== DWS FRAME =====
        dws_frame = ctk.CTkFrame(f)
        dws_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        dws_frame.columnconfigure(1, weight=1)
        ctk.CTkLabel(dws_frame, text="DWS9-11", font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))

        ctk.CTkLabel(dws_frame, text="API URL").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.dws_url = ctk.CTkEntry(dws_frame, width=50)
        self.dws_url.grid(row=1, column=1, pady=5, padx=(0, 10), sticky="ew")

        ctk.CTkLabel(dws_frame, text="Token").grid(row=2, column=0, sticky="w", padx=10, pady=(0, 10))
        self.dws_token = ctk.CTkEntry(dws_frame, width=50)
        self.dws_token.grid(row=2, column=1, pady=(0, 10), padx=(0, 10), sticky="ew")

        # ===== JMS FRAME =====
        jms_frame = ctk.CTkFrame(f)
        jms_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        jms_frame.columnconfigure(1, weight=1)
        ctk.CTkLabel(jms_frame, text="JMS", font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))

        ctk.CTkLabel(jms_frame, text="AuthToken").grid(row=1, column=0, sticky="w", padx=10, pady=(5, 10))
        self.jms_token = ctk.CTkEntry(jms_frame, width=50)
        self.jms_token.grid(row=1, column=1, pady=(5, 10), padx=(0, 10), sticky="ew")

        # ===== PATH FRAME =====
        path_frame = ctk.CTkFrame(f)
        path_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        path_frame.columnconfigure(1, weight=1)
        ctk.CTkLabel(path_frame, text="Output", font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))

        ctk.CTkLabel(path_frame, text="Save Path").grid(row=1, column=0, sticky="w", padx=10, pady=(5, 10))
        self.path = ctk.CTkEntry(path_frame, width=45)
        self.path.grid(row=1, column=1, pady=(5, 10), sticky="ew")

        self.btn_browse = ctk.CTkButton(path_frame, text="Browse", command=self.browse, width=90)
        self.btn_browse.grid(row=1, column=2, padx=8, pady=(5, 10))

        # ===== OUTPUT FRAME =====
        output_frame = ctk.CTkFrame(f)
        output_frame.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
        output_frame.columnconfigure(1, weight=1)
        ctk.CTkLabel(output_frame, text="File Names", font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))

        ctk.CTkLabel(output_frame, text="DWS File").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.name_dws = ctk.CTkEntry(output_frame, width=45)
        self.name_dws.grid(row=1, column=1, padx=(0, 10), sticky="ew")

        ctk.CTkLabel(output_frame, text="AUTOPDA File").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.name_auto = ctk.CTkEntry(output_frame, width=45)
        self.name_auto.grid(row=2, column=1, padx=(0, 10), sticky="ew")

        ctk.CTkLabel(output_frame, text="DWSPDA File").grid(row=3, column=0, padx=10, pady=(5, 10), sticky="w")
        self.name_dwspda = ctk.CTkEntry(output_frame, width=45)
        self.name_realtime_db = ctk.CTkEntry(output_frame,width=45)
        self.name_dwspda.grid(row=3, column=1, padx=(0, 10), pady=(5, 10), sticky="ew")
        
        ctk.CTkLabel(
            output_frame,
            text="RealtimeDB File"
        ).grid(
            row=4,
            column=0,
            padx=10,
            pady=(5, 10),
            sticky="w"
        )

        self.name_realtime_db.grid(
            row=4,
            column=1,
            padx=(0, 10),
            pady=(5, 10),
            sticky="ew"
        )
        # ให้ stretch
        f.columnconfigure(0, weight=1)

    def toggle_start(self):
        if not self.scheduler_running:
            self.start()
            self.btn_start.configure(text="Stop")
        else:
            self.stop()
            self.btn_start.configure(text="Start")
    
    def toggle_run(self):
        if self.job_running:
            self.stop_requested = True
            self.status.configure(text="Status: Stopping...")
            self.btn_run.configure(text="Run Now")
            return

        self.stop_requested = False
        threading.Thread(target=self.execute_pipeline, daemon=True).start()
        self.btn_run.configure(text="Stop")

    # ===== UTIL =====
    def log(self, msg, tag="SYS", level="INFO"):
        now = datetime.now().strftime("%H:%M:%S")

        icons = {
            "INFO": "●",
            "SUCCESS": "✓",
            "WARN": "▲",
            "ERROR": "✕",
        }

        icon = icons.get(level, "●")

        self.log_text.insert(
            "end",
            f"{now}  {icon}  [{tag}]  {msg}\n"
        )

        self.log_text.see("end")

    def browse(self):
        p = filedialog.askdirectory()
        if p:
            self.path.delete(0, tk.END)
            self.path.insert(0, p)

    # ====== ใส่ LOGIC จริงของคุณตรงนี้ ======
    def run_dws(self):
        self.log("Requesting export file", "DWS")

        start, end = self.get_time_range()

        try:
            url = self.dws_url.get().strip()
            token = self.dws_token.get().strip()
            data = {
                "startTime": start,
                "endTime": end,
                "barcodeList": [],
                "businessTimeType": "",
                "dbCode": "2",
                "curPage": 1,
                "pageSize": 20
            }
            headers = {
                "Content-Type": "application/json;charset=UTF-8",
                "token": token,
                "Origin": "http://10.30.32.10",
                "Referer": "http://10.30.32.10/",
                "User-Agent": "Mozilla/5.0"
            }

            res = requests.post(url, json=data, headers=headers, timeout=(5, 10))

            # DEBUG
            self.log(f"STATUS: {res.status_code}", "DWS9-11")
            self.log(f"SIZE: {len(res.content)}", "DWS9-11")
            #self.log(f"RESP: {res.text[:200]}", "DWS9-11")

            if res.status_code == 200 and len(res.content) > 100:
                os.makedirs(self.path.get(), exist_ok=True)
                path = os.path.join(self.path.get(), self.name_dws.get())
                with open(path, "wb") as f:
                    f.write(res.content)

                self.log(
                    f"File saved → {self.name_dws.get()}",
                    "DWS",
                    "SUCCESS"
                )
            else:
                self.after(0, lambda: self.log(f"Download failed: {res.text[:200]}"))
                self.after(0, lambda: self.status.configure(text="Error"))

        except Exception as e:
            self.after(0, lambda: self.log(f"ERROR: {e}"))
            self.after(0, lambda: self.status.configure(text="Error"))

        finally:
            self.log("DWS finished", "SYS")


    def run_jms_auto(self):
        base = "https://jmsgw.jtexpress.co.th/operatingplatform"
        token = self.jms_token.get().strip()

        headers = {
            "Content-Type": "application/json;charset=UTF-8",
            "authtoken": token
        }

        start, end = self.get_time_range()

        self._export_jms(base, headers, start, end, "建包扫描", self.name_auto.get())


    def run_jms_pda(self):
        base = "https://jmsgw.jtexpress.co.th/operatingplatform"
        token = self.jms_token.get().strip()

        headers = {
            "Content-Type": "application/json;charset=UTF-8",
            "authtoken": token
        }

        start, end = self.get_time_range()

        self._export_jms(base, headers, start, end, "卸车扫描", self.name_dwspda.get())

    def run_realtime_db(self):
        self.log(
            "Creating export task",
            "REALTIME"
        )

        token = self.jms_token.get().strip()

        headers = {
            "Content-Type": "application/json;charset=UTF-8",
            "authtoken": token,
            "origin": "https://jms.jtexpress.co.th",
            "referer": "https://jms.jtexpress.co.th/",
            "routename": "TrackRealTimeMonitoringDB",
            "lang": "TH",
            "langtype": "TH",
        }

        base = "https://jmsgw.jtexpress.co.th"

        export_url = (
            f"{base}"
            "/businessindicator/bigdataReport/"
            "pageExcelByTask/trail_monitor_detail_doris"
        )

        export_payload = {
            "orderSourceCode": ["JMS"],
            "scanCode": "999004",
            "scanFranCode": "555090",
            "scanAgentCode": "555090",
            "countryId": "1",
            "modelName": "RealtimeDB"
        }

        res = requests.post(
            export_url,
            json=export_payload,
            headers=headers,
            timeout=(5, 10)
        )

        if "login" in res.text.lower():
            raise Exception("Realtime DB token expired")

        self.log(
            "Export task created",
            "REALTIME",
            "SUCCESS"
        )

        if not self.sleep_with_stop(5):
            return

        list_url = (
            f"{base}"
            "/businessindicator/bigdataReport/report/file/list"
        )

        today = datetime.now()

        list_payload = {
            "current": 1,
            "size": 20,
            "startTime": today.strftime("%Y-%m-%d 00:00:00"),
            "endTime": today.strftime("%Y-%m-%d 23:59:59"),
            "countryId": "1",
            "networkCode": "999004",
            "userId": 13144940
        }

        download_url = None

        for _ in range(30):

            if self.stop_requested:
                return

            res = requests.post(
                list_url,
                json=list_payload,
                headers=headers,
                timeout=(5, 10)
            )

            if "login" in res.text.lower():
                raise Exception("Realtime DB token expired")

            data = res.json()

            records = data.get("data", {}).get("list", [])

            records.sort(
                key=lambda x: x.get("createTime", ""),
                reverse=True
            )

            for record in records:

                status = record.get("status")
                down_url = record.get("downUrl")
                business = record.get("business", "")

                if (
                    status == 2
                    and down_url
                    and business == "trail_monitor_detail_doris"
                ):

                    download_url = down_url

                    self.log(
                        "Download link ready",
                        "REALTIME",
                        "SUCCESS"
                    )

                    break

            if download_url:
                break

            self.log(
                "Waiting for file generation",
                "REALTIME"
            )

            if not self.sleep_with_stop(3):
                return

        if not download_url:
            raise Exception("Realtime DB download url not found")

        file = requests.get(
            download_url,
            timeout=(5, 30)
        )

        if len(file.content) < 1000:
            raise Exception("Realtime DB file invalid")

        os.makedirs(self.path.get(), exist_ok=True)

        save_path = os.path.join(
            self.path.get(),
            self.name_realtime_db.get()
        )

        with open(save_path, "wb") as f:
            f.write(file.content)

        self.log(
            f"File saved → {self.name_realtime_db.get()}",
            "REALTIME",
            "SUCCESS"
        )

    def _export_jms(self, base, headers, start, end, scanType, filename):
        self.log(
            f"Creating export task → {filename}",
            "JMS"
        )

        run_time = datetime.now()

        target_hour = run_time.replace(minute=0, second=0, microsecond=0)
        next_hour = target_hour + timedelta(hours=1)

        res = requests.post(f"{base}/scanningContrast/asyncDownExcel", json={
            "startTimeStr": start,
            "endTimeStr": end,
            "scanNetworkCode": "999004",
            "scanType": scanType,
            "excelType": "downExcelAll"
        }, headers=headers, timeout=(5, 10))

        if "login" in res.text.lower():
            self.log("JMS Token expired (generate)", "ERROR")
            raise Exception("Token expired")

        if not self.sleep_with_stop(50):
            return

        for _ in range(24):
            if self.stop_requested:
                self.log("Stopped by user")
                return

            if not self.sleep_with_stop(10):
                return

            res = requests.post(f"{base}/downLoadCenter/downLoadInfoList", json={
                "current": 1,
                "size": 20
            }, headers=headers, timeout=(5, 10))

            # CHECK TOKEN
            if "login" in res.text.lower():
                self.log("JMS Token expired (list)", "ERROR")
                raise Exception("Token expired")

            res = res.json()

            records = res.get("data", {}).get("records", [])
            records.sort(key=lambda x: x.get("downTime", ""), reverse=True)

            for r in records:
                try:
                    down_time = datetime.strptime(
                        r.get("downTime"), "%Y-%m-%d %H:%M:%S"
                    )
                except:
                    continue


                if (
                    scanType in r.get("queryJson", "")
                    and r.get("finishOrNot") == "1"
                    and r.get("downUrl")
                    and target_hour <= down_time < next_hour
                    and down_time >= run_time - timedelta(hours=1)
                ):

                    sign_res = requests.post(
                        f"{base}/downLoadCenter/getDownloadSignedUrl",
                        json=r,
                        headers=headers,
                        timeout=(5, 10)
                    )

                    if "login" in sign_res.text.lower():
                        self.log("JMS Token expired (signedUrl)", "ERROR")
                        raise Exception("Token expired")

                    sign = sign_res.json()

                    url = sign.get("data")
                    if not url:
                        continue

                    file = requests.get(
                        url if isinstance(url, str) else url[0],
                        timeout=(5,10)
                    )

                    # CHECK TOKEN / HTML
                    content_type = file.headers.get("Content-Type", "").lower()

                    if "html" in content_type or "login" in file.text.lower():
                        self.log("JMS Token expired (download)", "ERROR")
                        raise Exception("Token expired")

                    if len(file.content) < 100_000:
                        self.log("File not ready, retry...", "JMS")
                        continue

                    os.makedirs(self.path.get(), exist_ok=True)
                    path = os.path.join(self.path.get(), filename)

                    with open(path, "wb") as f:
                        f.write(file.content)

                    self.log(
                        f"File saved → {filename}",
                        "JMS",
                        "SUCCESS"
                    )
                    return

        raise Exception(f"{filename} fail")
        
    def run_now(self):
        if not self.validate():
            self.log("Missing config", "SYS")
            return

        if not self.validate_time():
            self.log("Invalid time range (Start > End)", "SYS")
            return

        if self.job_running:
            return
        
        threading.Thread(target=self.execute_pipeline, daemon=True).start()

    def get_time_range(self):
        start_str = f"{self.start_date.get()} {self.start_hour.get()}"
        end_str   = f"{self.end_date.get()} {self.end_hour.get()}"

        start = datetime.strptime(start_str, "%Y-%m-%d %H:%M")
        end   = datetime.strptime(end_str, "%Y-%m-%d %H:%M")

        # FIX สำคัญ: auto swap
        if start > end:
            start, end = end, start

        return (
            start.strftime("%Y-%m-%d %H:%M:%S"),
            end.strftime("%Y-%m-%d %H:%M:%S")
        )

    def sleep_with_stop(self, seconds):
        for _ in range(seconds):
            if self.stop_requested:
                return False
            time.sleep(1)
        return True

        # ===== NEW MAIN LOOP =====
    def main_loop(self):
        self.log("Main loop started", "SYS")

        run_minute = int(self.delay.get())

        # init รอบแรก
        if not hasattr(self, "next_run") or self.next_run is None:
            now = datetime.now()
            self.next_run = now.replace(minute=run_minute, second=0, microsecond=0)

            if now > self.next_run + timedelta(seconds=30):
                self.next_run += timedelta(hours=1)

        while self.scheduler_running:
            self.log(f"Next run at {self.next_run.strftime('%H:%M:%S')}", "SYS")

            # รอเวลา
            while datetime.now() < self.next_run:
                if not self.scheduler_running:
                    return
                time.sleep(1)

            # ยิงงาน
            if not self.job_running:
                self.log("Run job", "SYS")
                threading.Thread(target=self.execute_pipeline, daemon=True).start()
            else:
                self.log("Skip: previous job still running", "SYS")

            # ไปชั่วโมงถัดไป
            self.next_run += timedelta(hours=1)

    # ===== NEW PIPELINE =====
    def execute_pipeline(self):
        if self.job_running:
            self.log("Skip: job running", "SYS")
            return

        self.job_running = True
        self.save_config()
        self.set_ui(False)
        self.status.configure(text="Status: Running")

        try:
            if self.stop_requested:
                return
        
            self.run_dws()

            if self.stop_requested:
                return

            if not self.sleep_with_stop(5):
                raise Exception("Stopped")
            
            if self.stop_requested:
                return

            self.run_jms_auto()

            if self.stop_requested:
                return

            if not self.sleep_with_stop(5):
                raise Exception("Stopped")
            
            if self.stop_requested:
                return

            self.run_jms_pda()

            if self.stop_requested:
                return

            if not self.sleep_with_stop(5):
                raise Exception("Stopped")

            self.run_realtime_db()


            self.status.configure(text="Status: Success")

        except Exception as e:
            self.log(f"ERROR: {e}", "SYS")
            self.status.configure(text="Status: Error")

        finally:
            self.job_running = False
            self.stop_requested = False
            if self.scheduler_running and hasattr(self, "next_run") and self.next_run:
                self.log("Finished", "SYS")
                self.log(f"Next run at {self.next_run.strftime('%H:%M:%S')}", "SYS")

            self.after(0, lambda: self.btn_run.configure(text="Run Now"))

            # reset ปุ่ม Run Now
            if not self.scheduler_running:
                self.set_ui(True)

    def start(self):
        if self.scheduler_running:
            return

        self.stop_requested = False
        self.scheduler_running = True
        self.next_run = None

        self.set_ui(False)   # เพิ่มตรงนี้

        threading.Thread(target=self.main_loop, daemon=True).start()
        self.log("Auto scheduler started")
        self.btn_start.configure(text="Stop")

    def stop(self):
        self.scheduler_running = False
        self.stop_requested = True

        self.status.configure(text="Status: Stopped")

        self.set_ui(True)   # เปิด UI กลับแน่นอน

        self.btn_start.configure(text="Start")
        self.btn_run.configure(text="Run Now")
        
    def set_ui(self, enable):
        state = "normal" if enable else "disabled"

        # disable input
        self.dws_url.configure(state=state)
        self.dws_token.configure(state=state)
        self.jms_token.configure(state=state)
        self.path.configure(state=state)

        self.start_date.configure(state=state)
        self.end_date.configure(state=state)
        self.start_hour.configure(state=state)
        self.end_hour.configure(state=state)
        self.delay.configure(state=state)

        self.name_dws.configure(state=state)
        self.name_auto.configure(state=state)
        self.name_dwspda.configure(state=state)
        self.name_realtime_db.configure(state=state)

        self.btn_browse.configure(state=state)

        # ปุ่มควบคุม = กดได้ตลอด
        self.btn_run.configure(state="normal")
        self.btn_start.configure(state="normal")


    def validate(self):
        if not self.dws_url.get():
            return False
        if not self.dws_token.get():
            return False
        if not self.jms_token.get():
            return False
        if not self.path.get():
            return False
        return True
    
    def validate_time(self):
        start = datetime.strptime(f"{self.start_date.get()} {self.start_hour.get()}", "%Y-%m-%d %H:%M")
        end   = datetime.strptime(f"{self.end_date.get()} {self.end_hour.get()}", "%Y-%m-%d %H:%M")

        return start <= end
# ===== RUN =====
if __name__ == "__main__":
    App().mainloop()