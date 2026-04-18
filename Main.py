import tkinter as tk
from tkinter import ttk, filedialog
import threading, time
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import requests
import os
import configparser

CONFIG_FILE = "config.ini"

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bot Export Scada&Jms")
        self.geometry("650x450")
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
        }

        cfg["TIME"] = {
            "run_minute": self.delay.get() or "5"
        }

        with open(CONFIG_FILE, "w") as f:
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
            }

            cfg["TIME"] = {
                "run_minute": "5"
            }

            with open(CONFIG_FILE, "w") as f:
                cfg.write(f)
        # 🔥 โหลดไฟล์

        cfg.read(CONFIG_FILE)

        a = cfg["API"] if cfg.has_section("API") else {}
        # 🔥 set ค่า (ลบก่อน insert กันซ้ำ)
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
        
        t = cfg["TIME"] if cfg.has_section("TIME") else {}
        self.delay.set(t.get("run_minute", "5"))

    # ===== UI =====
    def create_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        self.tab_home = ttk.Frame(notebook)
        self.tab_setting = ttk.Frame(notebook)

        notebook.add(self.tab_home, text="Home")
        notebook.add(self.tab_setting, text="Setting")

        self.build_home()
        self.build_setting()

    # ===== HOME =====
    def build_home(self):
        f = self.tab_home

        # ===== ROW 1 : STATUS =====
        self.status = ttk.Label(f, text="Status: Idle")
        self.status.pack(anchor="w", padx=10, pady=5)

        # ===== FRAME =====
        frame = ttk.Frame(f)
        frame.pack(anchor="w", padx=10, pady=5)

        # ===== DATA =====
        hours = [f"{i:02d}:00" for i in range(24)]
        minutes = [str(i) for i in range(60)]

        # ===== ROW 2 : RUN MINUTE =====
        ttk.Label(frame, text="Run minute").grid(row=0, column=0, sticky="w")

        self.delay = ttk.Combobox(frame, values=minutes, width=5, state="readonly")
        self.delay.grid(row=0, column=1, padx=5)
        self.delay.set("5")

        # ===== ROW : START → END + BUTTON =====
        ttk.Label(frame, text="Start").grid(row=1, column=0, sticky="w")

        self.start_date = DateEntry(frame, width=12, date_pattern="yyyy-mm-dd", state="readonly")
        self.start_date.grid(row=1, column=1)

        self.start_hour = ttk.Combobox(frame, values=hours, width=5, state="readonly")
        self.start_hour.grid(row=1, column=2)
        self.start_hour.set("13:00")

        ttk.Label(frame, text="→").grid(row=1, column=3, padx=5)

        ttk.Label(frame, text="End").grid(row=1, column=4)

        self.end_date = DateEntry(frame, width=12, date_pattern="yyyy-mm-dd", state="readonly")
        self.end_date.grid(row=1, column=5)

        # default date
        today = datetime.now().date()
        self.start_date.set_date(today)
        self.end_date.set_date(today + timedelta(days=1))
        
        self.end_hour = ttk.Combobox(frame, values=hours, width=5, state="readonly")
        self.end_hour.grid(row=1, column=6)
        self.end_hour.set("23:00")

        # ===== BUTTON =====
        self.btn_start = ttk.Button(frame, text="Start", command=self.toggle_start)
        self.btn_start.grid(row=1, column=7, padx=10)

        self.btn_run = ttk.Button(frame, text="Run Now", command=self.toggle_run)
        self.btn_run.grid(row=1, column=8)

        # ===== LOG FRAME =====
        log_frame = ttk.Frame(f)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Scrollbar
        scrollbar = ttk.Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")

        # Text (log)
        self.log_text = tk.Text(log_frame, height=18, yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)

        # connect
        scrollbar.config(command=self.log_text.yview)

    # ===== SETTING =====
    def build_setting(self):
        f = self.tab_setting

        # ===== DWS FRAME =====
        dws_frame = ttk.LabelFrame(f, text="DWS9-11", padding=10)
        dws_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        dws_frame.columnconfigure(1, weight=1)

        ttk.Label(dws_frame, text="API URL").grid(row=0, column=0, sticky="w")
        self.dws_url = ttk.Entry(dws_frame, width=50)
        self.dws_url.grid(row=0, column=1, pady=5, sticky="ew")

        ttk.Label(dws_frame, text="Token").grid(row=1, column=0, sticky="w")
        self.dws_token = ttk.Entry(dws_frame, width=50)
        self.dws_token.grid(row=1, column=1, pady=5, sticky="ew")

        # ===== JMS FRAME =====
        jms_frame = ttk.LabelFrame(f, text="JMS", padding=10)
        jms_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        jms_frame.columnconfigure(1, weight=1)

        ttk.Label(jms_frame, text="AuthToken").grid(row=0, column=0, sticky="w")
        self.jms_token = ttk.Entry(jms_frame, width=50)
        self.jms_token.grid(row=0, column=1, pady=5, sticky="ew")

        # ===== PATH FRAME =====
        path_frame = ttk.LabelFrame(f, text="Output", padding=10)
        path_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        path_frame.columnconfigure(1, weight=1)

        ttk.Label(path_frame, text="Save Path").grid(row=0, column=0, sticky="w")
        self.path = ttk.Entry(path_frame, width=45)
        self.path.grid(row=0, column=1, pady=5, sticky="ew")

        self.btn_browse = ttk.Button(path_frame, text="Browse", command=self.browse)
        self.btn_browse.grid(row=0, column=2, padx=5)

        # ===== OUTPUT FRAME =====
        output_frame = ttk.LabelFrame(f, text="PATH", padding=10)
        output_frame.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
        output_frame.columnconfigure(1, weight=1)

        ttk.Label(output_frame, text="DWS File").grid(row=1, column=0)
        self.name_dws = ttk.Entry(output_frame, width=45)
        self.name_dws.grid(row=1, column=1, sticky="ew")

        ttk.Label(output_frame, text="AUTOPDA File").grid(row=2, column=0)
        self.name_auto = ttk.Entry(output_frame, width=45)
        self.name_auto.grid(row=2, column=1, sticky="ew")

        ttk.Label(output_frame, text="DWSPDA File").grid(row=3, column=0)
        self.name_dwspda = ttk.Entry(output_frame, width=45)
        self.name_dwspda.grid(row=3, column=1, sticky="ew")
        # ให้ stretch
        f.columnconfigure(0, weight=1)

    def toggle_start(self):
        if not self.scheduler_running:
            self.start()
            self.btn_start.config(text="Stop")
        else:
            self.stop()
            self.btn_start.config(text="Start")
    
    def toggle_run(self):
        if not self.job_running:
            self.stop_requested = False
            self.run_now()
            self.btn_run.config(text="Stop")
        else:
            self.stop_requested = True
            self.status.config(text="Status: Stopping...")
            self.btn_run.config(text="Run Now")

    # ===== UTIL =====
    def log(self, msg, tag="SYS"):
        now = datetime.now().strftime("%H:%M:%S")
        text = f"{now} [{tag}] {msg}\n"
        self.log_text.insert("end", text)
        self.log_text.see("end")

    def browse(self):
        p = filedialog.askdirectory()
        if p:
            self.path.delete(0, tk.END)
            self.path.insert(0, p)

    # ====== ใส่ LOGIC จริงของคุณตรงนี้ ======
    def run_dws(self):
        self.log("Start DWS request...", "DWS9-11")

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

            res = requests.post(url, json=data, headers=headers, timeout=60)

            # DEBUG
            self.log(f"STATUS: {res.status_code}", "DWS9-11")
            self.log(f"SIZE: {len(res.content)}", "DWS9-11")
            #self.log(f"RESP: {res.text[:200]}", "DWS9-11")

            if res.status_code == 200 and len(res.content) > 100:
                os.makedirs(self.path.get(), exist_ok=True)
                path = os.path.join(self.path.get(), self.name_dws.get())
                with open(path, "wb") as f:
                    f.write(res.content)

                self.log(f"Downloaded: {self.name_dws.get()}")
            else:
                self.after(0, lambda: self.log(f"Download failed: {res.text[:200]}"))
                self.after(0, lambda: self.status.config(text="Error"))

        except Exception as e:
            self.after(0, lambda: self.log(f"ERROR: {e}"))
            self.after(0, lambda: self.status.config(text="Error"))

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

    def _export_jms(self, base, headers, start, end, scanType, filename):
        self.log(f"Generate {filename}", "JMS")

        run_time = datetime.now()

        target_hour = run_time.replace(minute=0, second=0, microsecond=0)
        next_hour = target_hour + timedelta(hours=1)

        requests.post(f"{base}/scanningContrast/asyncDownExcel", json={
            "startTimeStr": start,
            "endTimeStr": end,
            "scanNetworkCode": "999004",
            "scanType": scanType,
            "excelType": "downExcelAll"
        }, headers=headers)

        if not self.sleep_with_stop(60):
            return

        for _ in range(12):
            if self.stop_requested:
                raise Exception("Stopped by user")

            if not self.sleep_with_stop(10):
                return

            res = requests.post(f"{base}/downLoadCenter/downLoadInfoList", json={
                "current": 1,
                "size": 20
            }, headers=headers).json()

            records = res.get("data", {}).get("records", [])
            records.sort(key=lambda x: x.get("downTime", ""), reverse=True)

            for r in records:
                try:
                    down_time = datetime.strptime(
                        r.get("downTime"), "%Y-%m-%d %H:%M:%S"
                    )
                except:
                    continue

                self.log(f"Check: {r.get('downTime')} | finish={r.get('finishOrNot')}", "DEBUG")

                if (
                    scanType in r.get("queryJson", "")
                    and r.get("finishOrNot") == "1"
                    and r.get("downUrl")
                    and target_hour <= down_time < next_hour
                    and down_time >= run_time - timedelta(minutes=7)
                ):
                    self.log(f"Matched file: {down_time}", "DEBUG")

                    sign = requests.post(
                        f"{base}/downLoadCenter/getDownloadSignedUrl",
                        json=r,
                        headers=headers
                    ).json()

                    url = sign.get("data")
                    if not url:
                        continue

                    file = requests.get(url if isinstance(url, str) else url[0])

                    if len(file.content) < 100_000:
                        self.log("File not ready, retry...", "JMS")
                        continue

                    os.makedirs(self.path.get(), exist_ok=True)
                    path = os.path.join(self.path.get(), filename)

                    with open(path, "wb") as f:
                        f.write(file.content)

                    self.log(f"Downloaded {filename}", "JMS")
                    return

        raise Exception(f"{filename} fail")
        
    # ===== CORE RUN =====
    def run_all(self):
        self.job_running = True
        self.save_config()
        self.set_ui(False)
        self.status.config(text="Status: Running")

        try:
            if self.stop_requested:
                return

            self.run_dws()

            if not self.sleep_with_stop(10):
                return

            if self.stop_requested:
                self.log("Stopped by user", "SYS")
                return

            self.run_jms_auto()

            if not self.sleep_with_stop(10):
                return

            self.run_jms_pda()

            if not self.stop_requested:
                self.status.config(text="Status: Success")

        except Exception as e:
            self.log(f"ERROR: {e}", "SYS")
            self.status.config(text="Status: Error")

        finally:
            self.job_running = False

            if not self.scheduler_running:   # 🔥 เพิ่มเงื่อนไขนี้
                self.set_ui(True)

            self.btn_run.config(text="Run Now")

            if self.stop_requested:
                self.status.config(text="Status: Stopped")

            now = datetime.now()
            run_minute = int(self.delay.get())

            next_run = now.replace(minute=run_minute, second=0, microsecond=0)

            if now >= next_run:
                next_run = next_run + timedelta(hours=1)

            self.log(f"Next run at {next_run.strftime('%H:%M:%S')}", "SYS")

    def run_now(self):
        if not self.validate():
            self.log("Missing config", "SYS")
            return

        if not self.validate_time():
            self.log("Invalid time range (Start > End)", "SYS")
            return

        if self.job_running:
            return

        self.job_running = True   # 🔥 เพิ่มตรงนี้

        threading.Thread(target=self.run_all, daemon=True).start()

    # ===== SCHEDULER =====
    def scheduler(self):
        while self.scheduler_running:

            now = datetime.now()
            run_minute = int(self.delay.get())

            target = now.replace(minute=run_minute, second=0, microsecond=0)

            if now >= target:
                target = target + timedelta(hours=1)

            wait_seconds = int((target - now).total_seconds())

            if wait_seconds <= 0:
                time.sleep(1)
                continue

            self.log(f"Next run at {target.strftime('%H:%M:%S')}", "SYS")

            for _ in range(wait_seconds):
                if not self.scheduler_running:
                    return
                time.sleep(1)

            if not self.validate():
                self.log("Missing config", "SYS")
                continue

            if not self.job_running:
                self.run_now()

    def get_time_range(self):
        start_str = f"{self.start_date.get()} {self.start_hour.get()}"
        end_str   = f"{self.end_date.get()} {self.end_hour.get()}"

        start = datetime.strptime(start_str, "%Y-%m-%d %H:%M")
        end   = datetime.strptime(end_str, "%Y-%m-%d %H:%M")

        # 🔥 FIX สำคัญ: auto swap
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

    def start(self):
        if self.scheduler_running:
            return

        self.scheduler_running = True

        self.set_ui(False)   # 🔥 เพิ่มตรงนี้

        threading.Thread(target=self.scheduler, daemon=True).start()
        self.log("Auto scheduler started")
        self.btn_start.config(text="Stop")

    def stop(self):
        self.scheduler_running = False
        self.stop_requested = True

        self.status.config(text="Status: Stopped")

        self.set_ui(True)   # 🔥 เปิด UI กลับแน่นอน

        self.btn_start.config(text="Start")
        self.btn_run.config(text="Run Now")
        
    def set_ui(self, enable):
        state = "normal" if enable else "disabled"

        # disable input
        self.dws_url.config(state=state)
        self.dws_token.config(state=state)
        self.jms_token.config(state=state)
        self.path.config(state=state)

        self.start_date.config(state=state)
        self.end_date.config(state=state)
        self.start_hour.config(state=state)
        self.end_hour.config(state=state)
        self.delay.config(state=state)

        self.name_dws.config(state=state)
        self.name_auto.config(state=state)
        self.name_dwspda.config(state=state)

        self.btn_browse.config(state=state)

        # 🔥 ปุ่มควบคุม = กดได้ตลอด
        self.btn_run.config(state="normal")
        self.btn_start.config(state="normal")


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