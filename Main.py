import configparser
import os
import threading
import time
from datetime import datetime, timedelta

import requests
from PySide6.QtCore import QDate, Qt, QTimer
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QPushButton,
    QPlainTextEdit,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

CONFIG_FILE = "config.ini"


class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bot Export Scada&Jms")
        self.resize(700, 520)

        self.scheduler_running = False
        self.job_running = False
        self.stop_requested = False
        self.log_queue = []

        self.create_ui()
        self.load_config()

        self.log_timer = QTimer(self)
        self.log_timer.timeout.connect(self.flush_log)
        self.log_timer.start(200)

    def create_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        main_layout = QVBoxLayout(root)

        self.notebook = QTabWidget()
        main_layout.addWidget(self.notebook)

        self.tab_home = QWidget()
        self.tab_setting = QWidget()
        self.notebook.addTab(self.tab_home, "Home")
        self.notebook.addTab(self.tab_setting, "Setting")

        self.build_home()
        self.build_setting()

    def build_home(self):
        layout = QVBoxLayout(self.tab_home)

        self.status = QLabel("Status: Idle")
        layout.addWidget(self.status)

        row = QGridLayout()
        layout.addLayout(row)

        hours = [f"{i:02d}:00" for i in range(24)]
        minutes = [str(i) for i in range(60)]

        row.addWidget(QLabel("Run minute"), 0, 0)
        self.delay = QComboBox()
        self.delay.addItems(minutes)
        self.delay.setCurrentText("5")
        row.addWidget(self.delay, 0, 1)

        row.addWidget(QLabel("Start"), 1, 0)
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDisplayFormat("yyyy-MM-dd")
        self.start_date.setDate(QDate.currentDate())
        row.addWidget(self.start_date, 1, 1)

        self.start_hour = QComboBox()
        self.start_hour.addItems(hours)
        self.start_hour.setCurrentText("13:00")
        row.addWidget(self.start_hour, 1, 2)

        row.addWidget(QLabel("→"), 1, 3)
        row.addWidget(QLabel("End"), 1, 4)

        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDisplayFormat("yyyy-MM-dd")
        self.end_date.setDate(QDate.currentDate().addDays(1))
        row.addWidget(self.end_date, 1, 5)

        self.end_hour = QComboBox()
        self.end_hour.addItems(hours)
        self.end_hour.setCurrentText("23:00")
        row.addWidget(self.end_hour, 1, 6)

        self.btn_start = QPushButton("Start")
        self.btn_start.clicked.connect(self.toggle_start)
        row.addWidget(self.btn_start, 1, 7)

        self.btn_run = QPushButton("Run Now")
        self.btn_run.clicked.connect(self.toggle_run)
        row.addWidget(self.btn_run, 1, 8)

        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

    def build_setting(self):
        layout = QVBoxLayout(self.tab_setting)

        dws_group = QGroupBox("DWS9-11")
        dws_form = QFormLayout(dws_group)
        self.dws_url = QLineEdit()
        self.dws_token = QLineEdit()
        dws_form.addRow("API URL", self.dws_url)
        dws_form.addRow("Token", self.dws_token)
        layout.addWidget(dws_group)

        jms_group = QGroupBox("JMS")
        jms_form = QFormLayout(jms_group)
        self.jms_token = QLineEdit()
        jms_form.addRow("AuthToken", self.jms_token)
        layout.addWidget(jms_group)

        path_group = QGroupBox("Output")
        path_layout = QHBoxLayout(path_group)
        self.path = QLineEdit()
        self.btn_browse = QPushButton("Browse")
        self.btn_browse.clicked.connect(self.browse)
        path_layout.addWidget(QLabel("Save Path"))
        path_layout.addWidget(self.path)
        path_layout.addWidget(self.btn_browse)
        layout.addWidget(path_group)

        output_group = QGroupBox("PATH")
        output_form = QFormLayout(output_group)
        self.name_dws = QLineEdit()
        self.name_auto = QLineEdit()
        self.name_dwspda = QLineEdit()
        output_form.addRow("DWS File", self.name_dws)
        output_form.addRow("AUTOPDA File", self.name_auto)
        output_form.addRow("DWSPDA File", self.name_dwspda)
        layout.addWidget(output_group)

        layout.addStretch()

    def log(self, msg, tag="SYS"):
        now = datetime.now().strftime("%H:%M:%S")
        self.log_queue.append(f"{now} [{tag}] {msg}")

    def flush_log(self):
        if not self.log_queue:
            return
        self.log_text.appendPlainText("\n".join(self.log_queue))
        self.log_queue.clear()

    def browse(self):
        p = QFileDialog.getExistingDirectory(self, "Select Folder")
        if p:
            self.path.setText(p)

    def save_config(self):
        cfg = configparser.ConfigParser()
        cfg["API"] = {
            "dws_url": self.dws_url.text(),
            "dws_token": self.dws_token.text(),
            "jms_token": self.jms_token.text(),
        }
        cfg["PATH"] = {"path": self.path.text()}
        cfg["FILE"] = {
            "name_dws": self.name_dws.text() or "DWS9-11.xlsx",
            "name_auto": self.name_auto.text() or "DWSXAUTOPDA.xlsx",
            "name_dwspda": self.name_dwspda.text() or "DWSPDA.xlsx",
        }
        cfg["TIME"] = {"run_minute": self.delay.currentText() or "5"}
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            cfg.write(f)

    def load_config(self):
        cfg = configparser.ConfigParser()
        if not os.path.exists(CONFIG_FILE):
            cfg["API"] = {"dws_url": "", "dws_token": "", "jms_token": ""}
            cfg["PATH"] = {"path": ""}
            cfg["FILE"] = {"name_dws": "DWS9-11.xlsx", "name_auto": "DWSXAUTOPDA.xlsx", "name_dwspda": "DWSPDA.xlsx"}
            cfg["TIME"] = {"run_minute": "5"}
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                cfg.write(f)

        cfg.read(CONFIG_FILE, encoding="utf-8")
        a = cfg["API"] if cfg.has_section("API") else {}
        self.dws_url.setText(a.get("dws_url", ""))
        self.dws_token.setText(a.get("dws_token", ""))
        self.jms_token.setText(a.get("jms_token", ""))
        p = cfg["PATH"] if cfg.has_section("PATH") else {}
        self.path.setText(p.get("path", ""))
        f = cfg["FILE"] if cfg.has_section("FILE") else {}
        self.name_dws.setText(f.get("name_dws", "DWS9-11.xlsx"))
        self.name_auto.setText(f.get("name_auto", "DWSXAUTOPDA.xlsx"))
        self.name_dwspda.setText(f.get("name_dwspda", "DWSPDA.xlsx"))
        t = cfg["TIME"] if cfg.has_section("TIME") else {}
        self.delay.setCurrentText(t.get("run_minute", "5"))

    def set_status(self, text):
        self.status.setText(text)

    # --- rest business logic stays same ---
    def toggle_start(self):
        if not self.scheduler_running:
            self.start()
            self.btn_start.setText("Stop")
        else:
            self.stop()
            self.btn_start.setText("Start")

    def toggle_run(self):
        if not self.job_running:
            self.stop_requested = False
            self.run_now()
            self.btn_run.setText("Stop")
        else:
            self.stop_requested = True
            self.set_status("Status: Stopping...")
            self.btn_run.setText("Run Now")

    def get_time_range(self):
        start_str = f"{self.start_date.date().toString('yyyy-MM-dd')} {self.start_hour.currentText()}"
        end_str = f"{self.end_date.date().toString('yyyy-MM-dd')} {self.end_hour.currentText()}"
        start = datetime.strptime(start_str, "%Y-%m-%d %H:%M")
        end = datetime.strptime(end_str, "%Y-%m-%d %H:%M")
        if start > end:
            start, end = end, start
        return start.strftime("%Y-%m-%d %H:%M:%S"), end.strftime("%Y-%m-%d %H:%M:%S")

    def sleep_with_stop(self, seconds):
        for _ in range(seconds):
            if self.stop_requested:
                return False
            time.sleep(1)
        return True

    def run_dws(self):
        self.log("Start DWS request...", "DWS9-11")
        start, end = self.get_time_range()
        try:
            url = self.dws_url.text().strip()
            token = self.dws_token.text().strip()
            data = {"startTime": start, "endTime": end, "barcodeList": [], "businessTimeType": "", "dbCode": "2", "curPage": 1, "pageSize": 20}
            headers = {"Content-Type": "application/json;charset=UTF-8", "token": token, "Origin": "http://10.30.32.10", "Referer": "http://10.30.32.10/", "User-Agent": "Mozilla/5.0"}
            res = requests.post(url, json=data, headers=headers, timeout=60)
            self.log(f"STATUS: {res.status_code}", "DWS9-11")
            self.log(f"SIZE: {len(res.content)}", "DWS9-11")
            if res.status_code == 200 and len(res.content) > 100:
                os.makedirs(self.path.text(), exist_ok=True)
                with open(os.path.join(self.path.text(), self.name_dws.text()), "wb") as f:
                    f.write(res.content)
                self.log(f"Downloaded: {self.name_dws.text()}")
            else:
                self.log(f"Download failed: {res.text[:200]}")
                self.set_status("Status: Error")
        except Exception as e:
            self.log(f"ERROR: {e}")
            self.set_status("Status: Error")
        finally:
            self.log("DWS finished", "SYS")

    def _export_jms(self, base, headers, start, end, scanType, filename):
        self.log(f"Generate {filename}", "JMS")
        run_time = datetime.now()
        target_hour = run_time.replace(minute=0, second=0, microsecond=0)
        next_hour = target_hour + timedelta(hours=1)
        requests.post(f"{base}/scanningContrast/asyncDownExcel", json={"startTimeStr": start, "endTimeStr": end, "scanNetworkCode": "999004", "scanType": scanType, "excelType": "downExcelAll"}, headers=headers)
        if not self.sleep_with_stop(60):
            return
        for _ in range(12):
            if self.stop_requested:
                raise Exception("Stopped by user")
            if not self.sleep_with_stop(10):
                return
            res = requests.post(f"{base}/downLoadCenter/downLoadInfoList", json={"current": 1, "size": 20}, headers=headers).json()
            records = res.get("data", {}).get("records", [])
            records.sort(key=lambda x: x.get("downTime", ""), reverse=True)
            for r in records:
                try:
                    down_time = datetime.strptime(r.get("downTime"), "%Y-%m-%d %H:%M:%S")
                except Exception:
                    continue
                if scanType in r.get("queryJson", "") and r.get("finishOrNot") == "1" and r.get("downUrl") and target_hour <= down_time < next_hour and down_time >= run_time - timedelta(minutes=7):
                    sign = requests.post(f"{base}/downLoadCenter/getDownloadSignedUrl", json=r, headers=headers).json()
                    url = sign.get("data")
                    if not url:
                        continue
                    file = requests.get(url if isinstance(url, str) else url[0])
                    if len(file.content) < 100_000:
                        self.log("File not ready, retry...", "JMS")
                        continue
                    os.makedirs(self.path.text(), exist_ok=True)
                    with open(os.path.join(self.path.text(), filename), "wb") as f:
                        f.write(file.content)
                    self.log(f"Downloaded {filename}", "JMS")
                    return
        raise Exception(f"{filename} fail")

    def run_jms_auto(self):
        base = "https://jmsgw.jtexpress.co.th/operatingplatform"
        headers = {"Content-Type": "application/json;charset=UTF-8", "authtoken": self.jms_token.text().strip()}
        start, end = self.get_time_range()
        self._export_jms(base, headers, start, end, "建包扫描", self.name_auto.text())

    def run_jms_pda(self):
        base = "https://jmsgw.jtexpress.co.th/operatingplatform"
        headers = {"Content-Type": "application/json;charset=UTF-8", "authtoken": self.jms_token.text().strip()}
        start, end = self.get_time_range()
        self._export_jms(base, headers, start, end, "卸车扫描", self.name_dwspda.text())

    def run_all(self):
        self.job_running = True
        self.save_config()
        self.set_ui(False)
        self.set_status("Status: Running")
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
                self.set_status("Status: Success")
        except Exception as e:
            self.log(f"ERROR: {e}", "SYS")
            self.set_status("Status: Error")
        finally:
            self.job_running = False
            if not self.scheduler_running:
                self.set_ui(True)
            self.btn_run.setText("Run Now")
            if self.stop_requested:
                self.set_status("Status: Stopped")

    def validate(self):
        return all([self.dws_url.text(), self.dws_token.text(), self.jms_token.text(), self.path.text()])

    def validate_time(self):
        start = datetime.strptime(f"{self.start_date.date().toString('yyyy-MM-dd')} {self.start_hour.currentText()}", "%Y-%m-%d %H:%M")
        end = datetime.strptime(f"{self.end_date.date().toString('yyyy-MM-dd')} {self.end_hour.currentText()}", "%Y-%m-%d %H:%M")
        return start <= end

    def run_now(self):
        if not self.validate():
            self.log("Missing config", "SYS")
            return
        if not self.validate_time():
            self.log("Invalid time range (Start > End)", "SYS")
            return
        if self.job_running:
            return
        threading.Thread(target=self.run_all, daemon=True).start()

    def scheduler(self):
        while self.scheduler_running:
            now = datetime.now()
            run_minute = int(self.delay.currentText())
            target = now.replace(minute=run_minute, second=0, microsecond=0)
            if now >= target:
                target += timedelta(hours=1)
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

    def start(self):
        if self.scheduler_running:
            return
        self.scheduler_running = True
        self.set_ui(False)
        threading.Thread(target=self.scheduler, daemon=True).start()
        self.log("Auto scheduler started")

    def stop(self):
        self.scheduler_running = False
        self.stop_requested = True
        self.set_status("Status: Stopped")
        self.set_ui(True)
        self.btn_start.setText("Start")
        self.btn_run.setText("Run Now")

    def set_ui(self, enable):
        for w in [self.dws_url, self.dws_token, self.jms_token, self.path, self.start_date, self.end_date, self.start_hour, self.end_hour, self.delay, self.name_dws, self.name_auto, self.name_dwspda, self.btn_browse]:
            w.setEnabled(enable)
        self.btn_run.setEnabled(True)
        self.btn_start.setEnabled(True)


if __name__ == "__main__":
    app = QApplication([])
    window = App()
    window.show()
    app.exec()
