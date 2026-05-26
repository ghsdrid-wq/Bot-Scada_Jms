import sys
import threading
import time
from datetime import datetime, timedelta
import requests
import os
import configparser
from PySide6.QtCore import QDate, QTimer
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
    QTextEdit,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

CONFIG_FILE = "config.ini"


class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bot Export Scada&Jms")
        self.setFixedSize(650, 450)
        self.scheduler_running = False
        self.job_running = False
        self.stop_requested = False

        self.create_ui()
        self.load_config()

    def save_config(self):
        cfg = configparser.ConfigParser()

        cfg["API"] = {
            "dws_url": self.dws_url.text(),
            "dws_token": self.dws_token.text(),
            "jms_token": self.jms_token.text(),
        }

        cfg["PATH"] = {
            "path": self.path.text(),
        }

        cfg["FILE"] = {
            "name_dws": self.name_dws.text() or "DWS9-11.xlsx",
            "name_auto": self.name_auto.text() or "DWSXAUTOPDA.xlsx",
            "name_dwspda": self.name_dwspda.text() or "DWSPDA.xlsx",
        }

        cfg["TIME"] = {
            "run_minute": self.delay.currentText() or "5"
        }

        with open(CONFIG_FILE, "w") as f:
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
            }
            cfg["TIME"] = {"run_minute": "5"}
            with open(CONFIG_FILE, "w") as f:
                cfg.write(f)

        cfg.read(CONFIG_FILE)

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
        idx = self.delay.findText(t.get("run_minute", "5"))
        if idx >= 0:
            self.delay.setCurrentIndex(idx)

    def create_ui(self):
        tabs = QTabWidget()
        self.setCentralWidget(tabs)

        self.tab_home = QWidget()
        self.tab_setting = QWidget()

        tabs.addTab(self.tab_home, "Home")
        tabs.addTab(self.tab_setting, "Setting")

        self.build_home()
        self.build_setting()

    def build_home(self):
        root = QVBoxLayout(self.tab_home)
        self.status = QLabel("Status: Idle")
        root.addWidget(self.status)

        form_wrap = QWidget()
        form = QGridLayout(form_wrap)
        root.addWidget(form_wrap)

        hours = [f"{i:02d}:00" for i in range(24)]
        minutes = [str(i) for i in range(60)]

        form.addWidget(QLabel("Run minute"), 0, 0)
        self.delay = QComboBox()
        self.delay.addItems(minutes)
        self.delay.setCurrentText("5")
        self.delay.setFixedWidth(60)
        form.addWidget(self.delay, 0, 1)

        form.addWidget(QLabel("Start"), 1, 0)
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDisplayFormat("yyyy-MM-dd")
        self.start_date.setDate(QDate.currentDate())
        form.addWidget(self.start_date, 1, 1)

        self.start_hour = QComboBox()
        self.start_hour.addItems(hours)
        self.start_hour.setCurrentText("13:00")
        form.addWidget(self.start_hour, 1, 2)

        form.addWidget(QLabel("→"), 1, 3)
        form.addWidget(QLabel("End"), 1, 4)

        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDisplayFormat("yyyy-MM-dd")
        self.end_date.setDate(QDate.currentDate().addDays(1))
        form.addWidget(self.end_date, 1, 5)

        self.end_hour = QComboBox()
        self.end_hour.addItems(hours)
        self.end_hour.setCurrentText("23:00")
        form.addWidget(self.end_hour, 1, 6)

        self.btn_start = QPushButton("Start")
        self.btn_start.clicked.connect(self.toggle_start)
        form.addWidget(self.btn_start, 1, 7)

        self.btn_run = QPushButton("Run Now")
        self.btn_run.clicked.connect(self.toggle_run)
        form.addWidget(self.btn_run, 1, 8)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        root.addWidget(self.log_text)

    def build_setting(self):
        root = QVBoxLayout(self.tab_setting)

        dws_group = QGroupBox("DWS9-11")
        dws_form = QFormLayout(dws_group)
        self.dws_url = QLineEdit()
        self.dws_token = QLineEdit()
        dws_form.addRow("API URL", self.dws_url)
        dws_form.addRow("Token", self.dws_token)
        root.addWidget(dws_group)

        jms_group = QGroupBox("JMS")
        jms_form = QFormLayout(jms_group)
        self.jms_token = QLineEdit()
        jms_form.addRow("AuthToken", self.jms_token)
        root.addWidget(jms_group)

        path_group = QGroupBox("Output")
        path_layout = QHBoxLayout(path_group)
        path_layout.addWidget(QLabel("Save Path"))
        self.path = QLineEdit()
        path_layout.addWidget(self.path)
        self.btn_browse = QPushButton("Browse")
        self.btn_browse.clicked.connect(self.browse)
        path_layout.addWidget(self.btn_browse)
        root.addWidget(path_group)

        output_group = QGroupBox("PATH")
        output_form = QFormLayout(output_group)
        self.name_dws = QLineEdit()
        self.name_auto = QLineEdit()
        self.name_dwspda = QLineEdit()
        output_form.addRow("DWS File", self.name_dws)
        output_form.addRow("AUTOPDA File", self.name_auto)
        output_form.addRow("DWSPDA File", self.name_dwspda)
        root.addWidget(output_group)

    def toggle_start(self):
        if not self.scheduler_running:
            self.start()
            self.btn_start.setText("Stop")
        else:
            self.stop()
            self.btn_start.setText("Start")

    def toggle_run(self):
        if self.job_running:
            self.stop_requested = True
            self.status.setText("Status: Stopping...")
            self.btn_run.setText("Run Now")
            return

        self.stop_requested = False
        threading.Thread(target=self.execute_pipeline, daemon=True).start()
        self.btn_run.setText("Stop")

    def log(self, msg, tag="SYS"):
        now = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"{now} [{tag}] {msg}")

    def browse(self):
        p = QFileDialog.getExistingDirectory(self, "Select Folder")
        if p:
            self.path.setText(p)

    def run_dws(self):
        self.log("Start DWS request...", "DWS9-11")
        start, end = self.get_time_range()
        try:
            url = self.dws_url.text().strip()
            token = self.dws_token.text().strip()
            data = {"startTime": start, "endTime": end, "barcodeList": [], "businessTimeType": "", "dbCode": "2", "curPage": 1, "pageSize": 20}
            headers = {"Content-Type": "application/json;charset=UTF-8", "token": token, "Origin": "http://10.30.32.10", "Referer": "http://10.30.32.10/", "User-Agent": "Mozilla/5.0"}
            res = requests.post(url, json=data, headers=headers, timeout=(5, 10))
            self.log(f"STATUS: {res.status_code}", "DWS9-11")
            self.log(f"SIZE: {len(res.content)}", "DWS9-11")
            if res.status_code == 200 and len(res.content) > 100:
                os.makedirs(self.path.text(), exist_ok=True)
                path = os.path.join(self.path.text(), self.name_dws.text())
                with open(path, "wb") as f:
                    f.write(res.content)
                self.log(f"Downloaded: {self.name_dws.text()}")
            else:
                self.log(f"Download failed: {res.text[:200]}")
                self.status.setText("Error")
        except Exception as e:
            self.log(f"ERROR: {e}")
            self.status.setText("Error")
        finally:
            self.log("DWS finished", "SYS")

    def run_jms_auto(self):
        base = "https://jmsgw.jtexpress.co.th/operatingplatform"
        token = self.jms_token.text().strip()
        headers = {"Content-Type": "application/json;charset=UTF-8", "authtoken": token}
        start, end = self.get_time_range()
        self._export_jms(base, headers, start, end, "建包扫描", self.name_auto.text())

    def run_jms_pda(self):
        base = "https://jmsgw.jtexpress.co.th/operatingplatform"
        token = self.jms_token.text().strip()
        headers = {"Content-Type": "application/json;charset=UTF-8", "authtoken": token}
        start, end = self.get_time_range()
        self._export_jms(base, headers, start, end, "卸车扫描", self.name_dwspda.text())

    def _export_jms(self, base, headers, start, end, scanType, filename):
        self.log(f"Generate {filename}", "JMS")
        run_time = datetime.now()
        target_hour = run_time.replace(minute=0, second=0, microsecond=0)
        next_hour = target_hour + timedelta(hours=1)
        res = requests.post(f"{base}/scanningContrast/asyncDownExcel", json={"startTimeStr": start, "endTimeStr": end, "scanNetworkCode": "999004", "scanType": scanType, "excelType": "downExcelAll"}, headers=headers, timeout=(5, 10))
        if "login" in res.text.lower():
            self.log("JMS Token expired (generate)", "ERROR")
            raise Exception("Token expired")
        if not self.sleep_with_stop(50):
            return
        for _ in range(10):
            if self.stop_requested:
                self.log("Stopped by user")
                return
            if not self.sleep_with_stop(10):
                return
            res = requests.post(f"{base}/downLoadCenter/downLoadInfoList", json={"current": 1, "size": 20}, headers=headers, timeout=(5, 10))
            if "login" in res.text.lower():
                self.log("JMS Token expired (list)", "ERROR")
                raise Exception("Token expired")
            res = res.json()
            records = res.get("data", {}).get("records", [])
            records.sort(key=lambda x: x.get("downTime", ""), reverse=True)
            for r in records:
                try:
                    down_time = datetime.strptime(r.get("downTime"), "%Y-%m-%d %H:%M:%S")
                except:
                    continue
                self.log(f"Check: {r.get('downTime')} | finish={r.get('finishOrNot')}", "DEBUG")
                if (scanType in r.get("queryJson", "") and r.get("finishOrNot") == "1" and r.get("downUrl") and target_hour <= down_time < next_hour and down_time >= run_time - timedelta(minutes=7)):
                    self.log(f"Matched file: {down_time}", "DEBUG")
                    sign_res = requests.post(f"{base}/downLoadCenter/getDownloadSignedUrl", json=r, headers=headers, timeout=(5, 10))
                    if "login" in sign_res.text.lower():
                        self.log("JMS Token expired (signedUrl)", "ERROR")
                        raise Exception("Token expired")
                    sign = sign_res.json()
                    url = sign.get("data")
                    if not url:
                        continue
                    file = requests.get(url if isinstance(url, str) else url[0], timeout=(5, 10))
                    content_type = file.headers.get("Content-Type", "").lower()
                    if "html" in content_type or "login" in file.text.lower():
                        self.log("JMS Token expired (download)", "ERROR")
                        raise Exception("Token expired")
                    if len(file.content) < 100_000:
                        self.log("File not ready, retry...", "JMS")
                        continue
                    os.makedirs(self.path.text(), exist_ok=True)
                    path = os.path.join(self.path.text(), filename)
                    with open(path, "wb") as f:
                        f.write(file.content)
                    self.log(f"Downloaded {filename}", "JMS")
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
        start_str = f"{self.start_date.date().toString('yyyy-MM-dd')} {self.start_hour.currentText()}"
        end_str = f"{self.end_date.date().toString('yyyy-MM-dd')} {self.end_hour.currentText()}"
        start = datetime.strptime(start_str, "%Y-%m-%d %H:%M")
        end = datetime.strptime(end_str, "%Y-%m-%d %H:%M")
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
        run_minute = int(self.delay.currentText())
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
                self.log("Run job", "SYS")
                threading.Thread(target=self.execute_pipeline, daemon=True).start()
            else:
                self.log("Skip: previous job still running", "SYS")
            self.next_run += timedelta(hours=1)

    def execute_pipeline(self):
        if self.job_running:
            self.log("Skip: job running", "SYS")
            return
        self.job_running = True
        self.save_config()
        self.set_ui(False)
        self.status.setText("Status: Running")
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
            self.status.setText("Status: Success")
        except Exception as e:
            self.log(f"ERROR: {e}", "SYS")
            self.status.setText("Status: Error")
        finally:
            self.job_running = False
            self.stop_requested = False
            if self.scheduler_running and hasattr(self, "next_run") and self.next_run:
                self.log("Finished", "SYS")
                self.log(f"Next run at {self.next_run.strftime('%H:%M:%S')}", "SYS")
            QTimer.singleShot(0, lambda: self.btn_run.setText("Run Now"))
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
        self.btn_start.setText("Stop")

    def stop(self):
        self.scheduler_running = False
        self.stop_requested = True
        self.status.setText("Status: Stopped")
        self.set_ui(True)
        self.btn_start.setText("Start")
        self.btn_run.setText("Run Now")

    def set_ui(self, enable):
        for w in [self.dws_url, self.dws_token, self.jms_token, self.path, self.start_date, self.end_date, self.start_hour, self.end_hour, self.delay, self.name_dws, self.name_auto, self.name_dwspda, self.btn_browse]:
            w.setEnabled(enable)
        self.btn_run.setEnabled(True)
        self.btn_start.setEnabled(True)

    def validate(self):
        return bool(self.dws_url.text() and self.dws_token.text() and self.jms_token.text() and self.path.text())

    def validate_time(self):
        start = datetime.strptime(f"{self.start_date.date().toString('yyyy-MM-dd')} {self.start_hour.currentText()}", "%Y-%m-%d %H:%M")
        end = datetime.strptime(f"{self.end_date.date().toString('yyyy-MM-dd')} {self.end_hour.currentText()}", "%Y-%m-%d %H:%M")
        return start <= end


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = App()
    w.show()
    sys.exit(app.exec())
