import sys
import os
import time
import subprocess
import json
import configparser
import winreg
from datetime import datetime

# --- PyQt5 修改區 ---
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QTextEdit, QProgressBar, QLabel, QPushButton, QMessageBox)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
# --------------------

class RunInWorker(QThread):
    # PyQt5 使用 pyqtSignal 定義信號
    sig_log = pyqtSignal(str)
    sig_finished = pyqtSignal(bool) # True=PASS, False=FAIL

    def __init__(self, logic_func):
        super().__init__()
        self.logic_func = logic_func

    def run(self):
        try:
            # 執行主流程
            self.logic_func()
            # 若無錯誤跑完，視為 PASS
            self.sig_finished.emit(True)
        except Exception as e:
            # 只要有 Exception (包含 exec_cmd_wait 拋出的)，就視為 FAIL 並停止
            self.sig_log.emit(f"!!! STOPPED: {str(e)} !!!")
            self.sig_finished.emit(False)

class BaseRunInApp(QMainWindow):
    def __init__(self, title="ODM Run-In Framework (PyQt5)"):
        super().__init__()
        self.setWindowTitle(title)
        self.resize(600, 500)
        self.state_file = "C:\\runin_state.json"
        
        # UI Setup
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setStyleSheet("background: black; color: #00FF00; font-family: Consolas;")
        
        self.btn_start = QPushButton("START TEST")
        self.btn_start.clicked.connect(lambda: self.start_test(False))
        
        layout.addWidget(QLabel(title))
        layout.addWidget(self.txt_log)
        layout.addWidget(self.btn_start)

        # Config
        self.config = configparser.ConfigParser()
        if os.path.exists("config.ini"):
            # --- 修正點：加入 encoding="utf-8" ---
            self.config.read("config.ini", encoding="utf-8")
        else:
            self.log("WARNING: config.ini not found!")

        # Resume Check
        if os.path.exists(self.state_file):
            self.start_test(is_resume=True)

    def start_test(self, is_resume=False):
        self.btn_start.setEnabled(False)
        self.btn_start.setText("RUNNING...")
        if is_resume: self.log(">>> RESUMED FROM REBOOT <<<")
        
        self.worker = RunInWorker(self.user_test_sequence)
        self.worker.sig_log.connect(self.log)
        self.worker.sig_finished.connect(self.on_finished)
        self.worker.start()

    def user_test_sequence(self):
        raise NotImplementedError

    def log(self, msg):
            # 寫入文字
            self.txt_log.append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
            # 1. 強制捲動到底部 (避免捲軸卡住導致渲染錯誤)
            self.txt_log.ensureCursorVisible()
            # 2. 強制處理 UI 事件 (讓介面有機會重繪)
            QApplication.processEvents()

    def exec_cmd_wait(self, cmd, timeout=None):
        """
        執行指令並等待結束。
        關鍵邏輯：若 Return Code != 0，直接 Raise Exception 讓 Worker 停止。
        """
        self.log(f"CMD > {cmd}")
        # 使用 subprocess.run 等待執行結束
        # shell=True 允許執行 BAT 檔與系統指令
        res = subprocess.run(cmd, shell=True, timeout=timeout)
        
        if res.returncode != 0:
            # 這行 Exception 會被 Worker 捕捉，進而觸發停止
            raise Exception(f"Command Failed (Ret: {res.returncode}): {cmd}")
        
        self.log("CMD < PASS")

    # --- 狀態管理 (Reboot用) ---
    def save_state(self, block, step, cycle=1):
        state = {"block": block, "step": step, "cycle": cycle}
        with open(self.state_file, "w") as f: json.dump(state, f)

    def load_state(self):
        if os.path.exists(self.state_file):
            with open(self.state_file, "r") as f: return json.load(f)
        return None

    def clear_state(self):
        if os.path.exists(self.state_file): os.remove(self.state_file)

    def trigger_reboot(self):
        try:
            # 設定 RunOnce 讓程式在下次登入時自動執行
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\RunOnce", 0, winreg.KEY_WRITE)
            cmd = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'
            winreg.SetValueEx(key, "ODM_RunIn", 0, winreg.REG_SZ, cmd)
            winreg.CloseKey(key)
            
            self.log("Reboot triggered. Shutting down...")
            subprocess.run("shutdown /r /t 0 /f", shell=True)
            
            # 強制卡住主執行緒，等待系統重啟
            while True: time.sleep(1)
        except Exception as e:
            self.log(f"Reboot Failed: {e}")
            raise e

    def on_finished(self, passed):
        self.btn_start.setEnabled(True)
        self.btn_start.setText("FINISHED")
        if passed:
            self.log("=== TEST FINISHED: PASS ===")
            self.clear_state()
        else:
            self.log("=== TEST STOPPED: FAIL ===")
            # 失敗不清除狀態檔，方便 Debug