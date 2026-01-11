import sys
import os
import time
import subprocess
import json
import configparser
import winreg
from datetime import datetime

# --- PyQt5 修改區 ---
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QTextEdit, QLabel, QPushButton)
from PyQt5.QtCore import QThread, pyqtSignal
# --------------------

class RunInWorker(QThread):
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
    sig_update_ui_log = pyqtSignal(str)

    def __init__(self, title="ODM Run-In Framework (PyQt5)"):
        super().__init__()
        self.setWindowTitle(title)
        self.resize(800, 600)
        
        # --- 變數初始化 ---
        self.current_proc = None 
        self.stop_flag = False
        
        # --- 路徑與資料夾設定 ---
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.state_file = os.path.join(self.base_dir, "runin_state.json")
        self.log_dir = os.path.join(self.base_dir, "log")
        
        # [新增] Result 資料夾設定
        self.result_dir = os.path.join(self.base_dir, "Result")
        
        # 確保資料夾存在
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)
        if not os.path.exists(self.result_dir):
            os.makedirs(self.result_dir)
            
        self.current_log_file = os.path.join(self.log_dir, "Runin_Debug.log")

        # UI 初始化
        self.sig_update_ui_log.connect(self.append_log_text)
        
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        
        # Log 區域
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setStyleSheet("background: black; color: #00FF00; font-family: Consolas;")
        
        # 按鈕區域
        btn_layout = QHBoxLayout()
        self.btn_start = QPushButton("START TEST")
        self.btn_start.setFixedHeight(40)
        self.btn_start.clicked.connect(lambda: self.start_test(False))
        
        self.btn_stop = QPushButton("STOP")
        self.btn_stop.setFixedHeight(40)
        self.btn_stop.setEnabled(False)
        self.btn_stop.setStyleSheet("background-color: #D32F2F; color: white; font-weight: bold;")
        self.btn_stop.clicked.connect(self.stop_test)
        
        btn_layout.addWidget(self.btn_start)
        btn_layout.addWidget(self.btn_stop)
        
        main_layout.addWidget(QLabel(title))
        main_layout.addWidget(self.txt_log)
        main_layout.addLayout(btn_layout)

        # 啟動檢查
        self.check_previous_log()

        # Config 讀取
        self.config = configparser.ConfigParser()
        if os.path.exists("config.ini"):
            self.config.read("config.ini", encoding="utf-8")
        else:
            self.log("WARNING: config.ini not found!")

        # 斷點續傳檢查
        if os.path.exists(self.state_file):
            self.start_test(is_resume=True)

    def check_previous_log(self):
        if os.path.exists(self.current_log_file):
            try:
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                with open(self.current_log_file, "a", encoding="utf-8") as f:
                    f.write(f"[{timestamp}] !!! DETECTED NEW STARTUP WHILE LOG EXISTS (PREVIOUS CRASH?) !!!\n")
            except:
                pass
            self.log("Found unfinished log. Archiving as Crash...")
            self.archive_log(prefix="Runin_Debug_Crash_")

    def start_test(self, is_resume=False):
        # [新增] 每次開始前，清除上次的結果檔
        self.cleanup_results()

        self.stop_flag = False
        self.btn_start.setEnabled(False)
        self.btn_start.setText("RUNNING...")
        self.btn_stop.setEnabled(True)
        
        if is_resume: self.log(">>> RESUMED FROM REBOOT <<<")
        
        self.worker = RunInWorker(self.user_test_sequence)
        self.worker.sig_log.connect(self.log)
        self.worker.sig_finished.connect(self.on_finished)
        self.worker.start()

    def stop_test(self):
        self.log("!!! USER PRESSED STOP BUTTON !!!")
        self.stop_flag = True
        
        self.btn_stop.setEnabled(False)
        self.btn_start.setEnabled(False)
        
        if self.current_proc:
            try:
                if self.current_proc.poll() is None:
                    pid = self.current_proc.pid
                    self.log(f"Killing external process (PID={pid})...")
                    subprocess.run(f"taskkill /F /T /PID {pid}", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception as e:
                self.log(f"Failed to kill process: {e}")

    def check_stop(self):
        if self.stop_flag:
            raise Exception("User Manually Stopped the Test.")

    def user_test_sequence(self):
        raise NotImplementedError

    def log(self, msg):
        self.sig_update_ui_log.emit(msg)
        print(msg)
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with open(self.current_log_file, "a", encoding="utf-8") as f:
                f.write(f"[{timestamp}] {msg}\n")
        except Exception as e:
            print(f"Write log failed: {e}")

    def append_log_text(self, msg):
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.txt_log.append(f"[{timestamp}] {msg}")
        self.txt_log.ensureCursorVisible()

    # 修改 exec_cmd_wait 函式，增加 capture_log 參數
    def exec_cmd_wait(self, cmd, timeout=None, capture_log=True):
        self.check_stop()
        
        # 根據是否抓 Log 來決定提示訊息
        if capture_log:
            self.log(f"CMD > {cmd}")
            # 使用 PIPE 抓取輸出 (原本的邏輯)
            self.current_proc = subprocess.Popen(
                cmd, 
                shell=True, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.STDOUT
            )
        else:
            self.log(f"CMD > {cmd} (Log Capture Disabled)")
            # [修正點] 不使用 PIPE，直接繼承 Console，避免 FDPCMD 崩潰
            self.current_proc = subprocess.Popen(cmd, shell=True)

        try:
            if capture_log:
                # 有抓 Log 才需要讀取迴圈
                while True:
                    output_bytes = self.current_proc.stdout.readline()
                    if not output_bytes and self.current_proc.poll() is not None:
                        break
                    if output_bytes:
                        try:
                            line = output_bytes.decode('cp950', errors='replace').strip()
                            if line:
                                self.log(f"[SYS] {line}")
                        except Exception as e:
                            print(f"Decode error: {e}")
            
            # 等待結束
            ret = self.current_proc.wait(timeout=timeout)
            
        except Exception as e:
            raise e
        finally:
            self.current_proc = None
            
        self.check_stop()

        if ret != 0:
            raise Exception(f"Command Failed (Ret: {ret}): {cmd}")
        
        self.log("CMD < PASS")

    def save_state(self, block, step, cycle=1, status="IDLE"):
        state = {"block": block, "step": step, "cycle": cycle, "status": status}
        with open(self.state_file, "w") as f: json.dump(state, f)

    def load_state(self):
        if os.path.exists(self.state_file):
            with open(self.state_file, "r") as f: return json.load(f)
        return None

    def clear_state(self):
        if os.path.exists(self.state_file): os.remove(self.state_file)

    def trigger_reboot(self):
        try:
            self.check_stop()
            
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\RunOnce", 0, winreg.KEY_WRITE)
            cmd = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'
            winreg.SetValueEx(key, "ODM_RunIn", 0, winreg.REG_SZ, cmd)
            winreg.CloseKey(key)
            
            self.log("Reboot triggered. Shutting down...")
            self.archive_log(prefix="Runin_Debug_Reboot_")
            
            subprocess.run("shutdown /r /t 0 /f", shell=True)
            while True: time.sleep(1)
        except Exception as e:
            self.log(f"Reboot Failed: {e}")
            raise e

    def on_finished(self, passed):
        # 決定最終結果：必須是 passed 為 True 且沒有被 STOP
        final_result = passed and not self.stop_flag

        if self.stop_flag:
            self.log("=== TEST STOPPED BY USER (Please Restart Application) ===")
            self.btn_start.setText("STOPPED")
            self.btn_start.setEnabled(False)
            self.btn_stop.setEnabled(False)
            self.clear_state()
        else:
            self.btn_start.setEnabled(True)
            self.btn_start.setText("FINISHED")
            self.btn_stop.setEnabled(False)
            
            if final_result:
                self.log("=== TEST FINISHED: PASS ===")
                self.clear_state()
            else:
                self.log("=== TEST STOPPED: FAIL ===")
        
        # [新增] 產生結果檔
        self.generate_result_file(final_result)
        
        self.archive_log()

    def closeEvent(self, event):
        if os.path.exists(self.current_log_file):
            self.log("User closed the application (Manual Abort).")
            if self.current_proc and self.current_proc.poll() is None:
                 subprocess.run(f"taskkill /F /T /PID {self.current_proc.pid}", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            # [新增] 如果還在跑的時候關閉，視為 FAIL
            # 判斷依據：如果 START 按鈕是 Disabled (且不是 STOPPED 狀態)，代表正在跑
            if not self.btn_start.isEnabled() and self.btn_start.text() == "RUNNING...":
                self.generate_result_file(False)

            self.archive_log(prefix="Runin_Debug_UserAbort_")
        event.accept()

    def archive_log(self, prefix="Runin_Debug_"):
        if os.path.exists(self.current_log_file):
            try:
                timestamp_str = datetime.now().strftime('%Y%m%d_%H%M%S')
                new_name = f"{prefix}{timestamp_str}.log"
                new_path = os.path.join(self.log_dir, new_name)
                if os.path.exists(new_path): os.remove(new_path)
                os.rename(self.current_log_file, new_path)
                self.sig_update_ui_log.emit(f"Log archived to: log\\{new_name}")
            except Exception as e:
                self.sig_update_ui_log.emit(f"Failed to archive log: {e}")

    # --- [新增] 結果檔管理功能 ---
    def cleanup_results(self):
        """清除上次的 PASS/FAIL 檔案"""
        for name in ["PASS", "FAIL"]:
            path = os.path.join(self.result_dir, name)
            if os.path.exists(path):
                try:
                    os.remove(path)
                    self.log(f"Cleared previous result: {name}")
                except Exception as e:
                    self.log(f"Failed to clear {name}: {e}")

    def generate_result_file(self, is_pass):
        """產生 PASS 或 FAIL 檔案"""
        filename = "PASS" if is_pass else "FAIL"
        filepath = os.path.join(self.result_dir, filename)
        
        try:
            with open(filepath, "w") as f:
                # 寫入時間戳記，方便追溯
                f.write(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                if not is_pass and self.stop_flag:
                    f.write("Reason: User Manually Stopped\n")
            
            self.log(f">>> Result Generated: {filename} <<<")
        except Exception as e:
            self.log(f"Failed to generate result file: {e}")
