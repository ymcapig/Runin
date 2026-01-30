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
from PyQt5.QtCore import QThread, pyqtSignal, Qt
# --------------------

class RunInWorker(QThread):
    sig_log = pyqtSignal(str)
    sig_finished = pyqtSignal(bool, str) # True=PASS, False=FAIL

    def __init__(self, logic_func):
        super().__init__()
        self.logic_func = logic_func

    def run(self):
        try:
            # 執行主流程
            self.logic_func()
            # 若無錯誤跑完，視為 PASS
            self.sig_finished.emit(True, "All Tests Passed")
        except Exception as e:
            # 只要有 Exception (包含 exec_cmd_wait 拋出的)，就視為 FAIL 並停止
            err_msg = str(e)
            self.sig_log.emit(f"!!! STOPPED: {err_msg} !!!")
            self.sig_finished.emit(False, err_msg)

class BaseRunInApp(QMainWindow):
    sig_update_ui_log = pyqtSignal(str)
    sig_update_status = pyqtSignal(str)

    def __init__(self, title="ACER Run-In Test"):
        super().__init__()
        self.setWindowTitle(title)
        self.resize(800, 600)
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        # --- 變數初始化 ---
        self.current_proc = None 
        self.stop_flag = False
        self.is_rebooting = False
        
        if getattr(sys, 'frozen', False):
            # 打包後：抓 .exe 的位置
            self.base_dir = os.path.dirname(sys.executable)
        else:
            # 開發時：抓 .py 的位置
            self.base_dir = os.path.dirname(os.path.abspath(__file__))
        # --- 路徑與資料夾設定 ---
        self.state_file = os.path.join(self.base_dir, "runin_state.json")
        self.log_dir = os.path.join(self.base_dir, "log")
        
        # Result 資料夾設定
        self.result_dir = os.path.join(self.base_dir, "Result")
        
        # 確保資料夾存在
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)
        if not os.path.exists(self.result_dir):
            os.makedirs(self.result_dir)
            
        self.current_log_file = os.path.join(self.log_dir, "Runin_Debug.log")

        # UI 初始化
        self.sig_update_ui_log.connect(self.append_log_text)
        # 連接狀態訊號到更新函式
        self.sig_update_status.connect(self.update_status_label)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        
        # 建立一個醒目的狀態標籤 (Status Label)
        self.lbl_status = QLabel("READY")
        self.lbl_status.setFixedHeight(60) # 設定高度
        self.lbl_status.setAlignment(Qt.AlignCenter) # 文字置中
        # 設定樣式: 深灰底、黃字、大字體、粗體
        self.lbl_status.setStyleSheet("""
            background-color: #333333; 
            color: #FFD700; 
            font-size: 24px; 
            font-weight: bold; 
            border: 2px solid #555;
            border-radius: 5px;
        """)
        # Log 區域
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setStyleSheet("background: black; color: #00FF00; font-family: Consolas; font-size: 15pt;")
        
        
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
        main_layout.addWidget(self.lbl_status)
        main_layout.addWidget(self.txt_log)
        main_layout.addLayout(btn_layout)

        # 啟動檢查
        self.check_previous_log()
        self.disable_runonce = "--factory" in sys.argv
        if self.disable_runonce:
            print("[System] Factory mode RunOnce Registry disabled (Managed by external launcher).")

        self.last_saved_state = {}
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
        # 如果狀態檔存在，代表這是同一次測試的延續，不應該切分 Log
        if os.path.exists(self.state_file):
            self.log(">>> Detected Resume State. Continuing with existing log. <<<")
            return
        # 如果沒有狀態檔，但 Log 卻存在，才視為上次的殘留檔進行封存
        if os.path.exists(self.current_log_file):
            try:
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                with open(self.current_log_file, "a", encoding="utf-8") as f:
                    f.write(f"[{timestamp}] !!! DETECTED NEW STARTUP WHILE LOG EXISTS (PREVIOUS CRASH?) !!!\n")
            except:
                pass
            self.log("Found old log from previous run. Archiving...")
            self.archive_log(prefix="Runin_Debug_Crash_")

    def start_test(self, is_resume=False):
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
        tool_path = os.path.join(self.base_dir, "RI", "DiagECtool.exe")
        try:
            self.log("Resetting Battery Mode...")
            subprocess.run(f"{tool_path} battery --mode auto", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)      
            self.log("Resetting Fan Mode...")
            subprocess.run(f"{tool_path} fan --mode auto", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            self.log(f"Error resetting hardware: {e}")

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
        
        popen_kwargs = {
            'shell': True
        }
        if capture_log:
            self.log(f"CMD > {cmd}")
            # 只有要抓 Log 時才使用 PIPE
            popen_kwargs['stdout'] = subprocess.PIPE
            popen_kwargs['stderr'] = subprocess.STDOUT
        else:
            self.log(f"CMD > {cmd} (Log Capture Disabled - Independent Console)")
            # 針對不抓 Log 的指令 (通常是 legacy tool 如 FDPCMD)
            # 強制開啟一個全新的 Console 視窗，避開 PyInstaller 無視窗環境的限制
            popen_kwargs['creationflags'] = subprocess.CREATE_NEW_CONSOLE

        self.current_proc = subprocess.Popen(cmd, **popen_kwargs)
        ret = -1
        try:
            if capture_log:
                # 有抓 Log 才需要讀取迴圈
                while True:
                    if self.stop_flag:
                        self.current_proc.terminate()
                        break
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
            
        except subprocess.TimeoutExpired:
            self.log(f"Command Timeout ({timeout}s): {cmd}")
            if self.current_proc:
                self.current_proc.kill()
            raise Exception(f"Command Timeout: {cmd}")           
        except Exception as e:
            self.log(f"Command Exception: {e}")
            if self.current_proc:
                self.current_proc.kill()
            raise e
        finally:
            self.current_proc = None
            
        self.check_stop()

        if ret != 0:
            raise Exception(f"Command Failed (Ret: {ret}): {cmd}")
        
        self.log("CMD < PASS")

    def run_external_tool_standalone(self, cmd_str):
        """
        專門用來執行像 FDPCMD 這種會搶 Console 的工具。
        它會彈出一個獨立的黑視窗執行，執行完後關閉。
        回傳: return code (0=Pass, 非0=Fail)
        """
        import subprocess
        
        self.log(f"Running standalone command: {cmd_str}")
        
        try:
            # CREATE_NEW_CONSOLE (0x00000010) 讓它擁有獨立視窗
            # 這樣 FDPCMD 就會像您手動執行一樣快樂
            creation_flags = subprocess.CREATE_NEW_CONSOLE
            
            # 使用 call 等待它執行完畢
            # 這裡我們不接管 stdout/stderr，直接讓它顯示在新視窗
            ret_code = subprocess.call(cmd_str, creationflags=creation_flags, shell=True)
            
            return ret_code
            
        except Exception as e:
            self.log(f"Error running standalone tool: {e}")
            return -1
        
    def save_state(self, block, step, cycle=1, status="IDLE"):
        state = {"block": block, "step": step, "cycle": cycle, "status": status}
        with open(self.state_file, "w") as f: json.dump(state, f)
        self.last_saved_state = state

    def load_state(self):
        if os.path.exists(self.state_file):
            with open(self.state_file, "r") as f: return json.load(f)
        return None

    def clear_state(self):
        if os.path.exists(self.state_file): os.remove(self.state_file)

    def trigger_reboot(self):
        try:
            self.check_stop()
            self.is_rebooting = True

            if not self.disable_runonce:
                key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\RunOnce", 0, winreg.KEY_WRITE)
                cmd = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'
                winreg.SetValueEx(key, "ODM_RunIn", 0, winreg.REG_SZ, cmd)
                winreg.CloseKey(key)
                self.log("RunOnce Registry Key Set (Standalone Mode).")
            else:
                self.log("Skipping RunOnce Registry (Managed Mode).")
            
            self.log("Reboot triggered. Shutting down...")            
            subprocess.run("shutdown /r /t 0 /f", shell=True)
            while True: time.sleep(1)
        except Exception as e:
            self.is_rebooting = False
            self.log(f"Reboot Failed: {e}")
            raise e
        
    def set_run_once_startup(self):
        """將目前的程式註冊到 Windows RunOnce，下次開機自動執行一次"""
        if self.disable_runonce:
            self.log("Skipping RunOnce for S4/ColdBoot (Managed Mode).")
            return
        try:
            # 1. 取得目前執行檔的路徑
            if getattr(sys, 'frozen', False):
                # PyInstaller 打包後：抓 .exe 完整路徑
                exe_path = sys.executable
                self.log(f"set_run_once_startup {exe_path}")
            else:
                exe_path = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'
                self.log("Dev mode: Skipping RunOnce registration (Run packaged EXE to test reboot)")
                return

            # 2. 開啟登錄檔路徑
            key = winreg.CreateKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\RunOnce",
                0, 
                winreg.KEY_SET_VALUE
            )
            
            # 3. 寫入數值 (Name: ODM_RunIn, Value: "C:\Path\To\ODM_RunIn.exe")
            # 建議加上引號，避免路徑有空白出錯
            winreg.SetValueEx(key, "ODM_RunIn", 0, winreg.REG_SZ, f'"{exe_path}"')
            winreg.CloseKey(key)
            
            self.log(f"RunOnce Registered: {exe_path}")
            
        except Exception as e:
            self.log(f"Failed to set startup: {e}")

    def on_finished(self, passed, msg=""):
        # 決定最終結果：必須是 passed 為 True 且沒有被 STOP
        final_result = passed and not self.stop_flag

        if not self.stop_flag:
            tool_path = os.path.join(self.base_dir, "RI", "DiagECtool.exe") 
            try:
                self.exec_cmd_wait(f"{tool_path} battery --mode auto", capture_log=True)
                self.exec_cmd_wait(f"{tool_path} fan --mode auto", capture_log=True)
            except Exception as e:
                self.log(f"Cleanup warning: {e}")

        if self.stop_flag:
            self.log("=== TEST STOPPED BY USER (Please Restart Application) ===")
            self.set_status("TEST STOPPED")
            self.btn_start.setText("STOPPED")
            self.btn_start.setEnabled(False)
            self.btn_stop.setEnabled(False)
            #self.clear_state()
        else:
            self.btn_start.setEnabled(False)
            self.btn_stop.setEnabled(False)
            
            if final_result:
                self.btn_start.setText("PASS")
                if self.last_saved_state:
                    self.save_state(
                        self.last_saved_state.get("block", "1"),
                        self.last_saved_state.get("step", 0),
                        self.last_saved_state.get("cycle", 1),
                        status="FINISHED_PASS"
                    )
            else:
                self.btn_start.setText("FAIL")
                if self.last_saved_state:
                    self.save_state(
                        self.last_saved_state.get("block", "1"),
                        self.last_saved_state.get("step", 0),
                        self.last_saved_state.get("cycle", 1),
                        status="FINISHED_FAIL"
                    )
        
        # 產生結果檔
        self.generate_result_file(final_result)     
        self.archive_log()

    def closeEvent(self, event):
        if self.is_rebooting:
            event.accept()
            return
        
        if hasattr(self, 'fan_thread') and self.fan_thread is not None:
            if self.fan_thread.isRunning():
                self.log("Stopping fan monitor...")
                self.fan_thread.stop() # FanMonitorThread 定義的 stop()
                # Fan Thread 很快，通常不用太久的 wait，但為了保險可以 wait
                self.fan_thread.wait(1000)

        # --- 安全停止 Worker ---
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.log("Close event detected. Stopping worker thread...")
            self.stop_flag = True  # 通知 Worker 停止           
            # 嘗試等待一下讓 Worker 收屍 (選用，避免卡住 UI 太久)
            max_wait_time = 60000 
            start_time = time.time()
            while self.worker.isRunning():
                if self.worker.wait(100): 
                    break # 如果 Worker 結束了就跳出
                QApplication.processEvents()
                if (time.time() - start_time) * 1000 > max_wait_time:
                    self.log("Worker thread timeout (30s), forcing close.")
                    break
        # -----------------------------        
        try:
            print("Resetting hardware to Auto mode...")
            # 假設 DiagECtool 在 .\RI\DiagECtool.exe
            tool_exe = os.path.join(os.getcwd(), "RI", "DiagECtool.exe")            
            subprocess.run(f"{tool_exe} battery --mode auto", shell=True, timeout=2, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            print("Resetting battery to Auto mode...")
            subprocess.run(f"{tool_exe} fan --mode auto", shell=True, timeout=2, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            print("Resetting fan to Auto mode...")        
        except Exception as e:
            print(f"Hardware reset error: {e}")
        # ----------------------------- 
        if self.current_proc and self.current_proc.poll() is None:
             subprocess.run(f"taskkill /F /T /PID {self.current_proc.pid}", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        # 判斷依據：如果 START 按鈕是 Disabled (且不是 STOPPED 狀態)，代表正在跑
        if not self.btn_start.isEnabled() and self.btn_start.text() == "RUNNING...":
            self.generate_result_file(False)

        if os.path.exists(self.current_log_file):
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
    
    def set_status(self, text):
        # 透過 emit 發送訊號，確保在主執行緒更新 UI (Thread-Safe)
        self.sig_update_status.emit(text)

    def update_status_label(self, text):
        self.lbl_status.setText(text)
