import sys
import os
import time
import psutil 
import csv
import subprocess
import threading
import ctypes
import shutil  # 用於複製檔案
from datetime import datetime, timedelta
from PyQt5.QtWidgets import QApplication, QMessageBox
from core import BaseRunInApp
from PyQt5.QtCore import Qt, QThread, QLockFile, QDir, QTimer, pyqtSignal

# ==========================================
# Helper: EC io txrx class
# ==========================================
class DirectEC:
    def __init__(self, dll_folder):
        # 設定 inpoutx64.dll 路徑
        dll_path = os.path.join(dll_folder, "inpoutx64.dll")
        if not os.path.exists(dll_path):
            print(f"[DirectEC] Error: DLL not found at {dll_path}")
            self.dll = None
            return

        try:
            self.dll = ctypes.WinDLL(dll_path)
            self.cmd_port = 0x6C
            self.dat_port = 0x68
            self.initialized = True
            print("[DirectEC] DLL Loaded successfully.")
        except Exception as e:
            print(f"[DirectEC] Failed to load DLL: {e}")
            self.dll = None
            self.initialized = False

    def wait_ibf_clear(self, timeout_s=0.5):
        """等待 Input Buffer Full 清除"""
        t0 = time.perf_counter()
        while time.perf_counter() - t0 < timeout_s:
            if (self.dll.Inp32(self.cmd_port) & 0x02) == 0:
                return True
            time.sleep(0.001)
        return False

    def wait_obf_set(self, timeout_s=0.5):
        """等待 Output Buffer Full 設定 (有資料可讀)"""
        t0 = time.perf_counter()
        while time.perf_counter() - t0 < timeout_s:
            if (self.dll.Inp32(self.cmd_port) & 0x01) != 0:
                return True
            time.sleep(0.001)
        return False

    def txrx(self, cmd, data_payload, expect_len, wait_s=0.05):
        if not self.initialized: return None

        try:
            # 1. Write Command
            if not self.wait_ibf_clear(): return None
            self.dll.Out32(self.cmd_port, cmd)
            time.sleep(0.05) # 模擬 ecio 的 command delay

            # 2. Write Payload
            for d in data_payload:
                if not self.wait_ibf_clear(): return None
                self.dll.Out32(self.dat_port, d)
                time.sleep(0.005) # 模擬 ecio 的 data delay

            # 3. Read Response
            resp = []
            for _ in range(expect_len):
                if self.wait_obf_set(timeout_s=wait_s):
                    val = self.dll.Inp32(self.dat_port) & 0xFF
                    resp.append(val)
                else:
                    break # Timeout or no more data
            
            return resp
        except Exception as e:
            print(f"[DirectEC] txrx error: {e}")
            return None
    def get_fan_rpm(self, fan_id):
        # CMD=0x20, SubCmd=0x05, Payload=[0x05, FanID]
        # 回傳: [LowByte, HighByte]
        if not self.initialized: return 0
        
        CMD = 0x20
        payload = [0x05, fan_id]
        
        # 嘗試 3 次 (如原始碼)
        for _ in range(3):
            resp = self.txrx(CMD, payload, expect_len=2, wait_s=0.05)
            if resp and len(resp) == 2:
                rpm = resp[0] | (resp[1] << 8)
                return rpm
        return 0

    def get_ts2_temp(self):
        # CMD=0x28, TS2_SubCmd=0x05
        # 回傳: [Temp]
        if not self.initialized: return 0
        
        CMD = 0x28
        payload = [0x05] # 0x05 對應 thermaltest.py 中的 "ts2"
        
        for _ in range(3):
            resp = self.txrx(CMD, payload, expect_len=1, wait_s=0.05)
            if resp and len(resp) == 1:
                return resp[0]
        return 0
    
    def get_charging_current(self):
        if not self.initialized: return 0
        
        CMD = 0x31
        payload = [0x05]
        
        for _ in range(3):
            resp = self.txrx(CMD, payload, expect_len=2, wait_s=0.1)
            if resp and len(resp) == 2:
                # 1. 先組合成 Unsigned 16-bit (0 ~ 65535)
                raw_val = resp[0] | (resp[1] << 8)               
                # 2. Signed 16-bit (Two's Complement)
                if raw_val >= 32768:
                    raw_val -= 65536
                    
                return raw_val
        return 0
        
# ==========================================
# Helper: 風扇監控執行緒 (背景執行)
# ==========================================
class FanMonitorThread(QThread):
    update_signal = pyqtSignal(int, int, int)
    def __init__(self, csv_path, interval=1):
        super().__init__()
        self.csv_path = csv_path
        self.interval = interval
        self.running = True

        if getattr(sys, 'frozen', False):
            self.base_dir = os.path.dirname(sys.executable)
        else:
            self.base_dir = os.path.dirname(os.path.abspath(__file__))
        ri_folder = os.path.join(self.base_dir, "RI")
        self.ec = DirectEC(ri_folder)
    def run(self):
        # 確保目錄存在
        os.makedirs(os.path.dirname(self.csv_path), exist_ok=True)
        
        # 寫入 CSV Header
        with open(self.csv_path, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(["Timestamp", "Fan1_RPM", "Fan2_RPM", "TS2"])

        while self.running:
            start_time = time.time()
            try:
                now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                rpm1 = self.ec.get_fan_rpm(1)
                rpm2 = self.ec.get_fan_rpm(2)
                ts2 = self.ec.get_ts2_temp()
                with open(self.csv_path, 'a', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow([now, rpm1, rpm2, ts2])

                self.update_signal.emit(rpm1, rpm2, ts2)
            except Exception as e:
                pass 

            #間隔休息
            elapsed = time.time() - start_time
            sleep_time = max(0.1, self.interval - elapsed)
            time.sleep(sleep_time)

    def get_rpm(self, fan_id):
        try:
            # 假設 DiagECtool 在 .\RI 目錄下
            cmd = f".\\RI\\DiagECtool.exe fan --get-rpm --id {fan_id}"
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            output = subprocess.check_output(cmd, startupinfo=si, shell=True).decode().strip()
            
            # 解析 "Fan RPM: 3000" 格式
            if ":" in output:
                val_str = output.split(":")[-1].strip()
                return int(val_str) if val_str.isdigit() else 0
            
            # 若直接回傳數字
            return int(output) if output.isdigit() else 0
        except:
            return 0
        
    def get_ts2(self):
        try:
            # 假設 DiagECtool 在 .\RI 目錄下
            cmd = f".\\RI\\DiagECtool.exe temp --sensor ts2"
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            output = subprocess.check_output(cmd, startupinfo=si, shell=True).decode().strip()
            
            # 解析 "Fan RPM: 3000" 格式
            if ":" in output:
                val_str = output.split(":")[-1].strip()
                return int(val_str) if val_str.isdigit() else 0
            
            # 若直接回傳數字
            return int(output) if output.isdigit() else 0
        except:
            return 0

    def stop(self):
        self.running = False
        self.wait()

# ==========================================
# 主程式邏輯
# ==========================================
class ODM_RunIn_Project(BaseRunInApp):

    def __init__(self, title="ODM Run-In"):
        super().__init__(title=title)   
        # 設定一個 Timer，在介面顯示後 1 秒檢查是否要 Auto Run
        # 這樣可以確保 UI 已經完全 Load 好
        if getattr(sys, 'frozen', False):
            # 打包後：抓 .exe 的位置
            self.base_dir = os.path.dirname(sys.executable)
        else:
            # 開發時：抓 .py 的位置
            self.base_dir = os.path.dirname(os.path.abspath(__file__))
        QTimer.singleShot(1000, self.check_auto_run)
        ri_folder = os.path.join(self.base_dir, "RI")
        self.ec = DirectEC(ri_folder)

    # Auto Run 檢查邏輯
    def check_auto_run(self):
        try:
            # 讀取 Config [Global] AutoRun
            # 注意: 需確保 config.ini 有 [Global] Section
            if 'Global' in self.config and self.config['Global'].getboolean('AutoRun'):
                self.log("[AutoRun] Config detected. Starting test automatically...")              
                # 假設 BaseRunInApp 有一個 self.btn_start 按鈕
                if hasattr(self, 'btn_start'):
                    self.btn_start.click()
                else:
                    self.log("Error: Start button not found, cannot auto run.")
        except Exception as e:
            self.log(f"AutoRun Error: {e}")

    def user_test_sequence(self):
        state = self.load_state()
        current_block = "1"
        current_step = 0
        current_cycle = 1
        last_status = "IDLE"

        # 斷點續傳恢復邏輯
        if state:
            current_block = state.get("block", "1")
            current_step = state.get("step", 0)
            current_cycle = state.get("cycle", 1)
            last_status = state.get("status", "IDLE")

            self.log(f">>> RESUMING Block {current_block}, Step {current_step} <<<")
            if last_status == "FINISHED_PASS":
                self.log("Test Previously PASSED. Showing Result...")
                self.worker.sig_finished.emit(True, "Loaded from State History")
                return 
            
            if last_status == "FINISHED_FAIL":
                self.log("Test Previously FAILED. Showing Result...")
                self.worker.sig_finished.emit(False, "Loaded from State History")
                return 
            # 偵測非預期當機 (上次狀態仍為 RUNNING)
            if last_status == "REBOOTING":
                self.log(">>> System returned from Reboot Test. Resuming...")

            if last_status == "RUNNING":
                self.log(f"!!! DETECTED CRASH/HANG AT BLOCK {current_block} STEP {current_step} !!!")
                self.log("Detected Crash. Retrying this step (Production Mode)...")
                
                # 簡單的 Fail Log
                with open("RunIn_Crash.log", "a") as f:
                    f.write(f"Crash at Block {current_block} Step {current_step} - Retrying...\n")

                # 跳過該步驟，避免無限重啟
                # current_step += 1
                # self.save_state(current_block, current_step, current_cycle, status="IDLE")

        # ---------------------------------------------------
        # Block 1: Thermal
        # ---------------------------------------------------
        if current_block == "1":
            if self.config['Block1_Thermal'].getboolean('Enabled'):
                self.run_block_1(start_from_step=current_step, current_cycle=current_cycle)
                current_block = "2"
                current_step = 0
                self.save_state("2", 0, current_cycle, status="IDLE")
            else:
                self.log("Block 1 is Disabled. Skipping to Block 2...")
                current_block = "2"
                self.save_state("2", 0, current_cycle, status="IDLE")

        # ---------------------------------------------------
        # Block 2: Aging
        # ---------------------------------------------------
        if current_block == "2":
            if self.config['Block2_Aging'].getboolean('Enabled'):
                self.run_block_2(start_from_step=current_step, current_cycle=current_cycle)
                current_block = "3"
                current_step = 0
                self.save_state("3", 0, current_cycle, status="IDLE")
            else:
                self.log("Block 2 is Disabled. Skipping to Block 3...")
                current_block = "3"
                self.save_state("3", 0, current_cycle, status="IDLE")

        # ---------------------------------------------------
        # Block 3: Battery
        # ---------------------------------------------------
        if current_block == "3":
            if self.config['Block3_Battery'].getboolean('Enabled'):
                self.run_block_3(current_cycle=current_cycle)
            else:
                self.log("Block 3 is Disabled. Skipping...")

            try:
                total_cycles = int(self.config['Global']['Total_RunIn_Cycles'])
            except:
                total_cycles = 1 # 預設值
            
            if current_cycle < total_cycles:
                next_cycle = current_cycle + 1
                self.log(f"=== Cycle {current_cycle} Finished. Starting Cycle {next_cycle}/{total_cycles} ===")                
                # 設定狀態回到 Block 1, Step 0，並更新 Cycle數
                self.save_state("1", 0, next_cycle, status="IDLE")
                # 為了確保測試環境乾淨，通常建議重開機進入下一輪
                self.log("Rebooting system for the next cycle...")
                self.trigger_reboot()
                return # 結束本次執行，等待重開機
            
            self.clear_state()
            self.log("=== ALL BLOCKS FINISHED ===")
        
            # if 'Global' in self.config and self.config['Global'].getboolean('AutoClose'):
            #     self.log("[AutoClose] Closing application in 3 seconds...")
            #     # 為了讓使用者看到 PASS，延遲 3 秒再關閉
            #     import time
            #     time.sleep(3)
            #     QApplication.quit()

    # --- Helper: 計算 CSV 平均值 ---
    def analyze_fan_log_average(self, csv_path, col_idx, duration_sec=120):
        self.log(f"Analyzing {os.path.basename(csv_path)}...")
        if not os.path.exists(csv_path):
            self.log("Error: Log file not found.")
            return 0.0
            
        values = []
        cutoff_time = datetime.now() - timedelta(seconds=duration_sec)
        
        try:
            with open(csv_path, 'r') as f:
                reader = csv.reader(f)
                next(reader, None) # Skip Header
                for row in reader:
                    if len(row) <= col_idx: continue
                    try:
                        t_str = row[0]
                        # 假設 CSV 時間格式 YYYY-mm-dd HH:MM:SS
                        t_obj = datetime.strptime(t_str, '%Y-%m-%d %H:%M:%S')
                        if t_obj >= cutoff_time:
                            values.append(float(row[col_idx]))
                    except: continue
            
            if not values: 
                self.log("Warning: No valid data found in timeframe.")
                return 0.0
            
            avg = sum(values) / len(values)
            self.log(f"Average: {avg:.2f}")
            return avg
        except Exception as e:
            self.log(f"Analysis Error: {e}")
            return 0.0
        
    def archive_fan_log(self, src_path, prefix_name):
        if not os.path.exists(src_path):
            return None # 檔案不存在回傳 None
        try:
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            new_filename = f"{timestamp}_{prefix_name}.csv"
            dest_dir = os.path.dirname(src_path)
            dest_path = os.path.join(dest_dir, new_filename)
            
            # 使用 move 取代 copy2，這樣原檔就會被刪除 (移動)
            shutil.move(src_path, dest_path)
            
            self.log(f"Fan Log archived (moved): {new_filename}")
            return dest_path # [新增] 回傳新路徑供分析使用
            
        except Exception as e:
            self.log(f"Error archiving fan log: {e}")
            return None

    def analyze_ptat_log(self, csv_path, col_idx, duration_sec=120):
            """ 
            讀取 PTAT CSV: 
            - Date 在第 1 欄 (MM/DD/YYYY)
            - Time 在第 2 欄 (HH:MM:SS:fff) (毫秒用冒號分隔)
            """
            self.log(f"[PTAT Analysis] {os.path.basename(csv_path)}")
            if not os.path.exists(csv_path):
                self.log("Error: Log file not found.")
                return 0.0
                
            values = []
            cutoff_time = datetime.now() - timedelta(seconds=duration_sec)
            
            try:
                with open(csv_path, 'r', encoding='utf-8', errors='ignore') as f:
                    reader = csv.reader(f)
                    next(reader, None) # Skip Header
                    
                    for row in reader:
                        # PTAT 至少要有 3 欄 (Version, Date, Time) + 數據
                        if len(row) <= col_idx or len(row) < 3: continue
                        
                        try:
                            date_str = row[1] # e.g. 10/01/2026
                            time_str = row[2] # e.g. 08:53:58:985
                            
                            # 修正 PTAT 的毫秒格式: 08:53:58:985 -> 08:53:58.985
                            if time_str.count(':') == 3:
                                last_colon = time_str.rfind(':')
                                time_str = time_str[:last_colon] + '.' + time_str[last_colon+1:]
                            
                            full_time_str = f"{date_str} {time_str}"
                            # 解析合併後的時間
                            try:
                                # 嘗試使用 DD/MM/YYYY (適配您目前的 CSV)
                                t_obj = datetime.strptime(full_time_str, '%d/%m/%Y %H:%M:%S.%f')
                            except ValueError:
                                # 若失敗，嘗試原本的 MM/DD/YYYY (相容舊格式)
                                t_obj = datetime.strptime(full_time_str, '%m/%d/%Y %H:%M:%S.%f')                            
                            if t_obj >= cutoff_time:
                                values.append(float(row[col_idx]))
                        except ValueError:
                            continue
                
                if not values: 
                    self.log("Warning: No valid PTAT data found in timeframe.")
                    return 9999.0
                
                avg = sum(values) / len(values)
                return avg  # 這裡不印 log 避免洗版，由上層呼叫者印
                
            except Exception as e:
                self.log(f"PTAT Analysis Error: {e}")
                return 9999.0
    # --- Helper: 尋找最新 Log ---
    def find_latest_log(self, folder, prefix="PTATMonitor", extension=".csv"):
        try:
            if not os.path.exists(folder): return None
            files = [os.path.join(folder, f) for f in os.listdir(folder) 
                     if f.startswith(prefix) and f.endswith(extension)]
            if not files: return None
            return max(files, key=os.path.getmtime)
        except Exception as e:
            self.log(f"Error finding log: {e}")
            return None

    # --- Helper: 確保 Process 關閉 ---
    def ensure_process_killed(self, process_name):
        self.log(f"Stopping {process_name}...")
        subprocess.run(f"taskkill /F /IM {process_name}", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(1)
        # Retry check
        for _ in range(10):
            self.log(f"Stopping {process_name} retry...")
            is_running = False
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] and proc.info['name'].lower() == process_name.lower():
                    is_running = True
                    break
            if not is_running: return
            time.sleep(2)
            subprocess.run(f"taskkill /F /IM {process_name}", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    # --- Helper: PTAT 檢查 (依 Config 欄位) ---
    def check_ptat_metrics(self, csv_path, test_mode="Test1"):
        self.log(f"Verifying PTAT Metrics in {os.path.basename(csv_path)}...")
        errors = []
        detailed_data = []
        if not os.path.exists(csv_path): return ["PTAT Log not found"]

        # 1. 找出 Config 定義的 Keys
        ptat_keys = []
        for key, value in self.config['Block1_Thermal'].items():
            if key.lower().startswith('ptat_key_'):
                ptat_keys.append(value)
        
        if not ptat_keys:
            self.log("No PTAT_Key defined in Config. Skipping check.")
            return [], []

        # 2. 建立 Header Map
        header_map = {}
        with open(csv_path, 'r') as f:
            reader = csv.reader(f)
            headers = next(reader, None)
            if not headers: return ["PTAT CSV is empty"]
            header_map = {name.strip(): idx for idx, name in enumerate(headers)}

        # 3. 檢查每個 Key
        for target_col_name in ptat_keys:
            if target_col_name not in header_map:
                msg = f"Config Error: Column '{target_col_name}' not found in CSV"
                self.log(msg)
                errors.append(msg) 
                continue
                
            col_idx = header_map[target_col_name]
            avg_val = self.analyze_ptat_log(csv_path, col_idx, duration_sec=120)
            
            try:
                cfg_low_key = f"{test_mode}_Low"
                cfg_high_key = f"{test_mode}_High"
                
                limit_low = float(self.config[target_col_name][cfg_low_key])
                limit_high = float(self.config[target_col_name][cfg_high_key])
                item_result = "PASS"
                if avg_val < limit_low or avg_val > limit_high:
                    msg = f"{target_col_name} FAIL: {avg_val:.2f} (Spec: {limit_low}~{limit_high})"
                    self.log(msg)
                    errors.append(msg)
                    item_result = "FAIL"
                else:
                    self.log(f"PASS: {target_col_name} = {avg_val:.2f} (Spec: {limit_low}~{limit_high})")

                detailed_data.append({
                    "Item": target_col_name,
                    "Value": f"{avg_val:.2f}",
                    "Min": limit_low,
                    "Max": limit_high,
                    "Result": item_result
                })

            except KeyError:
                self.log(f"WARNING: Config key '{cfg_low_key}/{cfg_high_key}' missing for [{target_col_name}]")
            except ValueError:
                self.log(f"WARNING: Invalid value for [{target_col_name}]")
        
        return errors, detailed_data

    def get_ptat_avg_power_value(self, csv_path, test_mode="Test1"):
        self.log(f"Calculating PTAT Power Avg ({test_mode})...") # Log 可視需求開關        
        if not os.path.exists(csv_path): 
            return 0.0

        # 1. 找出 Config 定義的 Watt Key (只取第一個找到的)
        target_col_name = None
        for key, value in self.config['Block1_Thermal'].items():
            if key.lower() == 'ptat_watt_key': # 精確比對 key 名稱
                target_col_name = value
                break
        
        if not target_col_name:
            target_col_name = "Power-Package Power(Watts)"

        try:
            # 2. 讀取 Header 並找 Index
            col_idx = -1
            with open(csv_path, 'r', encoding='utf-8', errors='ignore') as f:
                reader = csv.reader(f)
                headers = next(reader, None)
                if headers:
                    # 去除空白並比對
                    clean_headers = [h.strip() for h in headers]
                    if target_col_name in clean_headers:
                        col_idx = clean_headers.index(target_col_name)
            
            if col_idx != -1:
                # 3. 呼叫底層函式計算 (重用 analyze_ptat_log)
                avg_val = self.analyze_ptat_log(csv_path, col_idx, duration_sec=120)
                # self.log(f"PTAT Power ({target_col_name}): {avg_val:.2f} W")
                return avg_val
            else:
                self.log(f"Warning: PTAT Watt Key '{target_col_name}' not found in CSV.")
                return 0.0

        except Exception as e:
            self.log(f"Error getting PTAT power: {e}")
            return 0.0
        
    def analyze_gpumon_log(self, csv_path, col_idx, duration_sec=120):
        if not os.path.exists(csv_path): return 0.0
        values = []
        cutoff_time = datetime.now() - timedelta(seconds=duration_sec)       
        try:
            with open(csv_path, 'r', encoding='utf-8', errors='ignore') as f:
                reader = csv.reader(f)
                next(reader, None) # Skip Header
                for row in reader:
                    # GPUMon CSV: Iteration(0), Date(1), Timestamp(2), Data(3...)
                    if len(row) <= col_idx or len(row) < 3: continue
                    try:
                        date_str = row[1].strip() # 2026/01/11
                        time_str = row[2].strip() # 22:34:39:692
                        
                        # 處理毫秒分隔符號
                        if time_str.count(':') == 3:
                            last_colon = time_str.rfind(':')
                            time_str = time_str[:last_colon] + '.' + time_str[last_colon+1:]
                        
                        full_time_str = f"{date_str} {time_str}"
                        # 格式: YYYY/MM/DD HH:MM:SS.f
                        t_obj = datetime.strptime(full_time_str, '%Y/%m/%d %H:%M:%S.%f')
                        
                        if t_obj >= cutoff_time:
                            values.append(float(row[col_idx]))
                    except ValueError: continue

            if not values: return 0.0
            return sum(values) / len(values)
        except Exception as e:
            self.log(f"GPUMon Analysis Error: {e}")
            return 0.0
    

    def check_gpumon_metrics(self, csv_path, test_mode="Test1"):
        self.log(f"Verifying GPUMon Metrics ({test_mode})...")
        errors = []       
        detailed_data = []
        # 1. 找出 Config Keys
        gpu_keys = []
        for key, value in self.config['Block1_Thermal'].items():
            if key.lower().startswith('gpumon_key_'):
                gpu_keys.append(value)
        
        if not gpu_keys: return [], []
        # 2. 建立 Header Map
        header_map = {}
        with open(csv_path, 'r') as f:
            reader = csv.reader(f)
            headers = next(reader, None)
            if not headers: return ["GPUMon CSV is empty"]
            header_map = {name.strip(): idx for idx, name in enumerate(headers)}
        # 3. 檢查數值
        for target_col in gpu_keys:
            if target_col not in header_map:
                errors.append(f"GPUMon Column '{target_col}' not found")
                continue           
            avg_val = self.analyze_gpumon_log(csv_path, header_map[target_col], duration_sec=120)
            
            try:
                cfg_low = f"{test_mode}_Low"
                cfg_high = f"{test_mode}_High"
                limit_low = float(self.config[target_col][cfg_low])
                limit_high = float(self.config[target_col][cfg_high])
                
                item_result = "PASS"
                if avg_val < limit_low or avg_val > limit_high:
                    msg = f"GPUMon {target_col} FAIL: {avg_val:.2f} (Spec: {limit_low}~{limit_high})"
                    self.log(msg)
                    errors.append(msg)
                    item_result = "FAIL"
                else:
                    self.log(f"GPUMon PASS: {target_col} = {avg_val:.2f}")
                
                detailed_data.append({
                    "Item": f"GPUMon_{target_col}",
                    "Value": f"{avg_val:.2f}",
                    "Min": limit_low,
                    "Max": limit_high,
                    "Result": item_result
                })
            except Exception as e:
                errors.append(f"GPUMon Config Error [{target_col}]: {e}")               
        return errors, detailed_data
    
    def get_gpumon_power_avg(self, csv_path, test_mode="Test1"):
        # self.log(f"Calculating GPUMon Power Avg ({test_mode})...")
        
        if not os.path.exists(csv_path):
            return 0.0
            
        target_col_name = None
        for key, value in self.config['Block1_Thermal'].items():
            if key.lower() == 'gpumon_watt_key':
                target_col_name = value
                break
        
        if not target_col_name:
            target_col_name = "1:TGP (W)" # 預設值

        try:
            col_idx = -1
            with open(csv_path, 'r', encoding='utf-8', errors='ignore') as f:
                reader = csv.reader(f)
                headers = next(reader, None)
                if headers:
                    clean_headers = [h.strip() for h in headers]
                    if target_col_name in clean_headers:
                        col_idx = clean_headers.index(target_col_name)

            if col_idx != -1:
                # 這裡假設 analyze_gpumon_log 參數與 analyze_ptat_log 類似
                # 如果 analyze_gpumon_log 邏輯不同，請對應調整
                avg_val = self.analyze_gpumon_log(csv_path, col_idx, duration_sec=120)                
                if test_mode == "Test1": return 0.0                 
                self.log(f"GPUMon Power ({target_col_name}): {avg_val:.2f} W")
                return avg_val
            else:
                return 0.0

        except Exception as e:
            self.log(f"Error getting GPUMon power: {e}")
            return 0.0
    # ==========================================
    # Helper: 取得風扇轉速 (單純讀取版)
    # ==========================================
    def get_fan_rpm(self, fan_id, tool_path):
        """
        執行指令取得 RPM，不包含 Retry 邏輯。
        回傳: int (若失敗或無法解析則回傳 0)
        """
        cmd = f'"{tool_path}" fan --get-rpm --id {fan_id}'
        try:
            # 執行指令並抓取輸出
            output = subprocess.check_output(cmd, shell=True).decode().strip()
            
            # 解析 "Fan RPM: 3000" 或 純數字
            if ":" in output:
                val_str = output.split(":")[-1].strip()
                return int(val_str) if val_str.isdigit() else 0
            
            return int(output) if output.isdigit() else 0
        except Exception:
            return 0
        
    def run_fan_curve_test(self):
            self.log("[Fan Speed Test] Starting...")
            
            log_dir = r"C:\Diag\Thermal"
            os.makedirs(log_dir, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            log_file = os.path.join(log_dir, f"{timestamp}_fan_rpm_test.log")   
            tool_path = os.path.join(self.base_dir, "RI", "DiagECtool.exe")  
            try:
                try:
                    fan_count = int(self.config['Block1_Thermal']['Test2_Fan_Count'])
                    fan_ids = [str(i) for i in range(1, fan_count + 1)]
                except ValueError:
                    raise Exception("Config Error: Test2_Fan_Count must be an integer")

                try:
                    sample_count = int(self.config['Block1_Thermal']['Test2_Sample_Count'])
                    target_duty = self.config['Block1_Thermal']['Test2_Duty']   
                    retry_limit = int(self.config['Block1_Thermal'].get('Fan_Retry_Count', '3'))
                except KeyError as e:
                    raise Exception(f"Config Error: Missing key {e}")                                
                # A. 切換 Mode Debug
                self.log("Set Fan Mode: DEBUG")
                self.exec_cmd_wait(f"{tool_path} fan --mode debug", capture_log=True)
                
                all_failures = []
                
                with open(log_file, "w") as f:
                    f.write(f"Timestamp: {timestamp}\n")
                    f.write(f"Fan Test Start (Count: {fan_count}, Target Duty: {target_duty}%)\n")
                    f.write(f"Retry Limit: {retry_limit}\n")
                    f.write("========================================\n")

                # B. 迴圈測試 Levels
                try:
                    self.log(f"--- Testing Duty {target_duty}% ---")
                    
                    # 2. 設定所有風扇的 Duty
                    for fan_id in fan_ids:
                        cmd_set = f"{tool_path} fan --set-duty {target_duty} --id {fan_id}"
                        self.log(f"Setting Fan {fan_id} Duty: {cmd_set}")
                        subprocess.run(cmd_set, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    
                    # 3. 等待穩定
                    self.log("Waiting 5 seconds for fan to stabilize...")
                    for _ in range(5):
                        time.sleep(1)
                        QApplication.processEvents()

                    # 4. 抓取 RPM 並判定
                    for fan_id in fan_ids:
                        # Key 變更為 Test2_FanX_Min/Max (移除 Level 字眼)
                        key_min = f"Test2_Fan{fan_id}_Min"
                        key_max = f"Test2_Fan{fan_id}_Max"
                        
                        try:
                            spec_min = int(self.config['Block1_Thermal'][key_min])
                            spec_max = int(self.config['Block1_Thermal'][key_max])
                        except KeyError:
                            self.log(f"ERROR: Config key '{key_min}' or '{key_max}' missing!")
                            all_failures.append(f"Config Missing for Fan {fan_id}")
                            continue

                        # 抓取 RPM (取樣平均)
                        rpms = []
                        for i in range(sample_count):
                            val = 0
                            # [新增] Retry 機制：判斷 0 或 異常大值
                            for attempt in range(retry_limit):
                                val = self.get_fan_rpm(fan_id, tool_path)                          
                                # 若數值合理，直接跳出重試迴圈
                                if 0 < val < 10000:
                                    break                              
                                # 若數值異常，等待 1 秒後重試
                                self.log(f"Debug: Fan {fan_id} read {val}, retrying {attempt+1}/{retry_limit}...") 
                                time.sleep(0.5)
                            rpms.append(val)
                        avg_rpm = sum(rpms) / len(rpms) if rpms else 0
                        
                        result_msg = f"Fan {fan_id} | Duty {target_duty} | Avg RPM: {avg_rpm:.1f} (Spec: {spec_min}-{spec_max})"
                        self.log(result_msg)
                        
                        with open(log_file, "a") as f:
                            f.write(f"{datetime.now()} | {result_msg}\n")
                            f.write(f"    Raw Data: {rpms}\n")

                        if not (spec_min <= avg_rpm <= spec_max):
                            fail_msg = f"FAIL: Fan {fan_id} Duty {target_duty} RPM {avg_rpm:.1f} Out of Spec"
                            all_failures.append(fail_msg)
                            with open(log_file, "a") as f: f.write(f"    [RESULT] FAIL\n")
                        else:
                            with open(log_file, "a") as f: f.write(f"    [RESULT] PASS\n")
                except Exception as e:
                    raise e
                # C. 最終判斷
                if all_failures:
                    raise Exception(" | ".join(all_failures))
            except Exception as e:
                raise e
                
            finally:
                self.log("Teardown: Set Fan Mode AUTO")
                subprocess.run(f"{tool_path} fan --mode auto", shell=True)
    # ==========================================
    # Thermal test 
    # ==========================================
    def run_stress_test_common(self, test_name, furmark_cmd, prime95_cmd):
        self.log(f"[{test_name}] Stress Test Starting...")
        log_dir = r"C:\Diag\Thermal"
        tool_path = os.path.join(self.base_dir, "RI", "DiagECtool.exe")  
        os.makedirs(log_dir, exist_ok=True)
        try:
            # 動態讀取對應的 Duration (例如 Test1_Duration 或 Test3_Duration)
            duration = int(self.config['Block1_Thermal'][f'{test_name}_Duration'])
        except KeyError:
            self.log(f"Config Error: {test_name}_Duration not found, default to 1200")
            duration = 1200

        # Log 檔名 (區分 Test1 / Test3)
        fan_log = os.path.join(log_dir, f"{test_name}_Fan.csv")
        fan_thread = None

        # 檢查 GPUMon 是否啟用
        gpumon_keys = [k for k in self.config['Block1_Thermal'] if k.lower().startswith('gpumon_key_')]
        is_gpumon_enabled = len(gpumon_keys) > 0
        if is_gpumon_enabled:
            self.log(f"[{test_name}] GPUMon Enabled.")
        else:
            self.log(f"[{test_name}] GPUMon Disabled.")
        try:
            # 0. set fan mode
            fan_mode = self.config['Block1_Thermal'].get('Fan_Mode', None)
            if fan_mode and str(fan_mode).strip():
                fan_mode = str(fan_mode).strip()
                self.log(f"Set Fan Mode: {fan_mode}")
                self.exec_cmd_wait(f"{tool_path} raw --cmd 0x20 --subcmd 0x01 --data 0x03", capture_log=True)
                self.exec_cmd_wait(f"{tool_path} raw --cmd 0x20 --subcmd 0x06 --data 0x0{fan_mode}", capture_log=True)
                for _ in range(10):
                    self.check_stop()
                    QApplication.processEvents()
                    time.sleep(1)
            # --- 階段 A: 啟動壓力工具 (Staggered Start) ---          
            # 1. 啟動 Furmark (傳入的指令)
            self.log(f"Starting Furmark: {furmark_cmd}")
            p_furmark = subprocess.Popen(furmark_cmd, shell=True)
            # 等待 10 秒 (讓 GPU warmup and stable)
            for _ in range(10):
                self.check_stop()
                QApplication.processEvents()
                time.sleep(1)

            # 2. 啟動 Prime95
            self.log(f"Starting Prime95: {prime95_cmd}")
            p_prime95 = subprocess.Popen(prime95_cmd, shell=True)
            
            # 3. 等待 40 秒 (PTAT 前置緩衝)
            self.log("Waiting 40s before starting PTAT...")
            for _ in range(40):
                self.check_stop()
                QApplication.processEvents()
                time.sleep(1)

            # 4. 啟動 PTAT
            ptat_dir = r"C:\Program Files\Intel Corporation\Intel(R)PTAT"
            if os.path.exists(ptat_dir):                    
                ptat_cmd = "PTAT.exe -start -w=cpu.json"
                self.log(f"Starting PTAT: {ptat_cmd}")
                p_ptat = subprocess.Popen(ptat_cmd, cwd=ptat_dir, shell=True)
            else:
                self.log("PTAT not installed (dir not found)")
                raise Exception("PTAT not installed")

            # 5. 等待 20 秒 (Fan/GPUMon 前置緩衝)
            self.log("Waiting 20s before starting Fan Monitor...")
            for _ in range(20):
                self.check_stop()
                QApplication.processEvents()
                time.sleep(1)

            # 6. 啟動 Fan Monitor
            fan_thread = FanMonitorThread(fan_log)
            fan_thread.start(QThread.TimeCriticalPriority)

            # 7. 啟動 GPUMon (若啟用)
            if is_gpumon_enabled:
                gpu_mon_dir = os.path.join(self.base_dir, "RI", "GPUMon")
                if not os.path.exists(gpu_mon_dir): os.makedirs(gpu_mon_dir)
                gpu_ppab_cmd = f"GPUMonCmd.exe -db:0"
                self.log(f"Disable PPAB: {gpu_ppab_cmd}")
                subprocess.Popen(gpu_ppab_cmd, cwd=gpu_mon_dir, shell=True)
                for _ in range(10):
                    self.check_stop()
                    QApplication.processEvents()
                    time.sleep(1)
                # 這裡 Log 檔名先用暫存的，最後再備份改名
                gpu_temp_log = "cpu_gpumon.csv" 
                gpu_cmd = f"GPUMonCmd.exe -custom:timestamp,temp,pwr,clk -wake -log:{gpu_temp_log}"
                self.log(f"Starting GPUMon: {gpu_cmd}")
                p_gpumon = subprocess.Popen(gpu_cmd, cwd=gpu_mon_dir, shell=True)

            # --- 階段 B: 正式燒機測試 ---
            self.log(f"Running Stress for {duration} seconds...")
            for i in range(duration):
                if i % 10 == 0: QApplication.processEvents()
                self.check_stop()
                time.sleep(1)

        except Exception as e:
            self.log(f"[{test_name}] Interrupted or Error: {e}")
            raise e

        finally:
            # --- 階段 C: Teardown (停止工具) ---
            self.log("Stopping Tools (Teardown)...")           
            # 1. 停 PTAT (並等待寫入)
            ptat_dir = r"C:\Program Files\Intel Corporation\Intel(R)PTAT"
            ptat_cmd = "PTAT.exe -stop"
            try:
                p_ptat = subprocess.Popen(ptat_cmd, cwd=ptat_dir, shell=True)
                p_ptat.wait(timeout=60)
            except:
                pass            
            self.log("Waiting 15s for PTAT logs...")
            # 關鍵: 保持 Prime95/Furmark 活著，等待 PTAT 寫完
            for i in range(15):
                if i % 10 == 0: QApplication.processEvents()
                time.sleep(1)           
            self.ensure_process_killed("PTAT.exe")           
            # 2. 停 GPUMon
            if is_gpumon_enabled:
                gpu_ppab_cmd = f"GPUMonCmd.exe -db:1"
                self.log(f"Enable PPAB: {gpu_ppab_cmd}")
                subprocess.Popen(gpu_ppab_cmd, cwd=gpu_mon_dir, shell=True)
                for _ in range(10):
                    self.check_stop()
                    QApplication.processEvents()
                    time.sleep(1)
                self.ensure_process_killed("GPUMonCmd.exe")           
            # 3. 停 Fan Monitor
            if fan_thread: fan_thread.stop()
            
            # 4. 停 Stress Tools (最後才殺)
            self.ensure_process_killed("prime95.exe")
            # 殺 Furmark (注意: FurMark GUI 與 CLI 可能名稱不同，通殺)
            self.ensure_process_killed("FurMark_GUI.exe")
            self.ensure_process_killed("furmark.exe")    
            # 釋放資源
            for i in range(10):
                if i % 10 == 0: QApplication.processEvents()
                time.sleep(1)  
            self.log("Teardown: Set Fan Mode AUTO")
            subprocess.run(f"{tool_path} fan --mode auto", shell=True)    

        # --- 階段 D: 結果驗證 ---
        self.log(f"=== Verifying {test_name} Results ===")
        all_failures = []
        
        summary_csv_data = []
        # 1. 備份與檢查 Fan Log
        archived_fan_log = None

        if test_name == "Test1":
            archived_fan_log = self.archive_fan_log(fan_log, f"CPU_only_Fan")
        else:
            archived_fan_log = self.archive_fan_log(fan_log, f"Dual_Fan")
        target_log_to_analyze = archived_fan_log if (archived_fan_log and os.path.exists(archived_fan_log)) else fan_log    
        try:
            # 讀取對應 Test1 或 Test3 的 Fan Spec
            spec_f1_min = int(self.config['Block1_Thermal'][f'{test_name}_Fan1_Min'])
            spec_f1_max = int(self.config['Block1_Thermal'][f'{test_name}_Fan1_Max'])       
            spec_f2_min = int(self.config['Block1_Thermal'][f'{test_name}_Fan2_Min'])
            spec_f2_max = int(self.config['Block1_Thermal'][f'{test_name}_Fan2_Max'])
            
            avg_fan1 = self.analyze_fan_log_average(target_log_to_analyze, 1) 
            avg_fan2 = self.analyze_fan_log_average(target_log_to_analyze, 2)
            
            fan_failed = False
            # --- 驗證 Fan 1 ---
            res_f1 = "PASS"
            if not (spec_f1_min <= avg_fan1 <= spec_f1_max):
                msg = f"Fan1 RPM FAIL: {avg_fan1} (Spec: {spec_f1_min}-{spec_f1_max})"
                self.log(msg)
                all_failures.append(msg)
                fan_failed = True
                res_f1 = "FAIL"

            summary_csv_data.append({
                "Item": "Fan1_RPM", "Value": avg_fan1, 
                "Min": spec_f1_min, "Max": spec_f1_max, "Result": res_f1
            })            
            # --- 驗證 Fan 2 ---
            res_f2 = "PASS"
            if not (spec_f2_min <= avg_fan2 <= spec_f2_max):
                msg = f"Fan2 RPM FAIL: {avg_fan2} (Spec: {spec_f2_min}-{spec_f2_max})"
                self.log(msg)
                all_failures.append(msg)
                fan_failed = True
                res_f2 = "FAIL"
            
            summary_csv_data.append({
                "Item": "Fan2_RPM", "Value": avg_fan2, 
                "Min": spec_f2_min, "Max": spec_f2_max, "Result": res_f2
            })
            # 若兩者都沒失敗才算 PASS
            if not fan_failed:
                self.log(f"Fan RPM PASS: Fan1={avg_fan1}, Fan2={avg_fan2}")    

        except KeyError as k:
            all_failures.append(f"Config Key Missing: {k}")       
        except Exception as e:
            all_failures.append(f"Fan Check Error: {e}")

        # 2. 備份與檢查 PTAT Log
        ptat_power_avg_val = 0
        gpumon_power_avg_val = 0
        user_home = os.path.expanduser("~")
        ptat_log_dir = os.path.join(user_home, "Documents", "iPTAT", "log")           
        ptat_log = self.find_latest_log(ptat_log_dir, prefix="PTATMonitor")
        if ptat_log:
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            # 檔名加入 test_name (例如 Test3_CPU_PTAT.csv)
            if test_name == "Test1":
                new_filename = f"{timestamp}_CPU_only_PTAT.csv"
            else:
                new_filename = f"{timestamp}_Dual_PTAT.csv"
            dest_path = os.path.join(log_dir, new_filename)
            try:
                shutil.copy2(ptat_log, dest_path)
                # 傳入 test_mode=test_name，這樣就會去讀 Test3_Low/High
                ptat_errors, ptat_data = self.check_ptat_metrics(dest_path, test_mode=test_name) 
                ptat_power_avg_val = self.get_ptat_avg_power_value(dest_path, test_mode=test_name)
                if ptat_errors:
                    all_failures.extend(ptat_errors)
                else:
                    self.log("PTAT Check PASS.")

                summary_csv_data.extend(ptat_data)
            except Exception as e:
                all_failures.append(f"PTAT Error: {e}")
        else:
            all_failures.append("PTAT Log missing")

        # 3. 備份與檢查 GPUMon Log
        if is_gpumon_enabled:
            gpu_mon_dir = os.path.join(self.base_dir, "RI", "GPUMon")
            src_gpu_log = os.path.join(gpu_mon_dir, "cpu_gpumon.csv")        
            if os.path.exists(src_gpu_log):
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                # 檔名加入 test_name
                dest_gpu_name = f"{timestamp}_{test_name}_GPUMon.csv"
                dest_gpu_path = os.path.join(log_dir, dest_gpu_name)              
                try:
                    shutil.copy2(src_gpu_log, dest_gpu_path)
                    # 傳入 test_mode=test_name
                    gpu_errors, gpu_data = self.check_gpumon_metrics(dest_gpu_path, test_mode=test_name)
                    gpumon_power_avg_val = self.get_gpumon_power_avg(dest_gpu_path, test_mode=test_name)
                    if gpu_errors:
                        all_failures.extend(gpu_errors)
                    else:
                        self.log("GPUMon Check PASS.")    

                    summary_csv_data.extend(gpu_data)                   
                except Exception as e:
                    all_failures.append(f"GPUMon Error: {e}")
            else:
                self.log("WARNING: GPUMon Log missing!")
                all_failures.append("GPUMon Log missing")
                
        # ==========================================
        # 4. 功耗檢查 (Power Check)
        # ==========================================
        # 4.1 CPU
        try:
            # 讀取 Config (若沒設則預設 0~9999)
            cpu_min = float(self.config['Block1_Thermal'].get(f'{test_name}_CPUPower_Min', 0))
            cpu_max = float(self.config['Block1_Thermal'].get(f'{test_name}_CPUPower_Max', 9999))
            
            res_cpu = "PASS"
            if not (cpu_min <= ptat_power_avg_val <= cpu_max):
                msg = f"CPU Power FAIL: {ptat_power_avg_val:.2f}W (Spec: {cpu_min}~{cpu_max})"
                self.log(msg)
                all_failures.append(msg)
                res_cpu = "FAIL"
            else:
                self.log(f"CPU Power PASS: {ptat_power_avg_val:.2f}W")
            
            summary_csv_data.append({
                "Item": "CPU_Power_Avg", "Value": ptat_power_avg_val, 
                "Min": cpu_min, "Max": cpu_max, "Result": res_cpu
            })
        except Exception as e:
            all_failures.append(f"CPU Power Check Error: {e}")

        # 4.2 GPU Power
        if is_gpumon_enabled and test_name == "Test3":
            try:
                gpu_min = float(self.config['Block1_Thermal'].get(f'{test_name}_GPUPower_Min', 0))
                gpu_max = float(self.config['Block1_Thermal'].get(f'{test_name}_GPUPower_Max', 9999))
                
                res_gpu = "PASS"
                if not (gpu_min <= gpumon_power_avg_val <= gpu_max):
                    msg = f"GPU Power FAIL: {gpumon_power_avg_val:.2f}W (Spec: {gpu_min}~{gpu_max})"
                    self.log(msg)
                    all_failures.append(msg)
                    res_gpu = "FAIL"
                else:
                    self.log(f"GPU Power PASS: {gpumon_power_avg_val:.2f}W")
                
                summary_csv_data.append({
                    "Item": "GPU_Power_Avg", "Value": gpumon_power_avg_val, 
                    "Min": gpu_min, "Max": gpu_max, "Result": res_gpu
                })
            except Exception as e:
                all_failures.append(f"GPU Power Check Error: {e}")

        total_pwr = ptat_power_avg_val + gpumon_power_avg_val
        self.log(f"[Power Check] CPU: {ptat_power_avg_val:.2f}W + GPU: {gpumon_power_avg_val:.2f}W = Total: {total_pwr:.2f}W")
        
        try:
            # 從 Config 讀取 Total Power Spec
            # 格式: Test1_TotalPower_Min / Max
            spec_min = float(self.config['Block1_Thermal'].get(f'{test_name}_TotalPower_Min', 0))
            spec_max = float(self.config['Block1_Thermal'].get(f'{test_name}_TotalPower_Max', 9999))
            
            res_pwr = "PASS"
            if not (spec_min <= total_pwr <= spec_max):
                msg = f"Total Power FAIL: {total_pwr:.2f}W (Spec: {spec_min}~{spec_max})"
                self.log(msg)
                all_failures.append(msg)
                res_pwr = "FAIL"
            else:
                self.log(f"Total Power PASS")
            
            summary_csv_data.append({
                "Item": "Total_Power", "Value": total_pwr, 
                "Min": spec_min, "Max": spec_max, "Result": res_pwr
            })    
        except ValueError:
            self.log("Warning: Invalid Total Power Spec in Config (Check format).")
        except Exception as e:
            all_failures.append(f"Total Power Check Error: {e}")
        # =========Add thermal summary file=================================    
        try:
            timestamp_str = datetime.now().strftime("%Y%m%d%H%M%S")
            # 檔名規則: PASS/FAIL_%timestamp%_cpu_only/dual.csv
            status_prefix = "FAIL" if all_failures else "PASS"
            mode_suffix = "cpu_only" if test_name == "Test1" else "dual"
            
            summary_filename = f"{status_prefix}_{timestamp_str}_{mode_suffix}.csv"
            summary_path = os.path.join(log_dir, summary_filename)
            
            with open(summary_path, 'w', newline='') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=["Item", "Value", "Min", "Max", "Result"])
                writer.writeheader()
                for row in summary_csv_data:
                    if isinstance(row["Value"], float):
                        row["Value"] = f"{row['Value']:.1f}"
                    writer.writerow(row)
            
            self.log(f"Verification Summary saved to: {summary_filename}")
            
        except Exception as e:
            self.log(f"Error generating summary CSV: {e}")
        # ==========================================
        # 最終判定
        if all_failures:
            error_summary = " | ".join(all_failures)
            raise Exception(f"{test_name} FAILED: {error_summary}")
        
        self.log(f"{test_name} ALL PASS.")
    
    # ==========================================
    # 電量檢查
    # ==========================================
    def check_battery_threshold(self):
        try:
            threshold = int(self.config['Block1_Thermal']['Start_Battery_Threshold'])
        except KeyError:
            threshold = 90 # 預設值

        self.log(f"Waiting for Battery > {threshold}%...")          
        
        while True:
            # 檢查是否有人按 STOP
            self.check_stop()
            QApplication.processEvents() 
            
            try:
                bat = psutil.sensors_battery()
                if bat:
                    current_pct = bat.percent
                    is_plugged = bat.power_plugged
                    status_str = "Charging" if is_plugged else "Discharging"                   
                    self.log(f"Battery Status: {current_pct}% ({status_str}) / Target: {threshold}%")

                    charging_current = self.ec.get_charging_current()
                    self.log(f"Battery charging current: {charging_current}mA")

                    if current_pct >= threshold:
                        self.log("Battery Threshold Reached!")
                        break
                    
                    if not is_plugged:
                        self.log("WARNING: AC Adapter not plugged in!")
                else:
                    self.log("No battery detected. Skipping check.")
                    break
            except Exception as e:
                self.log(f"Error reading battery: {e}")
            
            # 等待 5 秒再檢查
            time.sleep(5)
    # ==========================================
    # Block 1 實作 (Thermal)
    # ==========================================
    def run_block_1(self, start_from_step=0, current_cycle=1):
        self.log("--- Block 1: Thermal Tool ---")
        log_dir = r"C:\Diag\Thermal"
        os.makedirs(log_dir, exist_ok=True)
        
        # Step 0: 電量檢查
        tool_path = os.path.join(self.base_dir, "RI", "DiagECtool.exe")
        self.exec_cmd_wait(f"{tool_path} battery --mode auto", capture_log=True)
        for i in range(5):
            if i % 10 == 0: QApplication.processEvents()
            time.sleep(1)
        if start_from_step == 0:
            self.check_battery_threshold() 
            self.save_state("1", 1, current_cycle, status="IDLE")
        # ==========================================
        # [Test 1] Single Stress
        # ==========================================
        if start_from_step <= 1:
            self.log("[Test 1] Single Stress Start")
            self.set_status(f"Cycle {current_cycle} | Running Test 1: Single Stress")
            cmd_prime95 = r".\RI\prime95\prime95.exe -t -small"
            cmd_furmark = r"start .\RI\FurMark\FurMark_GUI.exe"
            try:
                # 呼叫共用函式: 傳入 "Test1"
                self.check_battery_threshold()
                self.run_stress_test_common("Test1", cmd_furmark, cmd_prime95)
                
                # 若沒拋出 Exception 代表 PASS
                self.save_state("1", 2, current_cycle, status="IDLE")
                if self.config['Block1_Thermal'].getboolean('Test1_Reboot'):
                    self.log("Rebooting as per config...")
                    self.trigger_reboot()
            except Exception as e:
                # 錯誤已在 run_stress_test_common Log 過，這裡直接往上拋即可
                raise e
        # Step 2: Fan Max Speed
        if start_from_step <= 2:
            self.log("Sleep 3min before Fan test...")   
            self.set_status(f"Cycle {current_cycle} | Running Test 2: Fan Speed Test")      
            for i in range(180):
                if i % 10 == 0: QApplication.processEvents()
                time.sleep(1)
            self.log("[Test 2] Fan Speed Test Start")   
            self.log(f"Battery percentage:{psutil.sensors_battery().percent}")     
            try:
                # 呼叫剛剛寫好的 Helper 函式
                self.run_fan_curve_test()
                self.log("Test 2 PASS.")               
                # 測試通過後，儲存狀態並準備重開機
                self.save_state("1", 3, current_cycle, status="IDLE")                
                # 檢查 Config 是否需要重開機 (需求說 PASS 後 Reboot 往 Test 3)
                if self.config['Block1_Thermal'].getboolean('Test2_Reboot'):
                    self.log("Rebooting for Test 3...")
                    self.trigger_reboot()
                else:
                    self.log("Skipping Reboot, moving to Test 3...")
            except Exception as e:
                self.log(f"Test 2 FAILED: {e}")
                # 這裡直接 raise，core.py 會捕捉並停止測試 (STOPPED: FAIL)
                raise e           
        # Step 3: Dual Stress
        if start_from_step <= 3:
            self.log("Sleep 3min before Dual burn test...")    
            self.set_status(f"Cycle {current_cycle} | Running Test 3: Dual Stress")     
            for i in range(180):
                if i % 10 == 0: QApplication.processEvents()
                time.sleep(1)
            self.check_battery_threshold() 
            self.log("[Test 3] Dual Stress Test")
            cmd_prime95 = r".\RI\prime95\prime95.exe -t -small"
            cmd_furmark = r".\RI\FurMark\furmark.exe --demo furmark-gl --benchmark --max-time 1390 --no-score-box"
            try:
                # 呼叫共用函式: 傳入 "Test3"
                self.run_stress_test_common("Test3", cmd_furmark, cmd_prime95)
                self.log("Block 1 PASS. Moving to Block 2...")               
               
                # 若為最後一個 Step，準備接續
                if self.config['Block1_Thermal'].getboolean('Test3_Reboot'):
                     self.log("Rebooting before Block 2...")
                     self.save_state("2", 0, current_cycle, status="IDLE")
                     self.trigger_reboot()
                else:
                     self.save_state("2", 0, current_cycle, status="IDLE")

            except Exception as e:
                raise e
            self.log("Block 1 PASS. Moving to Block 2...")
            
            if self.config['Block1_Thermal'].getboolean('Test3_Reboot'):
                 self.log("Rebooting before Block 2...")
                 self.save_state("2", 0, current_cycle, status="IDLE")
                 self.trigger_reboot()

    # ==========================================
    # Block 2 實作 (Aging)
    # ==========================================
    def run_block_2(self, start_from_step=0, current_cycle=1):
        self.log("--- Block 2: Aging Test ---")
        items = []
        # 1. 嘗試從 Config 讀取測試項目
        if 'Block2_Aging_Items' in self.config:
            self.log("Loading Aging Items from Config...")
            section = self.config['Block2_Aging_Items']
            idx = 1
            while True:
                key = f"Item_{idx}"
                if key not in section:
                    break
                
                try:
                    # 格式: Name | Command | Interrupt(1/0) | CaptureLog(1/0)
                    raw_val = section[key]
                    parts = [p.strip() for p in raw_val.split('|')]
                    
                    if len(parts) >= 4:
                        name = parts[0]
                        cmd = parts[1]
                        will_interrupt = (parts[2] == '1')
                        capture_log = (parts[3] == '1')
                        items.append((name, cmd, will_interrupt, capture_log))
                    else:
                        self.log(f"Warning: Invalid format in {key}, skipping.")
                except Exception as e:
                    self.log(f"Error parsing {key}: {e}")
                
                idx += 1

        if not items:
            self.log("Config [Block2_Aging_Items] not found or empty. Using default list.")
            items = [
                ("Battery Info",   r"call .\RI\BatteryInfo.bat", False, False),
                ("Battery Aging",  r"call .\RI\Battery.bat",     False, True),
                ("Screen On/Off",  r"call .\RI\TurnOnOff.bat",   False, True),
                ("Camera Test",    r"call .\RI\RICamera.bat",    False, True),
                ("Cold Boot",      r"call .\RI\ColdBoot.bat",    True,  True),
                ("RTC Check",      r"call .\RI\RTC.bat",         False, True),
                ("Memory Stress",  r"call .\RI\Memory.bat",      False, False),
                ("Storage Test",   r"call .\RI\HDD_CMD.bat",     False, True),
                ("3DMark Test",    r"call .\RI\3DMark.bat",      False, True),
                ("Fan Speed Set",  r"call .\RI\SetFanSpeed.bat", False, True),
                ("S3 Sleep Test",  r"call .\RI\S3sleeptest.bat", True,  True),
                ("S4 Sleep Test",  r"call .\RI\S4sleeptest.bat", True,  True),
                ("Driver Check",   r"call .\RI\CheckDriver.bat", False, True),
                ("BT/WiFi Test",   r"call .\RI\BTWIFI.bat",      False, True)
            ]
        
        for idx, (test_name, cmd, will_interrupt, capture_log) in enumerate(items):
            if idx < start_from_step: continue           
            # 更新 UI 狀態
            self.set_status(f"Cycle {current_cycle} | Aging Step {idx+1}/{len(items)}: {test_name}")
            self.log(f"Starting {test_name}...")
            self.log(f"Battery percentage:{psutil.sensors_battery().percent}")
            # S3/S4 特殊處理 (Save RUNNING)
            if "sleeptest" in cmd.lower() or "coldboot" in cmd.lower(): 
                if "coldboot" in cmd.lower():
                    self.save_state("2", idx, current_cycle, status="REBOOTING")
                    self.set_run_once_startup()
                else:
                    self.save_state("2", idx, current_cycle, status="RUNNING")
                self.exec_cmd_wait(cmd, capture_log=capture_log)
                self.save_state("2", idx + 1, current_cycle, status="IDLE")
                continue

            # 一般重開機測試
            if will_interrupt:
                self.save_state("2", idx + 1, current_cycle, status="IDLE")
                self.exec_cmd_wait(cmd, capture_log=capture_log)
                if "Boot" in cmd or "RTC" in cmd: return 
            else:
                self.exec_cmd_wait(cmd, capture_log=capture_log)
                self.save_state("2", idx + 1, current_cycle, status="IDLE")

    # ==========================================
    # Block 3 實作 (Battery)
    # ==========================================
    def run_block_3(self, current_cycle=1):
        try:
            tool_path = os.path.join(self.base_dir, "RI", "DiagECtool.exe") 
            fan_mode = self.config['Block1_Thermal'].get('Fan_Mode', None)
            if fan_mode and str(fan_mode).strip():
                fan_mode = str(fan_mode).strip()
                self.log(f"Set Fan Mode: {fan_mode}")
                self.exec_cmd_wait(f"{tool_path} raw --cmd 0x20 --subcmd 0x01 --data 0x03", capture_log=True)
                self.exec_cmd_wait(f"{tool_path} raw --cmd 0x20 --subcmd 0x06 --data 0x0{fan_mode}", capture_log=True)
                for _ in range(10):
                    self.check_stop()
                    QApplication.processEvents()
                    time.sleep(1)
            self.log("--- Block 3: Battery Charge/Discharge ---")
            self.set_status(f"Cycle {current_cycle} | Running Block 3: Battery Charge/Discharge") 
            self.log(f"Battery percentage:{psutil.sensors_battery().percent}")
            self.save_state("3", 0, current_cycle, status="RUNNING")
            self.exec_cmd_wait(r"call .\RI\BatteryControl.bat")
        except Exception as e:
            self.log(f"Error : {e}")
            raise e
        finally:
            self.exec_cmd_wait(f"{tool_path} battery --mode auto", capture_log=True)
            self.exec_cmd_wait(f"{tool_path} fan --mode auto", capture_log=True)
if __name__ == "__main__":
    try:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            # 6 = SW_MINIMIZE (縮小到工作列)
            # 0 = SW_HIDE (完全隱藏, 若需要完全隱藏可改用 0)
            ctypes.windll.user32.ShowWindow(hwnd, 6)
    except Exception as e:
        print(f"Failed to minimize console: {e}")

    if getattr(sys, 'frozen', False):
        # 如果是打包後的 exe，路徑要抓執行檔本身的位置
        application_path = os.path.dirname(sys.executable)
    else:
        # 如果是跑 .py script
        application_path = os.path.dirname(os.path.abspath(__file__))  

    os.chdir(application_path)
    print(f"Current Working Directory: {os.getcwd()}")
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)
    app = QApplication(sys.argv)
    lock_file_path = os.path.join(QDir.tempPath(), 'odm_runin_instance.lock')
    lock_file = QLockFile(lock_file_path)
    # 嘗試鎖定，Timeout 設定 100ms
    if not lock_file.tryLock(100):
        # 如果鎖定失敗，代表已經有一個實例在執行
        QMessageBox.critical(None, "Error", "Run-In program is already running!\n(Please close the existing window first)")
        sys.exit(1)

    win = ODM_RunIn_Project(title="ACER Run-In Test")
    win.show()
    sys.exit(app.exec_())
