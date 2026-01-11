import sys
import os
import time
import psutil 
import csv
import subprocess
import threading
import shutil  # 用於複製檔案
from datetime import datetime, timedelta
from PyQt5.QtWidgets import QApplication, QMessageBox
from core import BaseRunInApp
from PyQt5.QtCore import Qt, QThread, QLockFile, QDir, QTimer

# ==========================================
# Helper: 風扇監控執行緒 (背景執行)
# ==========================================
class FanMonitorThread(QThread):
    def __init__(self, csv_path, interval=5):
        super().__init__()
        self.csv_path = csv_path
        self.interval = interval
        self.running = True

    def run(self):
        # 確保目錄存在
        os.makedirs(os.path.dirname(self.csv_path), exist_ok=True)
        
        # 寫入 CSV Header
        with open(self.csv_path, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(["Timestamp", "Fan1_RPM", "Fan2_RPM"])

        while self.running:
            try:
                now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                rpm1 = self.get_rpm(1)
                rpm2 = self.get_rpm(2)
                
                with open(self.csv_path, 'a', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow([now, rpm1, rpm2])
            except Exception as e:
                pass 

            # 間隔休息
            for _ in range(self.interval):
                if not self.running: break
                self.msleep(1000)

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
        QTimer.singleShot(1000, self.check_auto_run)

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

            # 偵測非預期當機 (上次狀態仍為 RUNNING)
            if last_status == "RUNNING":
                self.log(f"!!! DETECTED CRASH/HANG AT BLOCK {current_block} STEP {current_step} !!!")
                self.log("Marking this step as FAIL and skipping...")
                
                # 簡單的 Fail Log
                with open("RunIn_Crash.log", "a") as f:
                    f.write(f"Crash at Block {current_block} Step {current_step}\n")

                # 跳過該步驟，避免無限重啟
                current_step += 1
                self.save_state(current_block, current_step, current_cycle, status="IDLE")

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
                self.run_block_3()
            else:
                self.log("Block 3 is Disabled. Skipping...")
            
            self.clear_state()
            self.log("=== ALL BLOCKS FINISHED ===")
            # [新增] Auto Close 檢查邏輯
            if 'Global' in self.config and self.config['Global'].getboolean('AutoClose'):
                self.log("[AutoClose] Closing application in 3 seconds...")
                # 為了讓使用者看到 PASS，延遲 3 秒再關閉
                import time
                time.sleep(3)
                QApplication.quit()

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
            return
        try:
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            new_filename = f"{timestamp}_{prefix_name}.csv"
            # 假設 log_dir 在 run_block_1 已經定義，或直接存到 src_path 同目錄
            dest_dir = os.path.dirname(src_path)
            dest_path = os.path.join(dest_dir, new_filename)
            
            shutil.copy2(src_path, dest_path)
            self.log(f"Fan Log archived: {new_filename}")
        except Exception as e:
            self.log(f"Error archiving fan log: {e}")

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
        if not os.path.exists(csv_path): return ["PTAT Log not found"]

        # 1. 找出 Config 定義的 Keys
        ptat_keys = []
        for key, value in self.config['Block1_Thermal'].items():
            if key.lower().startswith('ptat_key_'):
                ptat_keys.append(value)
        
        if not ptat_keys:
            self.log("No PTAT_Key defined in Config. Skipping check.")
            return

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
                self.log(f"WARNING: Column '{target_col_name}' not found! Skipping.")
                continue
                
            col_idx = header_map[target_col_name]
            avg_val = self.analyze_ptat_log(csv_path, col_idx, duration_sec=120)
            
            try:
                cfg_low_key = f"{test_mode}_Low"
                cfg_high_key = f"{test_mode}_High"
                
                limit_low = float(self.config[target_col_name][cfg_low_key])
                limit_high = float(self.config[target_col_name][cfg_high_key])
                if avg_val < limit_low or avg_val > limit_high:
                    msg = f"{target_col_name} FAIL: {avg_val:.2f} (Spec: {limit_low}~{limit_high})"
                    self.log(msg)
                    errors.append(msg)
                else:
                    self.log(f"PASS: {target_col_name} = {avg_val:.2f} (Spec: {limit_low}~{limit_high})")

            except KeyError:
                self.log(f"WARNING: Config key '{cfg_low_key}/{cfg_high_key}' missing for [{target_col_name}]")
            except ValueError:
                self.log(f"WARNING: Invalid value for [{target_col_name}]")
        
        return errors
    
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
        # 1. 找出 Config Keys
        gpu_keys = []
        for key, value in self.config['Block1_Thermal'].items():
            if key.lower().startswith('gpumon_key_'):
                gpu_keys.append(value)
        
        if not gpu_keys: return []
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
                
                if avg_val < limit_low or avg_val > limit_high:
                    msg = f"GPUMon {target_col} FAIL: {avg_val:.2f} (Spec: {limit_low}~{limit_high})"
                    self.log(msg)
                    errors.append(msg)
                else:
                    self.log(f"GPUMon PASS: {target_col} = {avg_val:.2f}")
            except Exception as e:
                errors.append(f"GPUMon Config Error [{target_col}]: {e}")               
        return errors
    # ==========================================
    # Block 1 實作 (Thermal)
    # ==========================================
    def run_block_1(self, start_from_step=0, current_cycle=1):
        self.log("--- Block 1: Thermal Tool ---")
        log_dir = r"C:\Diag\Thermal"
        os.makedirs(log_dir, exist_ok=True)
        
        # Step 0: 電量檢查
        if start_from_step == 0:
            threshold = int(self.config['Block1_Thermal']['Start_Battery_Threshold'])
            self.log(f"Step 0: Waiting for Battery > {threshold}%...")
            
            while True:
                # [安全修正] 檢查停止訊號
                self.check_stop()
                QApplication.processEvents() 
                try:
                    bat = psutil.sensors_battery()
                    if bat:
                        current_pct = bat.percent
                        is_plugged = bat.power_plugged
                        status_str = "Charging" if is_plugged else "Discharging"
                        self.log(f"Battery Status: {current_pct}% ({status_str}) / Target: {threshold}%")

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
                time.sleep(5)
            
            self.save_state("1", 1, current_cycle, status="IDLE")

        # ==========================================
        # [Test 1] Single Stress
        # ==========================================
        if start_from_step <= 1:
            self.log("[Test 1] Single Stress Start")
            duration = int(self.config['Block1_Thermal']['Test1_Duration'])
            fan_log = os.path.join(log_dir, "CPU_Fan.csv")
            fan_thread = None
            gpumon_keys = [k for k in self.config['Block1_Thermal'] if k.lower().startswith('gpumon_key_')]
            is_gpumon_enabled = len(gpumon_keys) > 0
            if is_gpumon_enabled:
                self.log(f"GPUMon Enabled (Found {len(gpumon_keys)} keys).")
            else:
                self.log("GPUMon Disabled (No keys found in Config).")
            try:
                # 1. 啟動 Prime95 (Tool) & Furmark
                cmd_furmark = r"start .\RI\FurMark\FurMark_GUI.exe"
                self.log(f"Starting furmark: {cmd_furmark}")
                p_stress = subprocess.Popen(cmd_furmark, shell=True)

                for _ in range(10):
                    self.check_stop() # [安全修正] 檢查 STOP
                    QApplication.processEvents()
                    time.sleep(1)

                cmd_stress = r".\RI\prime95\prime95.exe -t -small -A16"
                self.log(f"Starting Prime95: {cmd_stress}")
                p_stress = subprocess.Popen(cmd_stress, shell=True)
                
                # 2. 等待 40 秒
                self.log("Waiting 40s before starting PTAT...")
                for _ in range(40):
                    self.check_stop() # [安全修正] 檢查 STOP
                    QApplication.processEvents()
                    time.sleep(1)

                # 3. 啟動 PTAT (Monitor)
                ptat_dir = r"C:\Program Files\Intel Corporation\Intel(R)PTAT"
                ptat_cmd = "PTAT.exe -start -w=cpu.json"
                self.log(f"Starting PTAT: {ptat_cmd}")
                p_ptat = subprocess.Popen(ptat_cmd, cwd=ptat_dir, shell=True)

                # 4. 等待 20 秒
                self.log("Waiting 20s before starting Fan Monitor...")
                for _ in range(20):
                    self.check_stop() # [安全修正] 檢查 STOP
                    QApplication.processEvents()
                    time.sleep(1)

                # 5. 啟動 Fan Monitor
                fan_thread = FanMonitorThread(fan_log)
                fan_thread.start()

                # --- [新增] 啟動 GPUMon ---
                if is_gpumon_enabled:
                    gpu_mon_dir = os.path.join(self.base_dir, "RI", "GPUMon")
                    # 確保資料夾存在，若無則建立(或是依賴已經存在)
                    if not os.path.exists(gpu_mon_dir): os.makedirs(gpu_mon_dir)
                    # 指令: -custom:timestamp,temp,pwr,clk -wake -log:cpu_gpumon.csv
                    gpu_cmd = "GPUMonCmd.exe -custom:timestamp,temp,pwr,clk -wake -log:cpu_gpumon.csv"
                    self.log(f"Starting GPUMon: {gpu_cmd}")
                    # 設定 cwd 為 RI\GPUMon，這樣 log 才會產在該資料夾內
                    p_gpumon = subprocess.Popen(gpu_cmd, cwd=gpu_mon_dir, shell=True)

                # 6. 正式等待測試時間
                self.log(f"Running Stress for {duration} seconds...")
                for i in range(duration):
                    if i % 10 == 0: QApplication.processEvents()
                    self.check_stop()
                    time.sleep(1)

            except Exception as e:
                self.log(f"Test 1 Interrupted or Error: {e}")
                raise e # 往上拋，讓主流程處理停止

            finally:
                # Teardown: 確保工具被殺掉 (無論是按 Stop 還是正常結束)
                self.log("Stopping Tools (Teardown)...")                
                # A. 殺 PTAT
                ptat_dir = r"C:\Program Files\Intel Corporation\Intel(R)PTAT"
                ptat_cmd = "PTAT.exe -stop"
                self.log(f"Stop PTAT: {ptat_cmd}")
                p_ptat = subprocess.Popen(ptat_cmd, cwd=ptat_dir, shell=True)
                try:
                    p_ptat.wait(timeout=60)
                except:
                    pass
                self.log("Waiting for PTAT to finish writing log...")
                self.ensure_process_killed("PTAT.exe")
                self.ensure_process_killed("GPUMonCmd.exe")
                for i in range(15):
                    if i % 10 == 0: QApplication.processEvents()
                    time.sleep(1)
                time.sleep(1) # 稍等檔案釋放
                # B. 停 Fan Monitor
                if fan_thread:
                    fan_thread.stop()
                #self.ensure_process_killed("FanControl.exe")
                
                # C. 殺 Prime95
                self.ensure_process_killed("prime95.exe")                
                time.sleep(2) 
                self.ensure_process_killed("FurMark_GUI.exe")  
                time.sleep(2) 
            # 8. 結果判定 (若前面被 Stop 中斷，這裡不會執行)
            self.log("Verifying Results...")
            all_failures = [] #錯誤收集器
            self.archive_fan_log(fan_log, "CPU_Fan")
            # (A) Fan Check
            # 讀取 Config
            try:
                spec_min = int(self.config['Block1_Thermal']['Test1_Fan_Min'])
                spec_max = int(self.config['Block1_Thermal']['Test1_Fan_Max'])
                self.log(f"CPU only fan spec_min:{spec_min}, CPU only fan spec_max:{spec_max}")
                avg_fan1 = self.analyze_fan_log_average(fan_log, 1) 
                avg_fan2 = self.analyze_fan_log_average(fan_log, 2) 
                if not (spec_min <= avg_fan1 <= spec_max) or not (spec_min <= avg_fan2 <= spec_max):
                    msg = f"Fan RPM FAIL: Fan1={avg_fan1}, Fan2={avg_fan2} (Spec: {spec_min}-{spec_max})"
                    self.log(msg)
                    all_failures.append(msg)
                else:
                    self.log(f"Fan RPM PASS: {avg_fan1}, {avg_fan2}")           
            except Exception as e:
                msg = f"Fan Check Error: {e}"
                self.log(msg)
                all_failures.append(msg)
            # (B) PTAT Check
            user_home = os.path.expanduser("~")
            ptat_log_dir = os.path.join(user_home, "Documents", "iPTAT", "log")           
            ptat_log = self.find_latest_log(ptat_log_dir, prefix="PTATMonitor")
            if ptat_log:
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                new_filename = f"{timestamp}_CPU_PTAT.csv"
                dest_path = os.path.join(log_dir, new_filename)
                
                try:
                    self.log(f"Copying PTAT Log to {dest_path}...")
                    shutil.copy2(ptat_log, dest_path)
                    
                    # 使用複製後的檔案進行檢查
                    ptat_errors = self.check_ptat_metrics(dest_path, test_mode="Test1") 
                    if ptat_errors:
                        all_failures.extend(ptat_errors) # 將 PTAT 錯誤加入總名單
                    else:
                        self.log("PTAT Check PASS (All items).")
                except Exception as e:
                    msg = f"PTAT File Process Error: {e}"
                    self.log(msg)
                    all_failures.append(msg)
            else:
                self.log(f"WARNING: No PTAT Log found in {ptat_log_dir}")
                all_failures.append("PTAT Log missing")
            # --- [新增] 3. GPUMon Check ---
            # 來源檔案: RI\GPUMon\cpu_gpumon.csv
            src_gpu_log = os.path.join(gpu_mon_dir, "cpu_gpumon.csv")        
            if os.path.exists(src_gpu_log):
                # 備份檔案: C:\Diag\Thermal\TIMESTAMP_Test1_GPUMon.csv
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                dest_gpu_name = f"{timestamp}_Test1_GPUMon.csv"
                dest_gpu_path = os.path.join(log_dir, dest_gpu_name)              
                try:
                    shutil.copy2(src_gpu_log, dest_gpu_path)
                    self.log(f"GPUMon Log archived: {dest_gpu_name}")                    
                    # 執行檢查 (Test1)
                    gpu_errors = self.check_gpumon_metrics(dest_gpu_path, test_mode="Test1")
                    if gpu_errors:
                        all_failures.extend(gpu_errors)
                    else:
                        self.log("GPUMon Check PASS.")                       
                except Exception as e:
                    self.log(f"GPUMon Copy/Check Error: {e}")
                    all_failures.append(f"GPUMon Error: {e}")
            else:
                self.log("WARNING: GPUMon Log not found!")
                all_failures.append("GPUMon Log missing") # 視需求開啟            
            # ==========================================
            # 最終判定 (若有任何錯誤，才 Raise)
            # ==========================================
            if all_failures:
                # 將所有錯誤組合成一個字串拋出
                error_summary = " | ".join(all_failures)
                raise Exception(f"Test 1 FAILED: {error_summary}")
            
            self.log("Test 1 PASS.")
            # 9. Reboot Check
            self.save_state("1", 2, current_cycle, status="IDLE")
            if self.config['Block1_Thermal'].getboolean('Test1_Reboot'):
                self.log("Rebooting as per config...")
                self.trigger_reboot()
            else:
                self.log("Continuing to next step...")

        # Step 2: Fan Max Speed
        if start_from_step <= 2:
            self.log("[Test 2] Fan Max Speed")
            self.exec_cmd_wait(r"call .\RI\Thermal_Fan.bat")
            
            self.save_state("1", 3, current_cycle, status="IDLE")
            if self.config['Block1_Thermal'].getboolean('Test2_Reboot'):
                self.log("Rebooting for Test 3...")
                self.trigger_reboot()

        # Step 3: Dual Stress
        if start_from_step <= 3:
            self.log("[Test 3] Dual Stress")
            self.exec_cmd_wait(r"call .\RI\Thermal_Dual.bat")
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
        items = [(r"call .\RI\BatteryInfo.bat", False, False),
                 (r"call .\RI\Battery.bat", False, True),
                 (r"call .\RI\TurnOnOff.bat", False, True),
                 (r"call .\RI\RICamera.bat", False, True),
                 # (r"call .\RI\ColdBoot.bat", True, True),
                 (r"call .\RI\RTC.bat", False, True),
                 (r"call .\RI\Memory.bat", False, True),
                 (r"call .\RI\HDD_CMD.bat", False, True),
                 (r"call .\RI\3DMark.bat", False, True),
                 (r"call .\RI\SetFanSpeed.bat", False, True),
                 (r"call .\RI\S3sleeptest.bat", True, True),
                 (r"call .\RI\S4sleeptest.bat", True, True),
                 (r"call .\RI\CheckDriver.bat", False, True),
                 (r"call .\RI\BTWIFI.bat", False, True)]
        
        for idx, (cmd, will_interrupt, capture_log) in enumerate(items):
            if idx < start_from_step: continue
            
            # S3/S4 特殊處理 (Save RUNNING)
            if "sleeptest" in cmd: 
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
    def run_block_3(self):
        self.log("--- Block 3: Battery Charge/Discharge ---")
        self.exec_cmd_wait(r"call .\RI\BatteryControl.bat")
        
if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
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
