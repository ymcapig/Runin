import sys
import os
import time
import psutil 
from PyQt5.QtWidgets import QApplication
from core import BaseRunInApp
from PyQt5.QtCore import Qt

class ODM_RunIn_Project(BaseRunInApp):

    def user_test_sequence(self):
        state = self.load_state()
        current_block = "1"
        current_step = 0
        current_cycle = 1
        # 斷點續傳恢復
        if state:
            current_block = state.get("block", "1")
            current_step = state.get("step", 0)
            current_cycle = state.get("cycle", 1)
            last_status = state.get("status", "IDLE") # 讀取上次狀態

            self.log(f">>> RESUMING Block {current_block}, Step {current_step} <<<")

            # 如果上次紀錄是 "RUNNING"，代表跑到一半當機或斷電重開
            if last_status == "RUNNING":
                self.log(f"!!! DETECTED CRASH/HANG AT BLOCK {current_block} STEP {current_step} !!!")
                self.log("Marking this step as FAIL and skipping...")
                
                # 1. 產生 FAIL 檔案
                self.generate_result_file(False)
                
                # 2. 強制跳過這個會當機的步驟，避免無限迴圈
                current_step += 1
                
                # 3. 更新狀態為 IDLE，讓流程繼續
                self.save_state(current_block, current_step, current_cycle, status="IDLE")
        # ---------------------------------------------------
        # Block 1: Thermal
        # ---------------------------------------------------
        if current_block == "1":
            if self.config['Block1_Thermal'].getboolean('Enabled'):
                self.run_block_1(start_from_step=current_step)
                # 跑完後切換到 Block 2
                current_block = "2"
                current_step = 0
                self.save_state("2", 0, current_cycle)
            else:
                self.log("Block 1 is Disabled. Skipping to Block 2...")
                current_block = "2" # 即使不跑，也要切換狀態

        # ---------------------------------------------------
        # Block 2: Aging
        # ---------------------------------------------------
        if current_block == "2":
            if self.config['Block2_Aging'].getboolean('Enabled'):
                self.run_block_2(start_from_step=current_step)
                current_block = "3"
                current_step = 0
                self.save_state("3", 0, current_cycle)
            else:
                self.log("Block 2 is Disabled. Skipping to Block 3...")
                current_block = "3"

        # ---------------------------------------------------
        # Block 3: Battery
        # ---------------------------------------------------
        if current_block == "3":
            if self.config['Block3_Battery'].getboolean('Enabled'):
                self.run_block_3()
            else:
                self.log("Block 3 is Disabled. Skipping...")
            
            # 全部跑完 (或全部跳過)
            self.clear_state()
            self.log("=== ALL BLOCKS FINISHED ===")
        
# --- Block 1 實作 ---
    def run_block_1(self, start_from_step=0):
        self.log("--- Block 1: Thermal Tool ---")
        
        # Step 0: 電量檢查 (每分鐘檢查一次，直到達標)
        if start_from_step == 0:
            threshold = int(self.config['Block1_Thermal']['Start_Battery_Threshold'])
            self.log(f"Step 0: Waiting for Battery > {threshold}%...")
            
            while True:
                self.check_stop()
                try:
                    bat = psutil.sensors_battery()
                    if bat:
                        current_pct = bat.percent
                        is_plugged = bat.power_plugged
                        status_str = "Charging" if is_plugged else "Discharging"
                        
                        self.log(f"Battery Status: {current_pct}% ({status_str}) / Target: {threshold}%")

                        if current_pct >= threshold:
                            self.log("Battery Threshold Reached! Starting tests...")
                            break
                        
                        if not is_plugged:
                            self.log("WARNING: AC Adapter not plugged in!")
                    else:
                        self.log("No battery detected (Desktop?). Skipping check.")
                        break

                except Exception as e:
                    self.log(f"Error reading battery: {e}")

                # 等待 60 秒 (1分鐘) 再檢查
                self.check_stop()
                time.sleep(5)
            
        # Step 1: 單燒
        if start_from_step <= 1:
            self.log("[Test 1] Single Stress")
            self.exec_cmd_wait(r"call .\RI\Thermal_Single.bat")
            self.save_state("1", 2)

        # Step 2: Fan
        if start_from_step <= 2:
            self.log("[Test 2] Fan Max Speed")
            self.exec_cmd_wait(r"call .\RI\Thermal_Fan.bat")
            self.log("Rebooting for Test 3...")
            self.save_state("1", 3)
            self.trigger_reboot()

        # Step 3: 雙燒
        if start_from_step <= 3:
            self.log("[Test 3] Dual Stress")
            self.exec_cmd_wait(r"call .\RI\Thermal_Dual.bat")
            self.log("Block 1 PASS. Rebooting to Block 2...")
            self.save_state("2", 0)
            self.trigger_reboot()

    # --- Block 2 實作 ---
    def run_block_2(self, start_from_step=0):
        self.log("--- Block 2: Aging Test ---")
        items = [("cd .\\RI && call .\\BatteryInfo.bat", False, False),
                ("cd .\\RI && call .\\Battery.bat", False, True),
                ("cd .\\RI && call .\\TurnOnOff.bat", False, True),
                ("cd .\\RI && call .\\RICamera.bat", False, True),
                #("cd .\\RI && call .\\ColdBoot.bat", True, True),
                ("cd .\\RI && call .\\RTC.bat", False, True),
                ("cd .\\RI && call .\\Memory.bat", False, True),
                ("cd .\\RI && call .\\HDD_CMD.bat", False, True),
                ("cd .\\RI && call .\\3DMark.bat", False, True),
                ("cd .\\RI && call .\\SetFanSpeed.bat", False, True),
                ("cd .\\RI && call .\\S3sleeptest.bat", True, True),
                ("cd .\\RI && call .\\S4sleeptest.bat", True, True),
                ("cd .\\RI && call .\\CheckDriver.bat", False, True),
                ("cd .\\RI && call .\\BTWIFI.bat", False, True)]
        # 這裡解包原本是 (cmd, will_interrupt)，現在變成 (cmd, will_interrupt, capture_log)
        for idx, (cmd, will_interrupt, capture_log) in enumerate(items):
            if idx < start_from_step: continue
            # [特例處理] S3 或 S4 測試
            if "sleeptest" in cmd: 
                # 1. 執行前：標記狀態為 RUNNING (正在跑)
                #    如果不幸睡死重開，下次啟動就會抓到這個標記
                self.save_state("2", idx, status="RUNNING")
                # 2. 執行測試
                self.exec_cmd_wait(cmd, capture_log=capture_log)
                # 3. 執行後 (成功醒來)：標記為 IDLE 並推進到下一步
                self.save_state("2", idx + 1, status="IDLE")
                continue

            # 一般的 ColdBoot / RTC 重開機測試
            # 這些測試本來就預期會重開，所以維持「先存下一步」的邏輯
            if will_interrupt:
                self.save_state("2", idx + 1, status="IDLE")
                self.exec_cmd_wait(cmd, capture_log=capture_log)
                if "Boot" in cmd or "RTC" in cmd: return 
            else:
                # 一般測試
                self.exec_cmd_wait(cmd, capture_log=capture_log)
                self.save_state("2", idx + 1, status="IDLE")

    # --- Block 3 實作 ---
    def run_block_3(self):
        self.log("--- Block 3: Battery Charge/Discharge ---")
        self.exec_cmd_wait(r"call .\RI\BatteryControl.bat")
        
if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)
    app = QApplication(sys.argv)
    win = ODM_RunIn_Project(title="ODM Run-In Framework (PyQt5)")
    win.show()
    # PyQt5 建議使用 exec_() (雖然新版 Python 支援 exec，但舊版 PyQt5 需加底線)
    sys.exit(app.exec_())
