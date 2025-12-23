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
            self.log(f">>> RESUMING Block {current_block}, Step {current_step} <<<")

        # ---------------------------------------------------
        # Block 1: Thermal
        # ---------------------------------------------------
        if current_block == "1" and self.config['Block1_Thermal'].getboolean('Enabled'):
            self.run_block_1(start_from_step=current_step)
            current_block = "2"
            current_step = 0
            self.save_state("2", 0, current_cycle)

        # ---------------------------------------------------
        # Block 2: Aging
        # ---------------------------------------------------
        if current_block == "2" and self.config['Block2_Aging'].getboolean('Enabled'):
            self.run_block_2(start_from_step=current_step)
            current_block = "3"
            current_step = 0
            self.save_state("3", 0, current_cycle)

        # ---------------------------------------------------
        # Block 3: Battery
        # ---------------------------------------------------
        if current_block == "3" and self.config['Block3_Battery'].getboolean('Enabled'):
            self.run_block_3()
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
        
        items = [
            ("call .\\RI\\BatteryInfo.bat", False),
            ("call .\\RI\\Battery.bat", False),
            ("call .\\RI\\TurnOnOff.bat", False),
            ("call .\\RI\\RICamera.bat", False),
            ("call .\\RI\\ColdBoot.bat", True),
            ("call .\\RI\\RTC.bat", True),
            ("call.\\RI\\Memory.bat", False),
            ("call .\\RI\\HDD_CMD.bat", False),
            ("call .\\RI\\3DMark.bat", False),
            ("call .\\RI\\SetFanSpeed.bat", False),
            ("call .\\RI\\S3sleeptest.bat", True),
            ("call .\\RI\\S4sleeptest.bat", True),
            ("call .\\RI\\CheckDriver.bat", False),
            ("call .\\RI\\BTWIFI.bat", False)
        ]

        for idx, (cmd, will_interrupt) in enumerate(items):
            if idx < start_from_step: continue

            if will_interrupt:
                self.save_state("2", idx + 1)
                self.exec_cmd_wait(cmd)
                if "Boot" in cmd or "RTC" in cmd: return 
            else:
                self.exec_cmd_wait(cmd)
                self.save_state("2", idx + 1)

    # --- Block 3 實作 ---
    def run_block_3(self):
        self.log("--- Block 3: Battery Charge/Discharge ---")
        self.exec_cmd_wait(r"call .\RI\Battery_Cycling_Test.bat")

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
