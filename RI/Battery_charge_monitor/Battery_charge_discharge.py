import subprocess
import time
import sys
import logging
import configparser
import datetime
import xml.etree.ElementTree as ET
from pathlib import Path
import shutil
# === 日誌設定 ===
def init_logger(base_dir, prefix):
    # 清除舊的 handlers 避免重複
    root = logging.getLogger()
    if root.handlers:
        for handler in root.handlers:
            root.removeHandler(handler)

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(), 
            logging.FileHandler(base_dir / f"{prefix}.log", encoding='utf-8', mode='a')
        ]
    )

# === 路徑設定 ===
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent.resolve()

RESULT_DIR = BASE_DIR / "result"
RESULT_FILE = RESULT_DIR / "result.txt"
CONFIG_FILE = BASE_DIR / "Config.ini"
PRIME95_DIR = BASE_DIR / "Prime95"
PRIME95_EXE = PRIME95_DIR / "prime95.exe"

# === 外部指令設定 ===
CMD_AUTO_MODE = ["DiagECtool.exe", "battery", "--mode", "auto"]
CMD_DEBUG_MODE = ["DiagECtool.exe", "battery", "--mode", "debug"]
CMD_DISCHARGE = ["DiagECtool.exe", "battery", "--discharge"]
CMD_TIMEOUT = 15
CMD_KILL_PRIME95 = ["taskkill", "/f", "/im", "prime95.exe"]

def remove_old_result():
    if RESULT_DIR.exists():
        import shutil
        shutil.rmtree(RESULT_DIR)
    RESULT_DIR.mkdir(exist_ok=True)

def write_result(status: str, message: str = ""):
    content = f"{status}\n{message}"
    RESULT_FILE.write_text(content, encoding="utf-8")
    logging.info(f"Result written: {status} - {message}")

def load_config():
    """讀取 Config.ini 並回傳設定字典"""
    config = configparser.ConfigParser()
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(f"Config file not found: {CONFIG_FILE}")
    
    config.read(CONFIG_FILE, encoding="utf-8")
    
    try:
        backup_path_str = config.get("Log", "BackupPath", fallback=r"C:\Diag\Thermal")
        cfg = {
            "max_p": int(config["Percentage"]["maxPercentage"]),
            "min_p": int(config["Percentage"]["minPercentage"]),
            "interval": int(config["Time_interval"]["CheckInterval_Sec"]),
            "timeout_min": int(config["Test_Settings"]["Test_Min"]),
            # Config 單位是 mA，轉換為 A
            "target_current_a": int(config["Test_Settings"]["Current"]) / 1000.0,
            "backup_path": Path(backup_path_str)
        }
        
        if cfg["max_p"] < cfg["min_p"]:
            raise ValueError("maxPercentage must be >= minPercentage")
            
        return cfg
    except KeyError as e:
        raise ValueError(f"Missing config key: {e}")
    except ValueError as e:
        raise ValueError(f"Config value error: {e}")

def get_serial_number():
    try:
        sn = subprocess.check_output("wmic bios get serialnumber", shell=True, text=True)
        return sn.strip().split('\n')[-1].strip()
    except Exception:
        return "UNKNOWN_SN"

def get_battery_info():
    """
    獲取電量百分比與電流(Amps)。
    嘗試使用 WMI BatteryStatus 計算電流。
    """
    percentage = None
    current_amps = 0.0
    
    try:
        # 1. 獲取百分比
        p_res = subprocess.check_output(
            ["powershell", "-command", "(Get-CimInstance Win32_Battery).EstimatedChargeRemaining"],
            text=True
        )
        if p_res.strip().isdigit():
            percentage = int(p_res.strip())

        # 2. 獲取電流 (PowerShell WMI)
        cmd = "Get-WmiObject -Namespace root/wmi -Class BatteryStatus | Select-Object ChargeRate, Voltage"
        c_res = subprocess.check_output(["powershell", "-command", cmd], text=True)
        
        lines = c_res.strip().split('\n')
        rate = 0
        voltage = 0
        
        for line in lines:
            if "ChargeRate" in line:
                try: rate = int(line.split(':')[-1].strip())
                except: pass
            if "Voltage" in line:
                try: voltage = int(line.split(':')[-1].strip())
                except: pass
        
        if voltage > 0:
            current_amps = rate / voltage
            
    except Exception as e:
        logging.debug(f"Read battery info failed: {e}")

    return percentage, current_amps

def run_command(cmd, action_name, max_retries=3):
    for attempt in range(1, max_retries + 1):
        try:
            subprocess.run(cmd, check=True, timeout=CMD_TIMEOUT, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            logging.info(f"[CMD Success] {action_name}")
            return True 
        except Exception:
            if attempt < max_retries:
                time.sleep(1)
    logging.error(f"[CMD Failed] {action_name}")
    return False

def is_prime95_running():
    # 檢查 tasklist 中是否有 prime95.exe
    try:
        output = subprocess.check_output('tasklist', shell=True, text=True)
        return "prime95.exe" in output.lower()
    except:
        return False
    
def start_prime95():
    if not PRIME95_EXE.exists():
        logging.warning("Prime95 not found, skipping stress.")
        return

    if is_prime95_running():
        logging.info("Prime95 is already running. Skipping start.")
        return
    
    try:
        subprocess.Popen(
            [str(PRIME95_EXE), "-t", "-small", "-A16"], 
            cwd=str(PRIME95_DIR),
            creationflags=subprocess.CREATE_NEW_CONSOLE
        )
        logging.info("Prime95 Started (Accelerate Discharge)")
    except Exception as e:
        logging.error(f"Failed to start Prime95: {e}")

def kill_prime95():
    try:
        subprocess.run(CMD_KILL_PRIME95, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        logging.info("Prime95 Stopped (Idle Mode)")
    except Exception:
        pass

def enable_charging():
    kill_prime95()
    time.sleep(5)
    run_command(CMD_AUTO_MODE, "Set Charge Mode")

def disable_charging():
    run_command(CMD_DEBUG_MODE, "Set Debug Mode")
    time.sleep(1)
    if run_command(CMD_DISCHARGE, "Set Discharge Mode"):
        time.sleep(5)
        start_prime95()

def save_xml_log(avg_current, current_list, cfg):
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f"BatCurrent_{timestamp}.xml"
    file_path = RESULT_DIR / filename
    
    root = ET.Element("BatteryTest")
    ET.SubElement(root, "Result").text = "PASS" if avg_current >= cfg['target_current_a'] else "FAIL"
    ET.SubElement(root, "AverageCurrent").text = f"{avg_current:.3f}A"
    ET.SubElement(root, "Requirement").text = f">={cfg['target_current_a']}A"
    ET.SubElement(root, "ConfigInterval").text = f"{cfg['interval']}s"
    
    data_elem = ET.SubElement(root, "DataPoints")
    for i, curr in enumerate(current_list):
        pt = ET.SubElement(data_elem, "Point")
        pt.set("seq", str(i+1))
        pt.text = f"{curr:.3f}"

    tree = ET.ElementTree(root)
    tree.write(file_path, encoding="utf-8", xml_declaration=True)
    logging.info(f"XML Log generated: {filename}")
    return filename

def perform_backup(xml_filename, cfg):
    """將生成的 XML 複製到 Config 指定的路徑"""
    source_file = RESULT_DIR / xml_filename
    dest_dir = cfg["backup_path"]

    if not source_file.exists():
        logging.error("Source XML file not found, cannot backup.")
        return

    try:
        # 確保目標資料夾存在
        if not dest_dir.exists():
            dest_dir.mkdir(parents=True, exist_ok=True)
            logging.info(f"Created backup directory: {dest_dir}")

        shutil.copy(source_file, dest_dir)
        logging.info(f"XML Backup Success: Copied to {dest_dir}")

    except Exception as e:
        logging.error(f"XML Backup Failed: {e}")

def perform_stage(target_p, is_charging, cfg, current_recorder_list=None):
    """
    執行階段，包含超時檢查
    """
    logging.info(f"--- Starting Stage: {'CHARGE' if is_charging else 'DISCHARGE'} to {target_p}% ---")
    
    if is_charging:
        enable_charging()
    else:
        disable_charging()
    
    # 計算此階段的超時時間 (使用設定檔的總時間做為保護，或可自行定義單階段時間)
    start_time = time.time()
    timeout_sec = cfg['timeout_min'] * 60
    
    while True:
        # 檢查超時
        if (time.time() - start_time) > timeout_sec:
            raise TimeoutError(f"Stage timeout after {cfg['timeout_min']} mins")

        bat, amps = get_battery_info()
        
        if bat is None:
            time.sleep(1)
            continue
        
        logging.info(f"Current Battery: {bat}% | Amps: {amps:.3f}A")

        if is_charging and current_recorder_list is not None:
            current_recorder_list.append(amps)

        # 檢查是否到達目標
        if is_charging:
            if bat >= target_p:
                logging.info(f"Target {target_p}% Reached.")
                break
        else:
            if bat <= target_p:
                logging.info(f"Target {target_p}% Reached.")
                break
        
        # 使用 Config 定義的間隔
        time.sleep(cfg['interval'])

def main():
    remove_old_result()
    init_logger(BASE_DIR, prefix="Battery_Charge_Test")
    logging.info("========== Battery Cycle Test  ==========")
    try:
        # 1. 讀取設定
        cfg = load_config()
        logging.info(f"Config Loaded: Range={cfg['min_p']}-{cfg['max_p']}%, Interval={cfg['interval']}s, TargetCurr={cfg['target_current_a']}A")
        
        # 2. 取得初始電量
        init_bat, _ = get_battery_info()
        while init_bat is None:
            logging.warning("Waiting for battery reading...")
            time.sleep(2)
            init_bat, _ = get_battery_info()

        logging.info(f"Initial Battery: {init_bat}%")
        
        charging_currents = []

        # 3. 執行邏輯 ( > Max 則先放電, 否則先充電)
        # 注意：這裡使用 Config 的 max_p 和 min_p
        
        if init_bat > cfg['max_p']:
            logging.info(f"Scenario: > {cfg['max_p']}%. Flow: Discharge({cfg['min_p']}) -> Charge({cfg['max_p']})")
            # Step 1: Discharge
            perform_stage(cfg['min_p'], is_charging=False, cfg=cfg)
            # Step 2: Charge
            perform_stage(cfg['max_p'], is_charging=True, cfg=cfg, current_recorder_list=charging_currents)
            
        else:
            logging.info(f"Scenario: <= {cfg['max_p']}%. Flow: Charge({cfg['max_p']}) -> Discharge({cfg['min_p']})")
            
            if init_bat < cfg['max_p']:
                 # Step 1: Charge
                 perform_stage(cfg['max_p'], is_charging=True, cfg=cfg, current_recorder_list=charging_currents)
            
            # Step 2: Discharge
            perform_stage(cfg['min_p'], is_charging=False, cfg=cfg)

        # 4. 結算
        logging.info("Cycle Finished. analyzing data...")
        
        if charging_currents:
            valid_currents = [c for c in charging_currents if c > 0]
            avg_cur = sum(valid_currents) / len(valid_currents) if valid_currents else 0.0
        else:
            avg_cur = 0.0

        logging.info(f"Average Charge Current: {avg_cur:.3f} A (Target: {cfg['target_current_a']} A)")
        
        xml_name = save_xml_log(avg_cur, charging_currents, cfg)
        perform_backup(xml_name, cfg)
        
        if avg_cur >= cfg['target_current_a']:
            write_result("PASS", f"Avg Current {avg_cur:.2f}A >= {cfg['target_current_a']}A. XML: {xml_name}")
        else:
            write_result("FAIL", f"Avg Current {avg_cur:.2f}A < {cfg['target_current_a']}A. XML: {xml_name}")

    except KeyboardInterrupt:
        logging.warning("Interrupted by user.")
    except Exception as e:
        logging.exception(f"Error: {e}")
        write_result("FAIL", str(e))
    finally:
        kill_prime95()
        try:
            run_command(CMD_AUTO_MODE, "Restore Charge Mode")
        except:
            pass

if __name__ == "__main__":
    main()