import subprocess
import time
import sys
import logging
import configparser
from pathlib import Path

# 引用原本的 log 設定
from log_setting import init_logger

# === 路徑設定 ===
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent.resolve()

RESULT_DIR = BASE_DIR / "result"
RESULT_FILE = RESULT_DIR / "result.txt"

# [新增] Prime95 相關路徑
PRIME95_DIR = BASE_DIR / "Prime95"
PRIME95_EXE = PRIME95_DIR / "prime95.exe"

# === 外部指令設定 ===
CMD_AUTO_MODE = ["DiagECtool.exe", "battery", "--mode", "auto"]
CMD_DEBUG_MODE = ["DiagECtool.exe", "battery", "--mode", "debug"]
CMD_DISCHARGE = ["DiagECtool.exe", "battery", "--discharge"]
CMD_TIMEOUT = 15

# [新增] 砍掉 Prime95 的指令
CMD_KILL_PRIME95 = ["taskkill", "/f", "/im", "prime95.exe"]

def remove_old_result():
    if RESULT_DIR.exists():
        import shutil
        shutil.rmtree(RESULT_DIR)
    RESULT_DIR.mkdir(exist_ok=True)

def write_result(status: str, message: str = ""):
    content = f"{status}\n{message}"
    RESULT_FILE.write_text(content, encoding="utf-8")
    logging.info(f"Result written: {status}")

def load_config():
    config_path = BASE_DIR / "Config.ini"
    config = configparser.ConfigParser()

    if not config_path.exists():
        raise FileNotFoundError("Config.ini not found")

    config.read(config_path, encoding="utf-8")

    try:
        max_p = int(config["Percentage"]["maxPercentage"])
        min_p = int(config["Percentage"]["minPercentage"])
        interval = int(config["Time_interval"]["CheckInterval_Sec"])
        
        duration_min = int(config.get("Test_Settings", "TestDuration_Min", fallback=60))
        tolerance = int(config.get("Test_Settings", "Tolerance_Percentage", fallback=2))

    except Exception:
        raise ValueError("Config.ini format error")

    if max_p < min_p:
        raise ValueError("maxPercentage must be >= minPercentage")

    return {
        "max": max_p,
        "min": min_p,
        "interval": interval,
        "duration_sec": duration_min * 60,
        "tolerance": tolerance
    }

def get_battery_percentage():
    try:
        result = subprocess.check_output(
            ["powershell", "-command", "(Get-CimInstance Win32_Battery).EstimatedChargeRemaining"],
            text=True
        )
        for line in result.splitlines():
            line = line.strip()
            if line.isdigit():
                return int(line)
    except Exception as e:
        logging.error(f"Read battery failed: {e}")
    return None

def run_command(cmd, action_name, max_retries=5):
    for attempt in range(1, max_retries + 1):
        try:
            logging.info(f"[RUN] {action_name} (Attempt {attempt}/{max_retries})")
            subprocess.run(cmd, check=True, timeout=CMD_TIMEOUT)
            return True 
        except subprocess.TimeoutExpired:
            logging.warning(f"[TIMEOUT] {action_name}")
        except subprocess.CalledProcessError as e:
            # Taskkill 如果找不到 process 會回傳錯誤，這是正常的，不需大驚小怪
            if "taskkill" in cmd[0]:
                 logging.info(f"{action_name} - Process might not be running (RC={e.returncode})")
                 return True
            logging.warning(f"[FAIL] {action_name} return code: {e.returncode}")
        except Exception as e:
            logging.warning(f"[ERROR] {action_name}: {e}")
        
        if attempt < max_retries:
            time.sleep(1)

    logging.error(f"[GIVE UP] {action_name} failed after {max_retries} attempts.")
    return False

# [新增] 啟動 Prime95 函式
def start_prime95():
    if not PRIME95_EXE.exists():
        logging.error(f"Prime95 not found at {PRIME95_EXE}, skipping stress test.")
        return

    try:
        logging.info("Starting Prime95 for faster discharge...")
        # 注意: 使用 Popen 讓它在背景執行，不會卡住程式
        # cwd 設定為 Prime95 目錄，避免找不到設定檔
        # 修正您的拼字: -samll -> -small
        subprocess.Popen(
            [str(PRIME95_EXE), "-t", "-small", "-A16"], 
            cwd=str(PRIME95_DIR),
            creationflags=subprocess.CREATE_NEW_CONSOLE # 視窗獨立，避免被主程式影響
        )
    except Exception as e:
        logging.error(f"Failed to start Prime95: {e}")

# [新增] 關閉 Prime95 函式
def kill_prime95():
    logging.info("Stopping Prime95...")
    # 這裡只嘗試一次即可，不用 retry loop，因為只要送出訊號就好
    try:
        subprocess.run(CMD_KILL_PRIME95, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass

def enable_charging():
    # 1. 先關閉 Prime95
    kill_prime95()
    
    # 2. 等待 5 秒
    logging.info("Waiting 5s before enabling charge...")
    time.sleep(5)
    
    # 3. 下指令充電
    if not run_command(CMD_AUTO_MODE, "Set Auto mode (Charge)"):
        logging.error("CRITICAL: Failed to enable charging after retries.")

def disable_charging():
    # 1. 先下指令放電
    if run_command(CMD_DEBUG_MODE, "Set Debug mode"):
        time.sleep(2)
        if run_command(CMD_DISCHARGE, "Set Discharging"):
            # 2. 指令成功後，等待 5 秒
            logging.info("Discharge command sent. Waiting 5s before Prime95...")
            time.sleep(5)
            
            # 3. 啟動 Prime95 加速放電
            start_prime95()
        else:
            logging.error("Discharge command failed, skipping Prime95.")
    else:
        logging.error("Debug mode failed, skipping discharge logic.")

def test_loop(cfg):
    start_time = time.time()
    end_time = start_time + cfg["duration_sec"]
    
    entered_safe_zone = False 
    validation_started = False
    recorded_data = [] 
    
    logging.info(f"Test Start. Duration: {cfg['duration_sec']/60} min.")
    logging.info(f"Target: {cfg['min']}% ~ {cfg['max']}%")

    while time.time() < end_time:
        battery = get_battery_percentage()

        if battery is None:
            logging.warning("Battery read failed, retrying...")
            time.sleep(3)
            continue

        # --- 1. 準備階段 ---
        if not entered_safe_zone:
            if cfg["min"] <= battery <= cfg["max"]:
                entered_safe_zone = True
                logging.info(f"Battery ({battery}%) inside safe range. Ready.")
            else:
                logging.info(f"Adjusting... Current: {battery}%")

        # --- 2. 控制邏輯 ---
        if battery <= cfg["min"]:
            logging.info(f"Battery {battery}% <= Min, Charging...")
            # enable_charging 裡面已經包含 kill prime95 -> sleep 5 -> EC command
            enable_charging()
        elif battery >= cfg["max"]:
            logging.info(f"Battery {battery}% >= Max, Discharging...")
            # disable_charging 裡面已經包含 EC command -> sleep 5 -> run prime95
            disable_charging()
        else:
            if not validation_started:
                logging.debug(f"Battery {battery}% in range. Waiting...")

        # --- 3. 驗證觸發 ---
        if entered_safe_zone and not validation_started:
            if battery <= cfg["min"] or battery >= cfg["max"]:
                validation_started = True
                logging.info(f"=== Boundary Triggered ({battery}%), VALIDATION STARTED ===")

        # --- 4. 記錄數據 ---
        if validation_started:
            logging.info(f"[Record] Check: {battery}%")
            recorded_data.append(battery)
        
        # --- 5. 等待 ---
        time.sleep(cfg["interval"])

    analyze_result(recorded_data, cfg)

def analyze_result(data, cfg):
    logging.info("========== Test Finished. Analyzing Data ==========")
    
    if not data:
        msg = "No data recorded."
        logging.error(msg)
        write_result("FAIL", msg)
        return

    upper_limit = cfg["max"] + cfg["tolerance"]
    lower_limit = cfg["min"] - cfg["tolerance"]
    
    logging.info(f"Pass Criteria: {lower_limit}% <= Battery <= {upper_limit}%")
    
    violations = []
    for val in data:
        if val > upper_limit or val < lower_limit:
            violations.append(val)

    if not violations:
        logging.info("Success! All points within limits.")
        write_result("PASS", f"Tested {len(data)} points.")
    else:
        msg = f"FAIL. Found {len(violations)} violations: {violations}"
        logging.error(msg)
        write_result("FAIL", msg)

def main():
    remove_old_result()
    init_logger(BASE_DIR, prefix="Battery_Test")
    logging.info("========== Battery Charge/Discharge Test Start ==========")

    try:
        config = load_config()
        test_loop(config)

    except KeyboardInterrupt:
        logging.warning("Program interrupted by user (Ctrl+C).")
        write_result("FAIL", "User Interrupted")

    except Exception as e:
        logging.exception(f"Program Error: {e}")
        write_result("FAIL", str(e))

    finally:
        logging.info(">>> FINAL CLEANUP <<<")
        # 1. 確保 Prime95 被關閉
        kill_prime95()
        
        # 2. 確保切回充電模式
        try:
            logging.info("Restoring Auto Mode (Charge)...")
            # 這裡不呼叫 enable_charging() 因為不需要再 wait 5s，直接下指令即可
            if not run_command(CMD_AUTO_MODE, "Force Charge"):
                 logging.error("Failed to restore charge mode.")
        except Exception as e:
            logging.error(f"Cleanup error: {e}")

if __name__ == "__main__":
    main()