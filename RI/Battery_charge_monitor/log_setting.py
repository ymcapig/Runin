from pathlib import Path
import logging
from datetime import datetime
import sys

def init_logger(base_dir: Path, prefix="log"):

    # === build log dir ===
    log_dir = base_dir / "log"
    log_dir.mkdir(exist_ok=True)

    # === build log filename ===
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = log_dir / f"{prefix}_{timestamp}.log"

    file_handler = logging.FileHandler(str(log_path), encoding='utf-8')
    stream_handler = logging.StreamHandler(sys.stdout)

    # === configure logging === 
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s,%(msecs)03d [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[file_handler, stream_handler],
        force=True
    )

    logging.debug(f"===== {prefix} =====")
    logging.info(f"Log file created at: {log_path}")
    
    return log_path