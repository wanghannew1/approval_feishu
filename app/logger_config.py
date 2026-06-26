import os
import sys
import logging
from pathlib import Path
from datetime import datetime

# Configure logging
LOG_DIR = Path(__file__).parent.parent / "log"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "app.log"

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8', mode='a'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger('FeishuApproval')