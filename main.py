# -*- coding: utf-8 -*-
"""
Entry point for the CopyTool AI Application.
"""
import logging
from copytool.config import LOG_FILE
from copytool.app import CopyToolApp

# --- LOGGING CONFIGURATION ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

if __name__ == '__main__':
    app = CopyToolApp()
    app.start()
