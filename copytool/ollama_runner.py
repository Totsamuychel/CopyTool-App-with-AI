# -*- coding: utf-8 -*-
import subprocess
import psutil
import logging

logger = logging.getLogger(__name__)

def is_ollama_running() -> bool:
    """Checks if 'ollama serve' is currently running on the system."""
    for p in psutil.process_iter(['name', 'cmdline']):
        try:
            if 'ollama' in p.info['name'].lower():
                if p.info['cmdline'] and any('serve' in arg for arg in p.info['cmdline']):
                    return True
        except Exception:
            continue
    return False

def start_ollama_server():
    """Starts the Ollama server in a new PowerShell window."""
    logger.info("Attempting to start 'ollama serve' in a new PowerShell window...")
    try:
        cmd = [
            "powershell",
            "-NoExit",
            "-Command", "echo 'Starting Ollama Server...'; ollama serve"
        ]
        # CREATE_NEW_CONSOLE = 0x00000010 (Windows specific flag)
        subprocess.Popen(cmd, creationflags=0x00000010)
        logger.info("Start command sent to PowerShell.")
        return True
    except Exception as e:
        logger.error(f"Failed to start Ollama server: {e}")
        return False
