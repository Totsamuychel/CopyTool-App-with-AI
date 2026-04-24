# -*- coding: utf-8 -*-
import os
from pathlib import Path
from dotenv import load_dotenv

# Load variables from .env file if it exists
load_dotenv()

# Model to use for text extraction
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "qwen2.5vl:7b")

# System prompt for Ollama
OLLAMA_PROMPT = os.getenv("OLLAMA_PROMPT", "Прочитай весь текст на этом изображении. Выведи только сам текст, без каких-либо объяснений или комментариев.")

# Ollama API URL (Local by default)
OLLAMA_LOCAL_URL = "http://localhost:11434/api/generate"

# Ngrok Configuration (Optional)
USE_NGROK = os.getenv("USE_NGROK", "False").lower() in ('true', '1', 't')
NGROK_AUTHTOKEN = os.getenv("NGROK_AUTHTOKEN", "")

# Hotkey for triggering the selection
HOTKEY = os.getenv("HOTKEY", "ctrl+shift+x")

# Application Paths
BASE_DIR = Path(__file__).resolve().parent.parent
LOG_FILE = BASE_DIR / 'app_log.log'
ICON_FILE = BASE_DIR / 'icon.png'
