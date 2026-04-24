# -*- coding: utf-8 -*-
import requests
import logging
from pyngrok import ngrok

from copytool.config import OLLAMA_MODEL, OLLAMA_PROMPT, OLLAMA_LOCAL_URL, NGROK_AUTHTOKEN, USE_NGROK

logger = logging.getLogger(__name__)

# Global variable to store the dynamic Ngrok URL if used
_ACTIVE_API_URL = OLLAMA_LOCAL_URL
_PUBLIC_NGROK_URL = None

def setup_api_url():
    """Sets up the Ollama API URL. Starts Ngrok if configured to do so."""
    global _ACTIVE_API_URL, _PUBLIC_NGROK_URL
    
    if USE_NGROK and NGROK_AUTHTOKEN:
        try:
            ngrok.set_auth_token(NGROK_AUTHTOKEN)
            http_tunnel = ngrok.connect(11434, "http")
            _PUBLIC_NGROK_URL = http_tunnel.public_url
            _ACTIVE_API_URL = f"{_PUBLIC_NGROK_URL}/api/generate"
            logger.info(f"Ngrok tunnel established. API URL: {_ACTIVE_API_URL}")
            return True
        except Exception as e:
            logger.error(f"Failed to start ngrok: {e}. Falling back to local URL.")
            _ACTIVE_API_URL = OLLAMA_LOCAL_URL
            return False
    else:
        _ACTIVE_API_URL = OLLAMA_LOCAL_URL
        logger.info(f"Using local API URL: {_ACTIVE_API_URL}")
        return True

def shutdown_ngrok():
    """Shuts down the ngrok tunnel if it is active."""
    if _PUBLIC_NGROK_URL:
        try:
            ngrok.disconnect(_PUBLIC_NGROK_URL)
            ngrok.kill()
            logger.info("Ngrok tunnel closed.")
        except Exception as e:
            logger.error(f"Error shutting down ngrok: {e}")

def extract_text_from_image(image_base64: str) -> str:
    """Sends the image to Ollama and returns the recognized text."""
    logger.info("Sending request to Ollama... Please wait.")
    try:
        payload = {
            "model": OLLAMA_MODEL,
            "prompt": OLLAMA_PROMPT,
            "images": [image_base64],
            "stream": False
        }
        response = requests.post(_ACTIVE_API_URL, json=payload, timeout=300)
        response.raise_for_status()
        
        response_data = response.json()
        text = response_data.get('response', '').strip()
        return text

    except requests.exceptions.RequestException as e:
        logger.error(f"Network error when calling Ollama: {e}")
        return f"Error: Network issue - {e}"
    except Exception as e:
        logger.error(f"Unknown error during Ollama request: {e}")
        return f"Error: {e}"
