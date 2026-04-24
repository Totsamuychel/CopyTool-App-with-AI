# -*- coding: utf-8 -*-
import mss
import io
import base64
import logging
from PIL import Image

logger = logging.getLogger(__name__)

def capture_screen_area(bbox: tuple):
    """
    Captures an image of the specified screen area.
    bbox format: (left, top, right, bottom)
    """
    try:
        with mss.mss() as sct:
            monitor = {
                'left': bbox[0], 
                'top': bbox[1], 
                'width': bbox[2] - bbox[0], 
                'height': bbox[3] - bbox[1]
            }
            sct_img = sct.grab(monitor)
            # Convert to PIL Image
            img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
            return img
    except Exception as e:
        logger.error(f"Error capturing screen area: {e}")
        return None

def image_to_base64(image: Image.Image) -> str:
    """
    Converts a PIL Image to a base64 encoded string.
    """
    try:
        buffered = io.BytesIO()
        image.save(buffered, format="PNG")
        return base64.b64encode(buffered.getvalue()).decode('utf-8')
    except Exception as e:
        logger.error(f"Error converting image to base64: {e}")
        return None
