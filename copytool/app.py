# -*- coding: utf-8 -*-
import os
import sys
import time
import threading
import logging
import keyboard
import pyperclip
from PIL import Image
from pystray import MenuItem as item, Icon as icon, Menu as menu

import win32gui
import win32con
import win32console

from copytool.config import HOTKEY, ICON_FILE
from copytool.ui import ScreenSelector
from copytool.capture import capture_screen_area, image_to_base64
from copytool.ollama_api import setup_api_url, extract_text_from_image, shutdown_ngrok
from copytool.ollama_runner import is_ollama_running, start_ollama_server

logger = logging.getLogger(__name__)

class CopyToolApp:
    def __init__(self):
        self.is_workflow_running = False
        self.last_trigger_time = 0
        self.console_hwnd = self._get_console_hwnd()
        self.tray_icon = None

    def _get_console_hwnd(self):
        try:
            hwnd = win32console.GetConsoleWindow()
            if not hwnd:
                hwnd = win32gui.GetForegroundWindow()
            return hwnd
        except Exception as e:
            logger.error(f"Failed to get console handle: {e}")
            return None

    def _workflow_logic(self):
        """Core workflow: selection -> capture -> OCR -> clipboard."""
        logger.info("Opening selection window...")
        try:
            selector = ScreenSelector()
            selection_box = selector.get_selection()

            if not selection_box:
                logger.info("Selection cancelled or area too small.")
                return

            logger.info(f"Area selected: {selection_box}")
            
            # 1. Capture Image
            captured_image = capture_screen_area(selection_box)
            if not captured_image:
                logger.error("Failed to capture image.")
                return
            
            # 2. Convert to Base64
            b64_image = image_to_base64(captured_image)
            if not b64_image:
                logger.error("Failed to encode image to base64.")
                return
            
            # 3. Request OCR from Ollama
            recognized_text = extract_text_from_image(b64_image)
            
            # 4. Copy to clipboard
            if recognized_text and not recognized_text.startswith("Error"):
                pyperclip.copy(recognized_text)
                logger.info("\\n--- Recognized Text Copied to Clipboard: ---\\n")
                logger.info(recognized_text)
                logger.info("\\n----------------------------------------------\\n")
            else:
                logger.error(f"Failed to extract text: {recognized_text}")

        except Exception as e:
            logger.error(f"Error in workflow: {e}")

    def run_workflow(self):
        """Thread wrapper for the main workflow."""
        current_time = time.time()
        
        # Debounce protection (1 second)
        if current_time - self.last_trigger_time < 1.0:
            return
        
        if self.is_workflow_running:
            logger.warning("Workflow is already running, ignoring keypress.")
            return
        
        self.last_trigger_time = current_time
        self.is_workflow_running = True
        logger.info(f"Hotkey '{HOTKEY}' pressed. Starting workflow...")
        
        def thread_target():
            try:
                self._workflow_logic()
            finally:
                self.is_workflow_running = False
                logger.info("Workflow completed. Ready for next capture.")

        threading.Thread(target=thread_target, daemon=True).start()

    def toggle_console(self, icon, item):
        """Toggles the visibility of the console window."""
        if not self.console_hwnd:
            logger.error("Console handle not found.")
            return
            
        try:
            is_visible = win32gui.IsWindowVisible(self.console_hwnd)
            if is_visible:
                win32gui.ShowWindow(self.console_hwnd, win32con.SW_HIDE)
                logger.info("Console hidden.")
            else:
                win32gui.ShowWindow(self.console_hwnd, win32con.SW_SHOW)
                win32gui.SetForegroundWindow(self.console_hwnd)
                logger.info("Console shown.")
        except Exception as e:
            logger.error(f"Error toggling console: {e}")

    def check_and_start_ollama(self, icon, item):
        """Checks if Ollama is running, starts it if not."""
        if is_ollama_running():
            logger.info("Ollama is already running!")
        else:
            logger.info("Starting Ollama server...")
            start_ollama_server()

    def exit_app(self, icon, item):
        """Clean shutdown procedure."""
        logger.info("Exiting application...")
        shutdown_ngrok()
        try:
            keyboard.unhook_all()
        except:
            pass
        
        if self.tray_icon:
            self.tray_icon.stop()
        os._exit(0)

    def start(self):
        """Initializes and runs the application."""
        logger.info("="*50)
        logger.info("Starting CopyTool AI App...")
        
        # Setup API URL (Local or Ngrok)
        setup_api_url()

        # Register Hotkey
        try:
            keyboard.add_hotkey(HOTKEY, self.run_workflow, suppress=False)
            logger.info(f"Hotkey '{HOTKEY}' registered successfully.")
        except Exception as e:
            logger.error(f"Failed to register hotkey: {e}")
            return

        # Setup System Tray
        try:
            image = Image.open(ICON_FILE)
        except FileNotFoundError:
            logger.warning(f"Icon file not found at {ICON_FILE}. Using default blue square.")
            image = Image.new('RGB', (64, 64), color='blue')

        tray_menu = menu(
            item('Распознать текст (Ctrl+Shift+X)', lambda _: self.run_workflow()),
            item('Запустить сервер Ollama', self.check_and_start_ollama),
            item('Показать/Скрыть консоль', self.toggle_console),
            item('Выход', self.exit_app)
        )
        
        self.tray_icon = icon('CopyTool', image, 'CopyTool AI', tray_menu)

        logger.info("Application is running!")
        logger.info(f"Press '{HOTKEY}' to select an area for OCR.")
        logger.info("="*50)

        try:
            self.tray_icon.run()
        except KeyboardInterrupt:
            logger.info("Received interrupt signal. Shutting down...")
        finally:
            self.exit_app(None, None)
