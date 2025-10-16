import os
import sys
import subprocess
import time
import psutil
from pystray import Icon as TrayIcon, Menu as Menu, MenuItem as Item
from PIL import Image, ImageDraw
import threading
import win32console
import win32gui
import win32con

def log(msg):
    print(f"[OllamaAutoStart] {msg}")
    sys.stdout.flush()

def is_ollama_running():
    """Проверяет, запущен ли ollama serve"""
    for p in psutil.process_iter(['name', 'cmdline']):
        try:
            if 'ollama' in p.info['name'].lower():
                if any('serve' in arg for arg in p.info['cmdline']):
                    return True
        except Exception:
            continue
    return False

def start_ollama_powershell():
    """Запускает ollama serve в отдельном PowerShell окне"""
    log("Пытаюсь запустить ollama serve в новом PowerShell-окне...")
    cmd = [
        "powershell",
        "-NoExit",
        "-Command", "echo Запуск Ollama...; ollama serve"
    ]
    subprocess.Popen(cmd, creationflags=subprocess.CREATE_NEW_CONSOLE)
    log("Команда отправлена в PowerShell (отдельное окно).")

def get_console_hwnd():
    try:
        return win32console.GetConsoleWindow()
    except Exception:
        return None

def show_hide_console(show=True):
    """Показывает или скрывает текущее консольное окно"""
    hwnd = get_console_hwnd()
    if hwnd:
        if show:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.SetForegroundWindow(hwnd)
        else:
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)

def create_image():
    img = Image.new('RGB', (64, 64), color='white')
    d = ImageDraw.Draw(img)
    d.rectangle((16, 16, 48, 48), fill='blue', outline='black')
    d.text((20, 20), "O", fill=(255, 255, 255))
    return img

class TrayApp:
    def __init__(self):
        self.console_visible = True # По умолчанию консоль показана
        self.icon = TrayIcon(
            "OllamaAutoStart",
            create_image(),
            "Ollama Server Starter",
            menu=Menu(
                Item(
                    lambda item: "Скрыть консоль" if self.console_visible else "Показать консоль",
                    self.toggle_console_visibility
                ),
                Item("Выйти", self.quit_app)
            )
        )

    def toggle_console_visibility(self, icon, item):
        self.console_visible = not self.console_visible
        show_hide_console(show=self.console_visible)
        log("Консоль скрыта" if not self.console_visible else "Консоль показана")

    def quit_app(self, icon, item):
        log("Выход по запросу пользователя.")
        icon.stop()
        sys.exit(0)

    def run(self):
        self.icon.run()

def main():
    log("=== Ollama Автостартер запущен ===")
    # Проверим, запущен ли Ollama
    if is_ollama_running():
        log("Ollama уже запущен! Повторный запуск не требуется.")
    else:
        log("Ollama не найден, пробую запустить...")
        start_ollama_powershell()
        time.sleep(2)
    log("Теперь Ollama должен быть запущен (проверьте PowerShell-окно).")

    # Трей-иконка с управлением консолью
    app = TrayApp()
    log("Добавлена иконка в трей (управление: ППКМ → скрыть/показать/выйти).")
    app.run()

if __name__ == "__main__":
    t = threading.Thread(target=main)
    t.start()
