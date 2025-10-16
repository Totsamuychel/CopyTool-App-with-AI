import tkinter as tk
from turtle import setup
import sys
import subprocess
import mss
import pyperclip
import requests
import base64
import io
import keyboard
import logging
import os
import threading
import win32gui
import win32con
import win32console
from PIL import Image
from pathlib import Path
from pystray import MenuItem as item, Icon as icon
from pyngrok import ngrok
import time

# --- НАСТРОЙКА ЛОГИРОВАНИЯ ---
log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app_log.log')

# Настраиваем логгер
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# --- КОНФИГУРАЦИЯ ---
OLLAMA_MODEL = 'qwen2.5vl:7b'
HOTKEY = 'ctrl+shift+x'
OLLAMA_API_URL = None
PROMPT = "Прочитай весь текст на этом изображении. Выведи только сам текст, без каких-либо объяснений или комментариев."
ICON_FILE = Path(__file__).resolve().parent / "icon.png"
NGROK_AUTHTOKEN = '31njXTz7EXHwlyOTDuCeapmoWLb_49DJ7r82hwFfoV4ESZ3PS'

# Глобальные переменные
tray_icon = None
console_hwnd = None
public_url = None
is_workflow_running = False
last_trigger_time = 0
hotkey_registered = False

class ScreenSelector:
    """Класс для создания окна выбора области на экране."""
    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes('-alpha', 0.3)
        self.root.attributes('-topmost', True)
        self.root.geometry(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}+0+0")
        self.root.configure(bg='black')

        self.canvas = tk.Canvas(self.root, cursor="cross", bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.start_x = None
        self.start_y = None
        self.rect = None
        self.selection_box = None

        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        
        self.root.bind('<Escape>', self.on_escape)
        self.root.focus_set()

    def on_escape(self, event):
        """Обработчик нажатия Escape для отмены выделения."""
        logging.info("Выделение отменено пользователем (Escape)")
        self.selection_box = None
        self.root.quit()

    def on_button_press(self, event):
        """Обработчик нажатия кнопки мыши."""
        self.start_x = event.x
        self.start_y = event.y
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y, 
            outline='red', width=2
        )

    def on_mouse_drag(self, event):
        """Обработчик движения мыши с зажатой кнопкой."""
        self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)

    def on_button_release(self, event):
        """Обработчик отпускания кнопки мыши."""
        end_x = event.x
        end_y = event.y

        x1 = min(self.start_x, end_x)
        y1 = min(self.start_y, end_y)
        x2 = max(self.start_x, end_x)
        y2 = max(self.start_y, end_y)
        
        if x2 - x1 > 10 and y2 - y1 > 10:
            self.selection_box = (int(x1), int(y1), int(x2), int(y2))
        
        self.root.quit()

    def get_selection(self):
        """Запускает окно и возвращает координаты выделения."""
        try:
            self.root.mainloop()
            return self.selection_box
        except Exception as e:
            logging.error(f"Ошибка в get_selection: {e}")
            return None
        finally:
            try:
                self.root.destroy()
            except:
                pass

def capture_screen_area(bbox):
    """Захватывает изображение указанной области экрана."""
    try:
        with mss.mss() as sct:
            monitor = {
                'top': bbox[1], 
                'left': bbox[0], 
                'width': bbox[2] - bbox[0], 
                'height': bbox[3] - bbox[1]
            }
            sct_img = sct.grab(monitor)
            img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
            return img
    except Exception as e:
        logging.error(f"Ошибка при захвате экрана: {e}")
        return None

def image_to_base64(image):
    """Конвертирует изображение Pillow в строку base64."""
    try:
        buffered = io.BytesIO()
        image.save(buffered, format="PNG")
        return base64.b64encode(buffered.getvalue()).decode('utf-8')
    except Exception as e:
        logging.error(f"Ошибка при конвертации в base64: {e}")
        return None

def get_text_from_ollama(image_base64):
    """Отправляет изображение в Ollama и возвращает распознанный текст."""
    logging.info("Отправка запроса в Ollama... Ожидайте.")
    try:
        payload = {
            "model": OLLAMA_MODEL,
            "prompt": PROMPT,
            "images": [image_base64],
            "stream": False
        }
        response = requests.post(OLLAMA_API_URL, json=payload, timeout=300)
        response.raise_for_status()
        
        response_data = response.json()
        text = response_data.get('response', '').strip()
        return text

    except requests.exceptions.RequestException as e:
        logging.error(f"Ошибка при обращении к Ollama: {e}")
        return f"Ошибка при обращении к Ollama: {e}"
    except Exception as e:
        logging.error(f"Произошла неизвестная ошибка: {e}")
        return f"Произошла неизвестная ошибка: {e}"

def main_workflow():
    """Основной процесс: выделение, распознавание, копирование."""
    logging.info("Запуск окна выделения области...")
    
    try:
        selector = ScreenSelector()
        selection_box = selector.get_selection()

        if selection_box:
            logging.info(f"Выделена область: {selection_box}")
            
            # Захватываем изображение
            captured_image = capture_screen_area(selection_box)
            if not captured_image:
                logging.error("Не удалось захватить изображение")
                return
            
            # Конвертируем в base64
            b64_image = image_to_base64(captured_image)
            if not b64_image:
                logging.error("Не удалось конвертировать изображение в base64")
                return
            
            # Отправляем в Ollama для распознавания
            recognized_text = get_text_from_ollama(b64_image)
            
            if recognized_text and not recognized_text.startswith("Ошибка"):
                # Копируем в буфер обмена
                pyperclip.copy(recognized_text)
                logging.info("\n--- Распознанный текст скопирован в буфер обмена: ---")
                logging.info(recognized_text)
                logging.info("---------------------------------------------------\n")
            else:
                logging.error(f"Не удалось распознать текст: {recognized_text}")
        else:
            logging.info("Область не была выделена или операция отменена.")
            
    except Exception as e:
        logging.error(f"Ошибка в main_workflow: {e}")

def hotkey_callback():
    """Callback функция для горячей клавиши."""
    global is_workflow_running, last_trigger_time
    
    current_time = time.time()
    
    # Защита от дребезга
    if current_time - last_trigger_time < 1.0:
        logging.info("Слишком быстрое повторное нажатие, игнорируется.")
        return
    
    if is_workflow_running:
        logging.warning("Процесс уже запущен, повторное нажатие игнорируется.")
        return
    
    last_trigger_time = current_time
    is_workflow_running = True
    
    logging.info(f"Горячая клавиша '{HOTKEY}' нажата. Запуск процесса...")
    
    # Запускаем в отдельном потоке
    workflow_thread = threading.Thread(target=run_workflow_with_cleanup, daemon=True)
    workflow_thread.start()

def run_workflow_with_cleanup():
    """Запускает workflow и гарантированно снимает блокировку."""
    global is_workflow_running
    try:
        main_workflow()
    except Exception as e:
        logging.error(f"Ошибка в workflow: {e}")
    finally:
        is_workflow_running = False
        logging.info("Процесс завершен, блокировка снята. Готов к следующему нажатию.")

def setup_ngrok():
    """Настройка ngrok туннеля."""
    global OLLAMA_API_URL, public_url
    
    try:
        ngrok.set_auth_token(NGROK_AUTHTOKEN)
        http_tunnel = ngrok.connect(11434, "http")
        public_url = http_tunnel.public_url
        OLLAMA_API_URL = f"{public_url}/api/generate"
        logging.info(f"Туннель ngrok запущен. URL: {public_url}")
        return True
    except Exception as e:
        logging.error(f"Не удалось запустить ngrok. Ошибка: {e}")
        return False

def register_hotkey():
    """Регистрация горячей клавиши."""
    global hotkey_registered
    
    try:
        if hotkey_registered:
            keyboard.unhook_all()
            hotkey_registered = False
        
        keyboard.add_hotkey(HOTKEY, hotkey_callback, suppress=False)
        hotkey_registered = True
        logging.info(f"Горячая клавиша '{HOTKEY}' зарегистрирована успешно.")
        return True
    except Exception as e:
        logging.error(f"Ошибка при регистрации горячей клавиши: {e}")
        return False

def exit_action(icon, item):
    """Функция для выхода из приложения."""
    logging.info("Выход из приложения...")
    
    try:
        if public_url:
            ngrok.disconnect(public_url)
        ngrok.kill()
    except:
        pass
    
    try:
        keyboard.unhook_all()
    except:
        pass
    
    icon.stop()
    os._exit(0)

def get_console_window():
    try:
        console_hwnd = win32console.GetConsoleWindow()
        logging.info(f"Handle консоли: {console_hwnd}")
        return console_hwnd
    except Exception as e:
        logging.error(f"Не удалось получить handle консоли: {e}")
        return None

def toggle_console_visibility(icon, item):
    """Переключает видимость консоли."""
    global console_hwnd
    
    console_hwnd = get_console_window()
    """if console_hwnd:
        try:
            if win32gui.IsWindowVisible(console_hwnd):
                win32gui.ShowWindow(console_hwnd, win32con.SW_HIDE)
                logging.info("Консоль скрыта.")
            else:
                win32gui.ShowWindow(console_hwnd, win32con.SW_SHOW)
                win32gui.SetForegroundWindow(console_hwnd)
                logging.info("Консоль показана.")
        except Exception as e:
            logging.error(f"Ошибка при переключении консоли: {e}")"""
    if console_hwnd:
        try:
            is_visible = win32gui.IsWindowVisible(console_hwnd)
            logging.info(f"Текущее состояние консоли: {'видима' if is_visible else 'скрыта'}")
            
            if is_visible:
                win32gui.ShowWindow(console_hwnd, win32con.SW_HIDE)
                logging.info("Консоль скрыта.")
            else:
                win32gui.ShowWindow(console_hwnd, win32con.SW_SHOW)
                win32gui.SetForegroundWindow(console_hwnd)
                logging.info("Консоль показана.")
        except Exception as e:
            logging.error(f"Ошибка при переключении консоли: {e}")
    else:
        logging.error("Handle консоли не найден.")


def main():
    global tray_icon, console_hwnd

    logging.info("="*50)
    logging.info("Запуск приложения...")
    
    # Настройка ngrok
    if not setup_ngrok():
        logging.error("Не удалось настроить ngrok. Завершение работы.")
        return

    # Получаем handle консоли
    try:
        console_hwnd = win32gui.GetForegroundWindow()
    except Exception as e:
        logging.error(f"Не удалось получить handle консоли: {e}")

    # Регистрируем горячую клавишу
    if not register_hotkey():
        logging.error("Не удалось зарегистрировать горячую клавишу. Завершение работы.")
        return

    # Настройка иконки трея
    try:
        image = Image.open(ICON_FILE)
    except FileNotFoundError:
        logging.warning(f"Файл иконки {ICON_FILE} не найден. Используется стандартная иконка.")
        image = Image.new('RGB', (64, 64), color='blue')

    menu = (
        item('Показать/скрыть консоль', toggle_console_visibility),
        item('Выход', exit_action)
    )
    
    tray_icon = icon('CopyTextTool', image, 'Инструмент для копирования текста', menu)

    logging.info("Приложение успешно запущено!")
    logging.info(f"Нажмите хоткей '{HOTKEY}' для выбора области экрана.")
    logging.info("Нажмите Escape в окне выделения для отмены.")
    logging.info("="*50)

    # Запускаем системный трей (блокирующий вызов)
    try:
        tray_icon.run()
    except KeyboardInterrupt:
        logging.info("Получен сигнал прерывания. Завершение работы...")
    except Exception as e:
        logging.error(f"Ошибка в главном цикле: {e}")
    finally:
        exit_action(None, None)

if __name__ == "__main__":
    main()