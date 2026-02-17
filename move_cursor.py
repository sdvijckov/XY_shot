
# -*- coding: utf-8 -*-

# Импортируем все необходимые библиотеки
import logging
import os
import ctypes
import win32api
import win32con
import win32gui
from pynput import keyboard

# --- Блок 0: Настройка ---
# Определяем абсолютный путь к директории, где лежит скрипт
script_dir = os.path.dirname(os.path.abspath(__file__))

# Настройка логирования
log_file_path = os.path.join(script_dir, 'scroller.log')
logging.basicConfig(
    filename=log_file_path,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Настройка файла блокировки
LOCK_FILE_PATH = os.path.join(script_dir, 'scroller.lock')

logging.info("Скрипт запущен.")

# --- Блок 1: Защита от двойного запуска (Файл блокировки) ---

def show_error_messagebox():
    """Показывает информационное сообщение об уже запущенной копии."""
    title = "Ошибка запуска Scroller"
    message = (
        "Программа для прокрутки консоли уже запущена или была завершена некорректно.\n\n"
        "Что делать:\n"
        "1. Откройте Диспетчер задач (Ctrl+Shift+Esc).\n"
        "2. Найдите и завершите процесс 'pythonw.exe'.\n"
        "3. Если это не помогло, удалите вручную файл 'scroller.lock' в папке со скриптом.\n\n"
        "После этого попробуйте запустить скрипт снова."
    )
    # MB_ICONERROR = 0x10, MB_OK = 0x0, MB_TOPMOST = 0x40000
    ctypes.windll.user32.MessageBoxW(0, message, title, 0x10 | 0x0 | 0x40000)

if os.path.exists(LOCK_FILE_PATH):
    logging.warning("Обнаружен файл блокировки. Показываем сообщение и завершаем работу.")
    show_error_messagebox()
    exit(0)
else:
    try:
        # Создаем файл блокировки
        with open(LOCK_FILE_PATH, 'w') as f:
            f.write(str(os.getpid()))
        logging.info(f"Файл блокировки создан: {LOCK_FILE_PATH}")
    except Exception as e:
        logging.error(f"Не удалось создать файл блокировки: {e}")
        exit(1)


# --- Блок 2: Основная логика прокрутки ---

CONSOLE_CLASS_NAME = "ConsoleWindowClass"

def on_press(key):
    """Эта функция вызывается при каждом нажатии клавиши на клавиатуре."""
    try:
        hwnd = win32gui.GetForegroundWindow()
        class_name = win32gui.GetClassName(hwnd)
        
        logging.info(f"Нажата клавиша: {key}. Активное окно: HWND={hwnd}, Класс='{class_name}'")

        if class_name == CONSOLE_CLASS_NAME:
            scroll_action = None
            if key == keyboard.Key.page_up:
                scroll_action = "SB_PAGEUP"
                # Возвращаемся к PostMessage, чтобы избежать зависаний
                win32api.PostMessage(hwnd, win32con.WM_VSCROLL, win32con.SB_PAGEUP, 0)
            elif key == keyboard.Key.page_down:
                scroll_action = "SB_PAGEDOWN"
                win32api.PostMessage(hwnd, win32con.WM_VSCROLL, win32con.SB_PAGEDOWN, 0)
            
            if scroll_action:
                logging.info(f"Окно является консолью. Отправлено (PostMessage) сообщение прокрутки: {scroll_action}")
        else:
            logging.info("Активное окно не является консолью. Действие пропущено.")

    except AttributeError:
        logging.debug(f"Нажата обычная клавиша {key}, игнорируется.")
    except Exception as e:
        logging.error(f"Произошла ошибка в on_press: {e}", exc_info=True)


def on_release(key):
    """Эта функция вызывается, когда клавиша отпускается."""
    if key == keyboard.Key.esc:
        logging.info("Нажата клавиша ESC. Завершение работы слушателя.")
        return False


# --- Блок 3: Запуск и удержание процесса ---

try:
    logging.info("Запуск слушателя клавиатуры.")
    with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()
except Exception as e:
    logging.critical(f"Произошла критическая ошибка в работе слушателя клавиатуры: {e}", exc_info=True)
finally:
    logging.info("Скрипт завершает работу. Удаление файла блокировки.")
    if os.path.exists(LOCK_FILE_PATH):
        try:
            os.remove(LOCK_FILE_PATH)
            logging.info("Файл блокировки успешно удален.")
        except Exception as e:
            logging.error(f"Не удалось удалить файл блокировки: {e}")
    logging.info("Программа полностью завершена.")
