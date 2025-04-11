import tkinter as tk
import time
from datetime import datetime
from constantes import *

def write_log(message):
    log_file = LOG_FILE
    with open(log_file, 'a') as f:
        f.write(f"{datetime.now()} - {message}\n")


def log_and_display(message, text_box, root, delay=0):
    if delay:
        time.sleep(delay)
    text_box.insert(tk.END, message + "\n")
    root.update()
    write_log(message)