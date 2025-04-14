# # # # # # # # # # # #
# But : Ce fichier contient l'ensemble des méthodes utilitaires autres
# Par : Estéban DESESSARD - e.desessard@burographic.fr
# Date : 11/04/2025
# # # # # # # # # # # #

import tkinter as tk
import time
from datetime import datetime
from constantes import *

# But : Écrire un message dans le fichier de log
def write_log(message):
    log_file = LOG_FILE
    with open(log_file, 'a') as f:
        f.write(f"{datetime.now()} - {message}\n")

# But : Écrire un message sur l'interface utilisateur et dans le fichier de log
def log_and_display(message, text_box, root, delay=0):
    if delay:
        time.sleep(delay)
    text_box.insert(tk.END, message + "\n")
    root.update()
    write_log(message)

def generate_report(datas):
    return datas