import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading, sys
import pythoncom   # <--- important pour COM/Outlook
from traitement_npai import pipeline
from couts_et_graphique import main as analyse_main


# ======== Redirection des logs ========
class RedirectLogs:
    def __init__(self, widget):
        self.widget = widget
    def write(self, msg):
        self.widget.insert("end", msg)
        self.widget.see("end")
    def flush(self):
        pass


# ======== Fonction générique pour exécuter un traitement ========
def run_task(func, btn, progress, use_com=False):
    def wrapper():
        btn.config(state="disabled")
        progress.start()
        try:
            if use_com:
                pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            func()
        except Exception as e:
            messagebox.showerror("Erreur", str(e))
        finally:
            if use_com:
                pythoncom.CoUninitialize()
            btn.config(state="normal")
            progress.stop()
    threading.Thread(target=wrapper, daemon=True).start()


# ======== Interface principale ========
root = tk.Tk()
root.title("Outil RA NPAI")
root.geometry("700x500")

notebook = ttk.Notebook(root)
frame1 = ttk.Frame(notebook)
frame2 = ttk.Frame(notebook)
notebook.add(frame1, text="Pipeline NPAI")
notebook.add(frame2, text="Analyse Frais")
notebook.pack(expand=True, fill="both")

# --- Onglet 1 : Pipeline NPAI ---
btn_run1 = ttk.Button(frame1, text="Lancer pipeline",
                      command=lambda: run_task(lambda: pipeline(reconstruction_totale=True),
                                               btn_run1, progress1, use_com=True))
btn_run1.pack(pady=5)

progress1 = ttk.Progressbar(frame1, mode="indeterminate")
progress1.pack(fill="x", padx=10, pady=5)

log1 = scrolledtext.ScrolledText(frame1, wrap="word", height=15)
log1.pack(expand=True, fill="both", padx=10, pady=5)

# Rediriger stdout/stderr vers log1
sys.stdout = RedirectLogs(log1)
sys.stderr = RedirectLogs(log1)


# --- Onglet 2 : Analyse Frais ---
btn_run2 = ttk.Button(frame2, text="Lancer analyse frais",
                      command=lambda: run_task(analyse_main, btn_run2, progress2))
btn_run2.pack(pady=5)

progress2 = ttk.Progressbar(frame2, mode="indeterminate")
progress2.pack(fill="x", padx=10, pady=5)

log2 = scrolledtext.ScrolledText(frame2, wrap="word", height=15)
log2.pack(expand=True, fill="both", padx=10, pady=5)


root.mainloop()
