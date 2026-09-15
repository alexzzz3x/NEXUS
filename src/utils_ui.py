import customtkinter as ctk
import tkinter as tk
import threading
import winsound

class ToastNotification:
    def __init__(self, parent, message, duration=3000, color="green"):
        self.toast = ctk.CTkToplevel(parent)
        self.toast.overrideredirect(True)
        self.toast.attributes("-alpha", 0.0)
        self.toast.attributes("-topmost", True)
        
        # Posicionamento no canto inferior direito
        screen_width = parent.winfo_screenwidth()
        screen_height = parent.winfo_screenheight()
        width = 320
        height = 80
        x = screen_width - width - 20
        y = screen_height - height - 60
        self.toast.geometry(f"{width}x{height}+{x}+{y}")
        
        bg_color = "#2E7D32" if color == "green" else "#C62828" if color == "red" else "#1565C0"
        self.toast.configure(fg_color=bg_color)
        
        lbl = ctk.CTkLabel(self.toast, text=message, text_color="white", font=("Segoe UI", 14, "bold"), wraplength=300)
        lbl.pack(expand=True, fill="both", padx=10, pady=10)
        
        self.duration = duration
        self.fade_in()
        
    def fade_in(self):
        try:
            alpha = self.toast.attributes("-alpha")
            if alpha < 0.9:
                self.toast.attributes("-alpha", alpha + 0.1)
                self.toast.after(20, self.fade_in)
            else:
                self.toast.after(self.duration, self.fade_out)
        except Exception:
            pass
            
    def fade_out(self):
        try:
            alpha = self.toast.attributes("-alpha")
            if alpha > 0.05:
                self.toast.attributes("-alpha", alpha - 0.1)
                self.toast.after(20, self.fade_out)
            else:
                self.toast.destroy()
        except Exception:
            pass

def show_toast(parent, message, duration=3000, type="info"):
    """
    Exibe um Toast Notification animado.
    types permitidos: "success", "error", "info"
    """
    color = "green" if type == "success" else "red" if type == "error" else "blue"
    ToastNotification(parent, message, duration, color)

def play_sound(type="success"):
    """
    Toca um som nativo do Windows em background.
    """
    def play():
        try:
            if type == "success":
                winsound.MessageBeep(winsound.MB_OK)
            elif type == "error":
                winsound.MessageBeep(winsound.MB_ICONHAND)
            elif type == "done":
                winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
        except:
            pass
    threading.Thread(target=play, daemon=True).start()

def run_async(func, *args, **kwargs):
    """
    Executa uma funcao em uma thread separada para não congelar a UI.
    """
    threading.Thread(target=func, args=args, kwargs=kwargs, daemon=True).start()
