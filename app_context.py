import tkinter as tk
import os

class AppContext:
    def __init__(self):
        # 初始化 root
        self.root = tk.Tk()

        # 設定 icon
        icon_path = os.path.join(os.path.dirname(__file__), 'static', 'icon', '32.ico')
        if os.path.exists(icon_path):
            self.root.iconbitmap(os.path.abspath(icon_path))
        
        # 設定主視窗標題
        self.root.title("Outlook Template Assistant")
