import tkinter as tk
import sys
import os
import time
from tkinter import messagebox
from pathlib import Path


# 确保模块可以正确导入
current_dir = Path(__file__).parent.absolute()
sys.path.append(str(current_dir))

# 导入自定义模块
from db_manager import DatabaseManager
from language_manager import LanguageManager
from template_manager import TemplateManager
from gui.main_window import MainWindow
import tkinter as tk

def main():
    """应用程序入口点"""
    print("初始化数据库...")

    
    # 创建数据库管理器
    db_manager = DatabaseManager()
    
    # 创建语言管理器
    language_manager = LanguageManager(db_manager=db_manager)
    # 加载语言偏好设置
    preferred_language = language_manager.load_user_preference()
    language_manager.set_language(preferred_language)
    
    # 创建模板管理器
    template_manager = TemplateManager(db_manager=db_manager)
    
    print("初始化应用程序...")
    # 创建主窗口
    root = tk.Tk()
    
    #设置应用程序图标（如果有的话）
    icon_path = current_dir / "static" / "icon" / "32.ico"
    if icon_path.exists():
            root.iconbitmap(icon_path)
    
    # 创建主窗口应用，传入语言管理器和模板管理器
    app = MainWindow(root, language_manager, template_manager)
    
    print("应用程序启动完成！")
    # 提示用户数据库位置和备份信息
    db_path = os.path.abspath(db_manager.db_file)
    print(f"数据库文件位置: {db_path}")
    print("请定期备份此文件以防数据丢失")

    # 显示备份提示对话框
    def show_backup_reminder():
        # 检查是否是首次运行或者上次提示已经过去了7天
        last_reminder = db_manager.get_setting('last_backup_reminder')
        current_time = str(int(time.time()))
        
        if not last_reminder or (int(current_time) - int(last_reminder)) > 7 * 24 * 60 * 60:  # 7天
            db_manager.save_setting('last_backup_reminder', current_time)
            messagebox.showinfo(
                language_manager.get_text("backup_reminder"), 
                language_manager.get_text("backup_message").format(db_path=db_path)
            )

    # 设置一个延迟，确保主窗口已经完全加载
    root.after(2000, show_backup_reminder)

    # 启动主循环
    root.mainloop()
    
    # 关闭数据库连接
    db_manager.close_connection()

if __name__ == "__main__":
    main()