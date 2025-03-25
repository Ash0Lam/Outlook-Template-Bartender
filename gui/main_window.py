import tkinter as tk
from tkinter import ttk, messagebox, StringVar, simpledialog
from typing import Dict, List, Any, Optional
import os
import sys
import re
import webbrowser
import threading
import time
import queue
import sqlite3
import importlib.util
from tkinter import ttk
from email_generator import EmailGenerator
from gui.edit_template import EditTemplateWindow


db_queue = queue.Queue()


class DBWorker(threading.Thread):
    def __init__(self, db_queue):
        threading.Thread.__init__(self)
        self.db_queue = db_queue
        self.daemon = True
        self.start()


    def run(self):
        while True:
            task = self.db_queue.get()  # Wait for the main thread's task
            if task is None:  # If a stop signal is received
                break
            event_type = task['event_type']
            # Perform database query operation
            templates = self.get_templates_for_event(event_type)
            task['callback'](templates)  # Execute callback function with query results


    def get_templates_for_event(self, event_type):
        # Use a new connection per thread
        conn = sqlite3.connect('your_database.db', check_same_thread=False)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM templates WHERE event_type=?", (event_type,))
        templates = cursor.fetchall()
        conn.close()
        return templates
    
class MainWindow:
    """主窗口類，處理主界面的顯示和用戶交互"""
    
    def __init__(self, root, language_manager=None, template_manager=None):
        """初始化主窗口"""
        self.root = root
        self.language_manager = language_manager
        self._ = language_manager.get_text if language_manager else lambda x: x
        self.outlook_checker = EmailGenerator()
        self.accounts = self.outlook_checker.get_outlook_accounts() if self.outlook_checker else []
        self.variable_values = {}
        
        # 初始化其他設置
        self.search_timer = None
        self.cached_template_content = {}
        
        # 確保 db_worker 線程只啟動一次
        if not hasattr(self, 'db_worker') or not self.db_worker.is_alive():
            self.db_worker = DBWorker(db_queue)
        
        self.root.title("Outlook Template Bartender")
        
        # 使用Grid布局管理器進行整體布局
        self.root.grid_rowconfigure(0, weight=1)  # 主內容區域可以擴展
        self.root.grid_rowconfigure(1, weight=0)  # 按鈕區域固定高度
        self.root.grid_columnconfigure(0, weight=1)  # 允許水平擴展
        
        # 設置適當的最小視窗尺寸
        self.root.minsize(width=775, height=775)
        
        # 初始化模板管理器和电子邮件生成器
        self.template_manager = template_manager if template_manager else TemplateManager()
        self.email_generator = EmailGenerator()
        
        # 事件类型和模板选择变量
        self.selected_event_type = StringVar()
        self.selected_template = StringVar()
        
        # 变量输入框引用
        self.var_entries = {}
        
        # 搜索变量
        self.search_var = StringVar()
        
        # 创建界面元素
        self._create_widgets()
        
        # 初始化事件类型下拉菜单
        self._update_event_types()
        
        # 在所有元素創建完成後居中窗口
        self.root.update_idletasks()
        self._center_window()
        self.root.deiconify()

    def _center_window(self):
        """將窗口置於螢幕中央"""
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        
        # 檢查是否已獲取有效大小
        if width <= 1 or height <= 1:  # 如果尚未正確渲染，使用默認大小
            width = 775
            height = 975
        
        # 獲取屏幕大小
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 計算位置
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        # 確保座標為正數
        x = max(0, x)
        y = max(0, y)
        
        # 設置窗口位置
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
        print(f"Window centered: {width}x{height} at position +{x}+{y}")

    
    def _create_widgets(self):
        """創建窗口小部件"""
        # 使用Grid布局替換Pack布局，讓底部按鈕固定
        self.root.grid_rowconfigure(0, weight=1)  # 主內容區域可以擴展
        self.root.grid_rowconfigure(1, weight=0)  # 按鈕區域固定高度
        self.root.grid_rowconfigure(2, weight=0)  # 狀態欄固定高度
        self.root.grid_columnconfigure(0, weight=1)  # 允許水平擴展
        
        # 創建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        version = self.template_manager.db_manager.get_app_info("version", self.language_manager.current_language)
        
        # 頂部菜單欄
        self._create_menu()
        
        # 標題標籤
        title_label = ttk.Label(main_frame, text=self._("app_title"), font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        title_label.language_key = "app_title" 

        conn = sqlite3.connect('data/app.db')
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM app_info WHERE key = 'version'")

        # 在標題標籤後添加版本和關於按鈕
        header_frame = ttk.Frame(main_frame)
        
        # 在右側添加版本號和關於按鈕
        version_frame = ttk.Frame(header_frame)
        
        # 版本號標籤
        version_label = ttk.Label(version_frame, text=f"Version: {version}", font=("Arial", 8))
        version_label.pack(pady=5)
        
        # 添加分隔線
        separator = ttk.Separator(main_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=10)
            
        # 左右分栏区域
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 左侧框架 - 事件类型、搜索和模板列表
        left_frame = ttk.Frame(content_frame, width=200)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        # 搜索框架
        search_frame = ttk.Frame(left_frame)
        search_frame.pack(fill=tk.X, pady=5)
        
        # 搜索标签和输入框
        self.search_label = ttk.Label(search_frame, text=self._("search") + ":")
        self.search_label.pack(side=tk.LEFT, padx=5)
        self.search_label.language_key = "search"
        
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        # 綁定Enter鍵進行搜索
        self.search_entry.bind("<Return>", self._perform_search_on_enter)
        
        self.clear_search_btn = ttk.Button(search_frame, text=self._("clear"), command=self._clear_search)
        self.clear_search_btn.pack(side=tk.RIGHT, padx=5)
        self.clear_search_btn.language_key = "clear"
        
        # 事件类型框架
        self.event_frame = ttk.LabelFrame(left_frame, text=self._("select_event_type"))
        self.event_frame.pack(fill=tk.X, pady=5)
        self.event_frame.language_key = "select_event_type"
        
        # 事件类型相关控件
        event_control_frame = ttk.Frame(self.event_frame)
        event_control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 事件类型下拉菜单
        self.event_type_combobox = ttk.Combobox(event_control_frame, textvariable=self.selected_event_type, state="readonly")
        self.event_type_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 添加/删除事件类型按钮
        event_button_frame = ttk.Frame(event_control_frame)
        event_button_frame.pack(side=tk.RIGHT)
        
        self.add_event_btn = ttk.Button(event_button_frame, text="+", width=2, command=self._add_event_type)
        self.add_event_btn.pack(side=tk.LEFT, padx=2)
        
        self.del_event_btn = ttk.Button(event_button_frame, text="-", width=2, command=self._delete_event_type)
        self.del_event_btn.pack(side=tk.LEFT, padx=2)
        
        # 模板列表框架
        self.templates_frame = ttk.LabelFrame(left_frame, text=self._("template_list"))
        self.templates_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.templates_frame.language_key = "template_list"
        
        # 模板列表
        self.template_listbox = tk.Listbox(self.templates_frame)
        self.template_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 为模板列表添加滚动条
        template_scrollbar = ttk.Scrollbar(self.templates_frame, orient=tk.VERTICAL, command=self.template_listbox.yview)
        template_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.template_listbox.configure(yscrollcommand=template_scrollbar.set)
        
        # 右侧框架 - 模板预览和变量填写
        right_frame = ttk.Frame(content_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        # HTML预览选项
        preview_option_frame = ttk.Frame(right_frame)
        preview_option_frame.pack(fill=tk.X, pady=5)
        
        # 预览框架
        self.preview_frame = ttk.LabelFrame(right_frame, text=self._("template_preview"))
        self.preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.preview_frame.language_key = "template_preview"
        
        # 变量输入框架
        self.variables_frame = ttk.LabelFrame(right_frame, text=self._("fill_variables"))
        self.variables_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.variables_frame.language_key = "fill_variables"
        
        # 添加右鍵菜單支持
        self._create_context_menus()
        
        # 綁定右鍵菜單到模板列表和事件類型下拉菜單
        self.template_listbox.bind("<Button-3>", self._show_template_context_menu)
        self.event_type_combobox.bind("<Button-3>", self._show_event_context_menu)

        # 变量框架中添加滚动条
        var_canvas = tk.Canvas(self.variables_frame)
        var_scrollbar = ttk.Scrollbar(self.variables_frame, orient="vertical", command=var_canvas.yview)
        var_scrollable_frame = ttk.Frame(var_canvas)
        
        var_scrollable_frame.bind(
            "<Configure>",
            lambda e: var_canvas.configure(scrollregion=var_canvas.bbox("all"))
        )
        
        var_canvas.create_window((0, 0), window=var_scrollable_frame, anchor="nw")
        var_canvas.configure(yscrollcommand=var_scrollbar.set)
        
        # 修改鼠标滚轮事件，只在鼠标悬停在特定区域时生效
        def _on_mousewheel(event, canvas):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            return "break"  # 阻止事件继续传播
        
        def _on_var_frame_enter(event):
            # 绑定滚轮事件到变量框架
            var_canvas.bind_all("<MouseWheel>", lambda e: _on_mousewheel(e, var_canvas))
        
        def _on_var_frame_leave(event):
            # 解绑滚轮事件
            var_canvas.unbind_all("<MouseWheel>")
        
        # 绑定鼠标进入和离开事件
        var_canvas.bind("<Enter>", _on_var_frame_enter)
        var_canvas.bind("<Leave>", _on_var_frame_leave)
        
        var_canvas.pack(side="left", fill="both", expand=True)
        var_scrollbar.pack(side="right", fill="y")
        
        self.var_scrollable_frame = var_scrollable_frame
        
        # 按钮框架 - 底部 (現在使用Grid布局固定在窗口底部)
        button_frame = ttk.Frame(self.root)
        button_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        
        # 左侧按钮组
        left_buttons = ttk.Frame(button_frame)
        left_buttons.pack(side=tk.LEFT)
        
        self.refresh_btn = ttk.Button(left_buttons, text=self._("refresh"), command=self._refresh_templates)
        self.refresh_btn.pack(side=tk.LEFT, padx=5)
        self.refresh_btn.language_key = "refresh"
        
        self.add_template_btn = ttk.Button(left_buttons, text=self._("add_template"), command=self._add_template)
        self.add_template_btn.pack(side=tk.LEFT, padx=5)
        self.add_template_btn.language_key = "add_template"
        
        self.edit_btn = ttk.Button(left_buttons, text=self._("edit"), command=self._edit_selected_template)
        self.edit_btn.pack(side=tk.LEFT, padx=5)
        self.edit_btn.language_key = "edit"
        
        self.delete_btn = ttk.Button(left_buttons, text=self._("delete"), command=self._delete_selected_template)
        self.delete_btn.pack(side=tk.LEFT, padx=5)
        self.delete_btn.language_key = "delete"
        
        # 簽名檔選項框架
        signature_frame = ttk.Frame(button_frame)
        signature_frame.pack(side=tk.RIGHT, padx=5)

        # 簽名檔標籤
        signature_label = ttk.Label(signature_frame, text=self._("signature") + ":")
        signature_label.pack(side=tk.LEFT, padx=2)

        # 獲取 Outlook 簽名檔列表
        signatures = self.email_generator.get_outlook_signatures()

        # 簽名檔下拉選單
        self.signature_var = StringVar(value="<Default>")
        self.signature_combobox = ttk.Combobox(
            signature_frame, 
            textvariable=self.signature_var,
            values=signatures,
            width=15,
            state="readonly"
        )
        self.signature_combobox.pack(side=tk.LEFT, padx=2)

        # 右侧生成邮件按钮
        self.generate_btn = ttk.Button(button_frame, text=self._("generate_email"), command=self._generate_email)
        self.generate_btn.pack(side=tk.RIGHT, padx=5)
        self.generate_btn.language_key = "generate_email"
        
        # 状态栏 (同樣使用Grid布局固定在窗口底部)
        self.status_var = StringVar()
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=2, column=0, sticky="ew")
        
        # 检查 Outlook 是否可用
        if not self.email_generator.is_outlook_available():
            self.status_var.set(self._("outlook_unavailable"))
            messagebox.showwarning(self._("warning"), self._("outlook_unavailable_msg"))
        else:
            self.status_var.set(self._("ready"))
        
        # 绑定事件
        self.event_type_combobox.bind("<<ComboboxSelected>>", self._on_event_type_selected)
        self.template_listbox.bind("<<ListboxSelect>>", self._on_template_selected)
        self.template_listbox.bind("<Double-1>", lambda e: self._edit_selected_template())
    
    def _create_menu(self):
        """創建菜單欄"""
        menu_bar = tk.Menu(self.root)
        self.root.config(menu=menu_bar)
        
        # 文件菜單
        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label=self._("file"), menu=file_menu)
        #file_menu.add_command(label=self._("import_templates"), command=self._import_templates)
        #file_menu.add_command(label=self._("export_templates"), command=self._export_templates)
        file_menu.add_separator()
        file_menu.add_command(label=self._("backup_now"), command=self._backup_database)
        file_menu.add_separator()
        file_menu.add_command(label=self._("exit"), command=self.root.quit)

        # 語言菜單
        language_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label=self._("language"), menu=language_menu)
        
        # 獲取可用語言
        available_languages = self.language_manager.get_available_languages()
        self.language_var = tk.StringVar(value=self.language_manager.current_language)
        
        # 添加語言選項
        for lang_code, lang_name in available_languages.items():
            language_menu.add_radiobutton(
                label=lang_name, 
                variable=self.language_var, 
                value=lang_code,
                command=self._change_language
            )

        # "More" or "說明" Menu item
        more_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label=self._("more"), menu=more_menu)  # "More" or "說明" menu item
        more_menu.add_command(label=self._("about"), command=self._show_about)  # Open About dialog
    
    def _perform_search_on_enter(self, event):
        """按下Enter鍵時執行搜索"""
        search_text = self.search_var.get().lower().strip()
        event_type = self.selected_event_type.get()
        if event_type:
            self._perform_search(search_text, event_type)

    def _perform_search(self, search_text, event_type):
        """执行搜索操作"""
        # 查询模板并更新模板列表
        results = self.template_manager.search_templates(search_text)
        filtered_templates = [template for template in results if template["event_type"] == event_type]
        
        # 清空模板列表并重新插入
        self.template_listbox.delete(0, tk.END)
        for template in filtered_templates:
            self.template_listbox.insert(tk.END, template["template"]["name"])

    def search_templates(self, keyword: str) -> List[Dict]:
        """搜索包含关键字的模板"""
        conn = self.get_connection()  # Each thread uses its own connection
        cursor = conn.cursor()

        keyword = f"%{keyword}%"  # Use LIKE for fuzzy matching

        cursor.execute(
            """SELECT t.id, t.name, t.recipient, t.cc, t.subject, t.body, t.note_en, t.tag_en, et.name as event_type
            FROM templates t
            JOIN event_types et ON t.event_type_id = et.id
            WHERE t.name LIKE ? OR t.subject LIKE ? OR t.body LIKE ?""",
            (keyword, keyword, keyword)
        )

        results = []
        for template in cursor.fetchall():
            template_dict = dict(template)

            # Modify key names to match expected output
            if 'recipient' in template_dict:
                template_dict['to'] = template_dict['recipient']
                del template_dict['recipient']

            # Get variables
            cursor.execute(
                "SELECT variable_name FROM template_variables WHERE template_id = ?",
                (template_dict['id'],)
            )

            template_dict['variables'] = [row[0] for row in cursor.fetchall()]
            results.append({
                "event_type": template_dict['event_type'],
                "template": {
                    "name": template_dict['name'],
                    "to": template_dict['to'],
                    "cc": template_dict['cc'],
                    "subject": template_dict['subject'],
                    "body": template_dict['body'],
                    "variables": template_dict['variables'],
                    "note_en": template_dict['note_en'],
                    "tag_en": template_dict['tag_en'],
                }
            })
            

        conn.close()  # Ensure the connection is closed
        return results

    def _clear_search(self):
        """清空搜索框"""
        self.search_var.set("")
        # 當清空搜索框時，重新加載當前事件類型的所有模板
        event_type = self.selected_event_type.get()
        if event_type:
            self._load_templates_for_event_type(event_type)
    
    def _on_event_type_selected(self, event):
        """事件类型选择时的回调"""
        # 更新模板列表 - 只在選擇事件類型時加載對應模板
        event_type = self.selected_event_type.get()
        if event_type:
            self._load_templates_for_event_type(event_type)
    
    def _load_templates_for_event_type(self, event_type):
        """加載特定事件類型的模板"""
        # 清空模板列表
        self.template_listbox.delete(0, tk.END)
        
        # 獲取該事件類型的所有模板
        templates = self.template_manager.get_templates_for_event(event_type)
        for template in templates:
            self.template_listbox.insert(tk.END, template["name"])
    
    def _on_template_selected(self, event):
        """模板选择时的回调"""
        # 获取所选模板
        selected_indices = self.template_listbox.curselection()
        if not selected_indices:
            return
            
        template_name = self.template_listbox.get(selected_indices[0])
        self.selected_template.set(template_name)  # 更新选定的模板
        event_type = self.selected_event_type.get()

        template = self.template_manager.get_template(event_type, template_name)
        if not template:
            return

        # 清空变量输入框
        self._clear_variable_entries()
        
        # 更新预览 - 根據HTML預覽設置決定顯示方式
        self._update_preview(template)
        
        # 创建变量输入框
        self._create_variable_entries(template)

    def _update_preview(self, template):
        """更新預覽"""
        # 首先清空预览区域
        for widget in self.preview_frame.winfo_children():
            widget.destroy()
            
        # 創建模板信息框架
        info_frame = ttk.Frame(self.preview_frame)
        info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 各种信息标签
        row = 0
        
        # 模板名称
        ttk.Label(info_frame, text=f"{self._('name')}:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(info_frame, text=template.get("name", "")).grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        row += 1
        
        # Note (使用大写的Note术语)
        ttk.Label(info_frame, text=f"{self._('Note')}:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(info_frame, text=template.get("note_en", "")).grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        row += 1
        
        # Tag (使用大写的Tag术语)
        ttk.Label(info_frame, text=f"{self._('Tag')}:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(info_frame, text=template.get("tag_en", "")).grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        row += 1
        
        # 新增: 寄件人
        ttk.Label(info_frame, text=f"{self._('sender')}:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        sender_value = template.get("sender", "")  # 確保正確獲取 sender 值
        ttk.Label(info_frame, text=sender_value).grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        row += 1

        # 收件人
        ttk.Label(info_frame, text=f"{self._('recipient')}:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(info_frame, text=template.get("to", "")).grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        row += 1
        
        # 抄送
        ttk.Label(info_frame, text=f"{self._('cc')}:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(info_frame, text=template.get("cc", "")).grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        row += 1
        
        # 主题
        ttk.Label(info_frame, text=f"{self._('subject')}:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(info_frame, text=template.get("subject", "")).grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        row += 1
        
        # 内容标签
        ttk.Label(self.preview_frame, text=f"{self._('content')}:").pack(anchor=tk.W, padx=10, pady=2)
        
        # 内容预览文本框
        preview_frame = ttk.Frame(self.preview_frame)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        preview_body = tk.Text(preview_frame, wrap=tk.WORD, height=10)
        preview_body.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 添加滚动条
        preview_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=preview_body.yview)
        preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        preview_body.configure(yscrollcommand=preview_scrollbar.set)
        
        # 配置变量标签样式（红色粗体）
        preview_body.tag_configure("variable", foreground="red", font=("TkDefaultFont", 10, "bold"))
        
        body = template.get("body", "")
        
        # 始終使用優化的HTML處理方式
        if "<html>" in body.lower() or any(tag in body.lower() for tag in ["<img", "<p>", "<div>", "<span>", "<table>", "<br"]):
            # 使用後台線程處理HTML以避免UI凍結
            self._load_html_preview(preview_body, body, template)
        else:
            # 純文本內容直接顯示
            self._insert_text_with_variable_highlight(preview_body, body)
            preview_body.config(state=tk.DISABLED)
            
        # 保存原始HTML内容到模板中（不改变）
        self.template = template
        
        # 為內容預覽添加鼠標滾輪事件處理
        def _on_preview_mousewheel(event):
            preview_body.yview_scroll(int(-1*(event.delta/120)), "units")
            return "break"  # 阻止事件继续传播
        
        def _on_preview_enter(event):
            # 绑定滚轮事件到预览文本框
            preview_body.bind("<MouseWheel>", _on_preview_mousewheel)
        
        def _on_preview_leave(event):
            # 解绑滚轮事件
            preview_body.unbind("<MouseWheel>")
        
        # 绑定鼠标进入和离开事件
        preview_body.bind("<Enter>", _on_preview_enter)
        preview_body.bind("<Leave>", _on_preview_leave)

    
    def _load_html_preview(self, preview_body, html_content, template):
        """在後台線程中載入HTML預覽，優化處理大量圖片"""
        # 使用緩存來提高性能
        template_id = f"{template.get('name', '')}"
        
        # 檢查是否有圖片數據（base64或cid引用）
        has_images = "data:image" in html_content or "cid:" in html_content
        
        # 檢查緩存中是否已有處理過的內容
        if template_id in self.cached_template_content:
            try:
                if preview_body.winfo_exists():  # 檢查widget是否存在
                    body_text = self.cached_template_content[template_id]
                    self._update_preview_content(preview_body, body_text)
            except Exception as e:
                print(f"更新預覽內容時出錯: {e}")
        else:
            try:
                # 先檢查preview_body是否存在
                if not preview_body.winfo_exists():
                    print("預覽窗口已不存在")
                    return
                    
                # 先清空預覽並顯示處理中的訊息
                preview_body.config(state=tk.NORMAL)
                preview_body.delete(1.0, tk.END)
                preview_body.insert(tk.END, "正在處理HTML內容...\n\n")
                
                if has_images:
                    preview_body.insert(tk.END, "[此模板包含圖片，為提高性能僅顯示文字內容]\n\n")
                
                preview_body.config(state=tk.DISABLED)
                
                # 開啟處理線程
                def process_html():
                    # 提取純文本用於預覽
                    processed_text = self._remove_html_tags(html_content)
                    self.cached_template_content[template_id] = processed_text
                    
                    # 使用 after 方法在主線程中更新 UI，但需確保元素仍存在
                    def safe_update():
                        try:
                            if preview_body.winfo_exists():  # 確認widget仍存在
                                self._update_preview_content(preview_body, processed_text)
                        except Exception as e:
                            print(f"執行UI更新時出錯: {e}")
                    
                    # 嘗試在主窗口或子窗口上調度更新
                    try:
                        if hasattr(self, 'root') and self.root.winfo_exists():
                            self.root.after(10, safe_update)
                        elif hasattr(self, 'window') and self.window.winfo_exists():
                            self.window.after(10, safe_update)
                        else:
                            print("找不到有效的窗口進行UI更新")
                    except Exception as e:
                        print(f"調度UI更新時出錯: {e}")
                
                # 啟動處理線程
                threading.Thread(target=process_html).start()
            except Exception as e:
                print(f"初始化HTML預覽時出錯: {e}")

    def _update_preview_content(self, text_widget, content):
        """更新預覽文本內容"""
        try:
            # 在操作前檢查widget是否仍存在
            if not text_widget.winfo_exists():
                print("文本窗口已不存在，無法更新內容")
                return
                
            text_widget.config(state=tk.NORMAL)
            text_widget.delete(1.0, tk.END)
            
            # 處理變量標記，使其為紅色粗體
            self._insert_text_with_variable_highlight(text_widget, content)
            text_widget.config(state=tk.DISABLED)  # 只讀
        except Exception as e:
            print(f"更新預覽內容時出錯: {e}")
    
    def _remove_html_tags(self, html):
        """移除 HTML 标签並將 HTML 內容轉換為文本，保留換行符，替換base64圖片為[圖片]標籤"""
        # 處理含有base64數據的圖片標籤
        html = re.sub(r'<img [^>]*?src="data:image[^>]*?/>', '[Pic]', html)
        
        # 處理其他圖片標籤，包括CID引用的圖片
        html = re.sub(r'<img [^>]*?src="cid:[^>]*?/>', '[Pic]', html)
        html = re.sub(r'<img [^>]*?/>', '[Pic]', html)
        
        # 替換 &nbsp; 為空格
        html = html.replace("&nbsp;", " ")
        
        # 處理列表項
        html = re.sub(r'<li>(.*?)</li>', r'• \1\n', html)
        html = re.sub(r'<ol>|</ol>|<ul>|</ul>', '', html)
        
        # 處理段落和換行
        html = re.sub(r'<p>(.*?)</p>', r'\1\n\n', html)  # 段落替換為內容加兩個換行
        html = re.sub(r'<br\s*/?>', '\n', html)          # <br>替換為換行
        html = html.replace('<p>', '').replace('</p>', '\n\n')  # 處理未成對的p標籤
        
        # 替換其他可能導致換行的標籤
        html = re.sub(r'</div>|</h[1-6]>|</tr>|</thead>|</tbody>|</table>', '\n', html)
        
        # 處理a標籤
        html = re.sub(r'<a [^>]*?>(.*?)</a>', r'\1', html)
        
        # 處理strong/b標籤
        html = re.sub(r'<(strong|b)>(.*?)</\1>', r'\2', html)
        
        # 清除剩餘所有 HTML 標籤
        clean_text = re.sub(r'<[^>]*>', '', html)
        
        # 標準化空白符
        clean_text = re.sub(r' +', ' ', clean_text)       # 多個空格變一個
        clean_text = re.sub(r'\n{3,}', '\n\n', clean_text)  # 3個以上換行變2個
        
        # 去除每行首尾空白
        lines = [line.strip() for line in clean_text.split('\n')]
        clean_text = '\n'.join(lines)
        
        # 去掉整個文本首尾的空白
        clean_text = clean_text.strip()
        
        return clean_text
        
    def _insert_text_with_variable_highlight(self, text_widget, content):
        """插入文本並高亮變數標記"""
        # 正則表達式匹配{variable}格式的變數
        pattern = r'(\{[^{}]+\})'
        
        # 分割文本，使變數部分和普通文本分開
        parts = re.split(pattern, content)
        
        # 遍歷所有部分並插入文本，為變數應用特殊標籤
        for part in parts:
            if re.match(pattern, part):
                # 這是一個變數，使用變數標籤
                text_widget.insert(tk.END, part, "variable")
            else:
                # 這是普通文本
                text_widget.insert(tk.END, part)
        
    def _create_variable_entries(self, template):
        """創建變量輸入框，根據郵件內容中出現的順序 (包括同一行多變數)"""
        # 清空變量框
        self._clear_variable_entries()
        
        if "variables" not in template or not template["variables"]:
            ttk.Label(self.var_scrollable_frame, text=self._("no_variables") if hasattr(self, '_') and callable(self._) else "No variables").pack(pady=10)
            return
        
        variables = template.get("variables", [])
        subject = template.get("subject", "")
        body = template.get("body", "")
        
        ordered_variables = []
        seen = set()
        
        # Helper function: extract in-order unique vars
        def extract_vars(text):
            matches = re.findall(r'\{([^{}]+)\}', text)
            for var in matches:
                if var in variables and var not in seen:
                    ordered_variables.append(var)
                    seen.add(var)
        
        # 1. 先處理主題行
        extract_vars(subject)
        
        # 2. 再處理正文
        extract_vars(body)
        
        # 3. 剩餘沒出現的
        for var in variables:
            if var not in seen:
                ordered_variables.append(var)
                seen.add(var)
        
        # 4. 建立變量輸入框
        for var_name in ordered_variables:
            var_frame = ttk.Frame(self.var_scrollable_frame)
            var_frame.pack(fill=tk.X, padx=5, pady=2)
            
            var_label = ttk.Label(var_frame, text=f"{var_name}:", width=15, anchor=tk.W)
            var_label.pack(side=tk.LEFT, padx=5)
            
            var_entry = ttk.Entry(var_frame, width=40)
            var_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            
            # 如果存在已保存的值，恢復它
            template_name = template.get("name", "")
            var_key = f"{template_name}:{var_name}"
            if var_key in self.variable_values:
                var_entry.insert(0, self.variable_values[var_key])
            
            # 綁定變更事件以保存值
            var_entry.bind("<KeyRelease>", lambda e, name=var_name, template_name=template_name: 
                        self._save_var_value(template_name, name, e.widget.get()))
            
            self.var_entries[var_name] = var_entry

    def _save_var_value(self, template_name, var_name, value):
        """保存變量值到緩存"""
        var_key = f"{template_name}:{var_name}"
        self.variable_values[var_key] = value

    
    def _create_context_menus(self):
        """創建右鍵菜單"""
        # 模板右鍵菜單
        self.template_context_menu = tk.Menu(self.root, tearoff=0)
        self.template_context_menu.add_command(label=self._("add_template"), command=self._add_template)
        self.template_context_menu.add_command(label=self._("edit_template"), command=self._edit_selected_template)
        self.template_context_menu.add_command(label=self._("delete_template"), command=self._delete_selected_template)
        self.template_context_menu.add_separator()
        self.template_context_menu.add_command(label=self._("move_to_event_type"), command=self._move_template_to_event_type)
        
        # 事件類型右鍵菜單
        self.event_context_menu = tk.Menu(self.root, tearoff=0)
        self.event_context_menu.add_command(label=self._("add_event_type"), command=self._add_event_type)
        self.event_context_menu.add_command(label=self._("rename_event_type"), command=self._rename_event_type)
        self.event_context_menu.add_command(label=self._("delete_event_type"), command=self._delete_event_type)

        # 綁定右鍵菜單
        self.template_listbox.bind("<Button-3>", self._show_template_context_menu)
        self.event_type_combobox.bind("<Button-3>", self._show_event_context_menu)

    def _show_template_context_menu(self, event):
        """顯示模板右鍵菜單"""
        # 獲取鼠標位置對應的列表項索引
        index = self.template_listbox.nearest(event.y)
        if index >= 0:
            self.template_listbox.selection_clear(0, tk.END)
            self.template_listbox.selection_set(index)
            self.template_listbox.activate(index)
            self._on_template_selected(None)  # 更新預覽
            self.template_context_menu.tk_popup(event.x_root, event.y_root)

    def _show_event_context_menu(self, event):
        """顯示事件類型右鍵菜單"""
        self.event_context_menu.tk_popup(event.x_root, event.y_root)

    def _move_template_to_event_type(self):
        """將選中的模板移動到另一個事件類型"""
        # 獲取當前選中的模板
        selected_indices = self.template_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning(self._("warning"), self._("please_select_template"))
            return
        
        template_name = self.template_listbox.get(selected_indices[0])
        current_event_type = self.selected_event_type.get()
        
        # 獲取模板數據
        template = self.template_manager.get_template(current_event_type, template_name)
        if not template:
            return
        
        # 獲取所有事件類型（排除當前事件類型）
        all_event_types = self.template_manager.get_event_types()
        other_event_types = [et for et in all_event_types if et != current_event_type]
        
        if not other_event_types:
            messagebox.showinfo(self._("information"), self._("no_other_event_types"))
            return
        
        # 創建事件類型選擇對話框
        dialog = tk.Toplevel(self.root)
        dialog.title(self._("select_target_event_type"))
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # 獲取主窗口位置和大小，以便選擇窗口居中顯示
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        
        # 計算居中位置
        dialog_width = 300
        dialog_height = 150
        x_pos = root_x + (root_width - dialog_width) // 2
        y_pos = root_y + (root_height - dialog_height) // 2
        
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x_pos}+{y_pos}")
        
        # 創建框架
        frame = ttk.Frame(dialog, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 添加說明標籤
        ttk.Label(frame, text=self._("move_template_description").format(name=template_name),
                wraplength=280).pack(pady=5)
        
        # 事件類型選擇
        target_event_type = tk.StringVar()
        target_combobox = ttk.Combobox(frame, textvariable=target_event_type, state="readonly")
        target_combobox['values'] = other_event_types
        target_combobox.pack(fill=tk.X, pady=10)
        target_combobox.current(0)  # 默認選擇第一個事件類型
        
        # 按鈕框架
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        # 取消按鈕
        ttk.Button(button_frame, text=self._("cancel"), command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        
        # 確定按鈕
        def confirm_move():
            target = target_event_type.get()
            if not target:
                return
            
            # 檢查目標事件類型中是否已存在同名模板
            existing_templates = self.template_manager.get_template_names_for_event(target)
            if template_name in existing_templates:
                # 詢問是否覆蓋
                if not messagebox.askyesno(
                    self._("confirm"),
                    self._("template_name_exists_in_target").format(name=template_name, target=target),
                    parent=dialog
                ):
                    return
            
            # 執行移動操作
            try:
                # 在目標事件類型中添加模板
                self.template_manager.add_template(target, template)
                # 從原事件類型中刪除模板
                self.template_manager.remove_template(current_event_type, template_name)
                
                # 顯示成功消息
                messagebox.showinfo(
                    self._("success"),
                    self._("template_moved_successfully").format(name=template_name, target=target)
                )
                
                # 刷新模板列表
                self._load_templates_for_event_type(current_event_type)
                
                dialog.destroy()
            except Exception as e:
                messagebox.showerror(
                    self._("error"),
                    self._("template_move_failed").format(error=str(e))
                )
        
        ttk.Button(button_frame, text=self._("move"), command=confirm_move).pack(side=tk.RIGHT, padx=5)

    def _rename_event_type(self):
        """重命名事件類型"""
        event_type = self.selected_event_type.get()
        if not event_type:
            messagebox.showwarning(self._("warning"), self._("please_select_event_type"))
            return
        
        # 詢問新名稱
        new_name = simpledialog.askstring(
            self._("rename_event_type"),
            self._("enter_new_event_type_name"),
            initialvalue=event_type,
            parent=self.root
        )
        
        if not new_name or new_name == event_type:
            return
        
        # 檢查新名稱是否已存在
        existing_types = self.template_manager.get_event_types()
        if new_name in existing_types:
            messagebox.showwarning(
                self._("warning"),
                self._("event_type_exists").format(name=new_name)
            )
            return
        
        # 執行重命名操作
        try:
            # 獲取所有該事件類型的模板
            templates = self.template_manager.get_templates_for_event(event_type)
            
            # 創建新事件類型
            self.template_manager.add_event_type(new_name)
            
            # 將所有模板移動到新事件類型
            for template in templates:
                self.template_manager.add_template(new_name, template)
            
            # 刪除舊事件類型（會自動刪除關聯的模板）
            self.template_manager.remove_event_type(event_type)
            
            # 刷新事件類型下拉菜單
            self._update_event_types()
            
            # 選擇新事件類型
            self.event_type_combobox.set(new_name)
            self._on_event_type_selected(None)
            
            messagebox.showinfo(
                self._("success"),
                self._("event_type_renamed_successfully").format(old=event_type, new=new_name)
            )
        except Exception as e:
            messagebox.showerror(
                self._("error"),
                self._("event_type_rename_failed").format(error=str(e))
            )
    def _clear_variable_entries(self):
        """清空變量輸入框"""
        # 在清空前保存當前值
        for var_name, entry in self.var_entries.items():
            template_name = self.selected_template.get()
            if template_name:
                var_key = f"{template_name}:{var_name}"
                self.variable_values[var_key] = entry.get()
        
        # 清空字典
        self.var_entries.clear()
        
        # 清空框架
        for widget in self.var_scrollable_frame.winfo_children():
            widget.destroy()
    
    def _add_template(self):
        """添加新模板"""
        event_type = self.selected_event_type.get()
        if not event_type:
            messagebox.showwarning(self._("warning"), self._("please_select_event_type"))
            return

        # 创建空白模板，只保留tag_en和note_en
        template = {
            "name": "",
            "to": "",
            "cc": "",
            "subject": "",
            "body": "",
            "variables": [],
            "note_en": "",
            "tag_en": ""
        }

        # 打开编辑窗口时传递 template
        edit_window = EditTemplateWindow(
            parent=self.root,
            template_manager=self.template_manager,
            event_type=event_type,
            template=template,
            is_new=True,
            language_manager=self.language_manager,
            refresh_callback=self._refresh_templates,
            accounts=self.accounts
        )
        self.root.wait_window(edit_window.window)

    def _edit_selected_template(self):
        """编辑所选模板"""
        selected_indices = self.template_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning(self._("warning"), self._("please_select_template"))
            return

        template_name = self.template_listbox.get(selected_indices[0])
        event_type = self.selected_event_type.get()

        template = self.template_manager.get_template(event_type, template_name)
        if not template:
            return

        # 打开编辑窗口时传递 template
        edit_window = EditTemplateWindow(
            parent=self.root,
            template_manager=self.template_manager,
            event_type=event_type,
            template=template,
            is_new=False,  # 编辑时設為False
            language_manager=self.language_manager,
            refresh_callback=self._refresh_templates,
            accounts=self.accounts
        )
        self.root.wait_window(edit_window.window)
        
    def _delete_selected_template(self):
        """删除所选模板，需要输入"Confirm"确认"""
        selected_indices = self.template_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning(self._("warning"), self._("please_select_template"))
            return
        
        template_name = self.template_listbox.get(selected_indices[0])
        event_type = self.selected_event_type.get()
        
        # 创建确认对话框
        confirm_dialog = tk.Toplevel(self.root)
        confirm_dialog.title(self._("confirm"))
        confirm_dialog.transient(self.root)
        confirm_dialog.grab_set()
        
        # 获取主窗口位置和大小，以便确认窗口居中显示
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        
        # 计算确认窗口位置
        dialog_width = 400
        dialog_height = 180
        x_pos = root_x + (root_width - dialog_width) // 2
        y_pos = root_y + (root_height - dialog_height) // 2
        
        confirm_dialog.geometry(f"{dialog_width}x{dialog_height}+{x_pos}+{y_pos}")
        confirm_dialog.resizable(False, False)
        
        # 警告消息
        message_frame = ttk.Frame(confirm_dialog, padding=10)
        message_frame.pack(fill=tk.BOTH, expand=True)
        
        warning_label = ttk.Label(
            message_frame, 
            text=self._("confirm_delete_template").format(name=template_name),
            wraplength=380,
            font=("", 10, "")
        )
        warning_label.pack(pady=5)
        
        instruction_label = ttk.Label(
            message_frame, 
            text=self._("delete_template_warning") if hasattr(self, '_') and callable(self._) and self._("delete_template_warning") != "delete_template_warning" else "Type 'Confirm' to delete this template:",
            font=("", 9, "italic"),
            foreground="red"
        )
        instruction_label.pack(pady=5)
        
        # 确认输入框
        confirm_var = tk.StringVar()
        confirm_entry = ttk.Entry(message_frame, textvariable=confirm_var, width=20)
        confirm_entry.pack(pady=5)
        confirm_entry.focus_set()
        
        # 按钮框架
        button_frame = ttk.Frame(message_frame)
        button_frame.pack(pady=10, fill=tk.X)
        
        # 取消按钮
        cancel_button = ttk.Button(
            button_frame, 
            text=self._("cancel"),
            command=confirm_dialog.destroy
        )
        cancel_button.pack(side=tk.RIGHT, padx=5)
        
        # 确认按钮（删除）
        delete_button = ttk.Button(
            button_frame,
            text=self._("delete"),
            command=lambda: self._confirm_delete_with_text(confirm_dialog, event_type, template_name, confirm_var.get())
        )
        delete_button.pack(side=tk.RIGHT, padx=5)
        
        # 绑定回车键
        confirm_entry.bind("<Return>", lambda e: self._confirm_delete_with_text(confirm_dialog, event_type, template_name, confirm_var.get()))
        
        # 等待窗口关闭
        self.root.wait_window(confirm_dialog)

    def _confirm_delete_with_text(self, dialog, event_type, template_name, confirmation_text):
        """根据输入文本确认删除模板"""
        if confirmation_text == "Confirm":
            dialog.destroy()
            # 删除模板
            self.template_manager.remove_template(event_type, template_name)
            # 更新模板列表
            self._load_templates_for_event_type(event_type)
            # 显示成功消息
            self.status_var.set(self._("template_deleted") if hasattr(self, '_') and callable(self._) and self._("template_deleted") != "template_deleted" else f"Template '{template_name}' deleted")
        else:
            messagebox.showinfo(
                self._("info"), 
                self._("confirm_text_mismatch") if hasattr(self, '_') and callable(self._) and self._("confirm_text_mismatch") != "confirm_text_mismatch" else "Please type 'Confirm' exactly to delete the template"
            )
    
    def _add_event_type(self):
        """添加新的事件类型"""
        event_type = simpledialog.askstring(self._("add"), self._("enter_event_type"))
        if not event_type:
            return
        
        # 检查是否已存在
        existing_types = self.template_manager.get_event_types()
        if event_type in existing_types:
            messagebox.showwarning(self._("warning"), self._("event_type_exists").format(name=event_type))
            return
        
        # 添加新事件类型
        self.template_manager.add_event_type(event_type)
        
        # 更新下拉菜单
        self._update_event_types()
        
        # 选择新添加的事件类型
        self.event_type_combobox.set(event_type)
        self._on_event_type_selected(None)
    
    def _delete_event_type(self):
        """刪除事件類型，需要用戶輸入確認"""
        event_type = self.selected_event_type.get()
        if not event_type:
            messagebox.showwarning(self._("warning"), self._("please_select_event_type"))
            return
        
        # 創建確認對話框
        confirm_dialog = tk.Toplevel(self.root)
        confirm_dialog.title(self._("confirm"))
        confirm_dialog.transient(self.root)
        confirm_dialog.grab_set()
        
        # 獲取主窗口位置和大小，以便確認窗口居中顯示
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        
        # 計算確認窗口位置
        dialog_width = 400
        dialog_height = 180
        x_pos = root_x + (root_width - dialog_width) // 2
        y_pos = root_y + (root_height - dialog_height) // 2
        
        confirm_dialog.geometry(f"{dialog_width}x{dialog_height}+{x_pos}+{y_pos}")
        confirm_dialog.resizable(False, False)
        
        # 警告訊息
        message_frame = ttk.Frame(confirm_dialog, padding=10)
        message_frame.pack(fill=tk.BOTH, expand=True)
        
        warning_label = ttk.Label(
            message_frame, 
            text=self._("confirm_delete_event").format(name=event_type),
            wraplength=380,
            font=("", 10, "")
        )
        warning_label.pack(pady=5)
        
        instruction_label = ttk.Label(
            message_frame, 
            text=self._("delete_confirm_instruction") if hasattr(self, '_') and self._("delete_confirm_instruction") != "delete_confirm_instruction" else "Type 'Confirm' to delete this event type:",
            font=("", 9, "italic"),
            foreground="red"
        )
        instruction_label.pack(pady=5)
        
        # 確認輸入框
        confirm_var = tk.StringVar()
        confirm_entry = ttk.Entry(message_frame, textvariable=confirm_var, width=20)
        confirm_entry.pack(pady=5)
        confirm_entry.focus_set()
        
        # 按鈕框架
        button_frame = ttk.Frame(message_frame)
        button_frame.pack(pady=10, fill=tk.X)
        
        # 取消按鈕
        cancel_button = ttk.Button(
            button_frame, 
            text=self._("cancel"),
            command=confirm_dialog.destroy
        )
        cancel_button.pack(side=tk.RIGHT, padx=5)
        
        # 確認按鈕（刪除）
        delete_button = ttk.Button(
            button_frame,
            text=self._("delete"),
            command=lambda: self._confirm_delete_event_type(confirm_dialog, event_type, confirm_var.get())
        )
        delete_button.pack(side=tk.RIGHT, padx=5)
        
        # 綁定回車鍵
        confirm_entry.bind("<Return>", lambda e: self._confirm_delete_event_type(confirm_dialog, event_type, confirm_var.get()))
        
        # 等待窗口關閉
        self.root.wait_window(confirm_dialog)

    def _confirm_delete_event_type(self, dialog, event_type, confirmation_text):
        """根据输入文本确认删除事件类型"""
        if confirmation_text == "Confirm":
            dialog.destroy()
            # 删除事件类型
            self.template_manager.remove_event_type(event_type)
            # 更新下拉菜单
            self._update_event_types()
            # 顯示成功訊息
            self.status_var.set(self._("event_type_deleted") if hasattr(self, '_') and self._("event_type_deleted") != "event_type_deleted" else f"Event type '{event_type}' deleted")
        else:
            messagebox.showinfo(
                self._("info"), 
                self._("confirm_text_mismatch") if hasattr(self, '_') and self._("confirm_text_mismatch") != "confirm_text_mismatch" else "Please type 'Confirm' exactly to delete the event type"
            )
    
    def _update_event_types(self):
        """更新事件类型下拉菜单"""
        event_types = self.template_manager.get_event_types()
        self.event_type_combobox['values'] = event_types
        
        if event_types:
            self.event_type_combobox.current(0)
            # 選擇事件類型後自動加載對應模板
            self._on_event_type_selected(None)
        else:
            # 如果沒有事件類型，清空模板列表
            self.template_listbox.delete(0, tk.END)
    
    def _refresh_templates(self):
        """刷新模板列表"""
        # 保存当前选择
        event_type = self.selected_event_type.get()
        current_template = self.selected_template.get()
        
        # 刷新列表
        if event_type:
            self._load_templates_for_event_type(event_type)
        
        # 尝试恢复选择
        if current_template:
            for i in range(self.template_listbox.size()):
                if self.template_listbox.get(i) == current_template:
                    self.template_listbox.selection_set(i)
                    self._on_template_selected(None)
                    break
    
    def _generate_email(self):
        """生成 Outlook 郵件"""
        # 檢查是否選擇了模板
        event_type = self.selected_event_type.get()
        template_name = self.selected_template.get()
        
        if not event_type or not template_name:
            messagebox.showwarning(self._("warning"), self._("please_select_template"))
            return
        
        # 獲取模板
        template = self.template_manager.get_template(event_type, template_name)
        if not template:
            messagebox.showerror(self._("error"), self._("template_not_found"))
            return
        
        # 獲取變數值
        variables = {}
        for var_name, entry in self.var_entries.items():
            variables[var_name] = entry.get()
        
        # 檢查是否所有變數都已填寫
        unfilled_vars = [var_name for var_name, value in variables.items() if not value]
        if unfilled_vars:
            unfilled_str = ", ".join(unfilled_vars)
            messagebox.showwarning(
                self._("warning"), 
                self._("please_fill_variables").format(vars=unfilled_str)
            )
            return
        
        # 從模板獲取寄件人
        sender = template.get("sender", "")
        if not sender:
            messagebox.showwarning(self._("warning"), self._("template_no_sender"))
            return
        
        # 獲取簽名檔選項
        signature_option = self.signature_var.get()
        
        # 如果選擇默認簽名檔，使用 Outlook 的內建機制
        if signature_option == "<Default>":
            # 不需要額外操作，Outlook 會自動添加默認簽名檔
            use_signature = True
        elif signature_option == "<None>":
            # 設置不使用簽名檔
            use_signature = False
        else:
            # 使用指定簽名檔 - 讓 EmailGenerator 處理
            use_signature = False  # 關閉默認簽名檔
            
        # 設置模板的簽名檔選項
        template["use_signature"] = use_signature
        template["signature_name"] = None if signature_option in ["<Default>", "<None>"] else signature_option
    
        # 生成郵件
        if self.email_generator.generate_email(template, variables, sender, signature_option):
            self.status_var.set(self._("email_generated").format(name=template_name))
        else:
            messagebox.showerror(self._("error"), self._("email_generation_failed"))
            self.status_var.set(self._("email_generation_failed"))
    
    def _import_templates(self):
        """導入模板"""
        from tkinter import filedialog
        
        # 打開文件對話框
        filename = filedialog.askopenfilename(
            title=self._("import_templates"),
            filetypes=[(self._("json_file"), "*.json"), (self._("all_files"), "*.*")]
        )
        
        if filename:
            if self.template_manager.import_templates(filename):
                messagebox.showinfo(self._("success"), self._("import_success"))
                self._update_event_types()
            else:
                messagebox.showerror(self._("error"), self._("import_error"))
    
    def _export_templates(self):
        """導出模板"""
        from tkinter import filedialog
        
        # 打開文件對話框
        filename = filedialog.asksaveasfilename(
            title=self._("export_templates"),
            defaultextension=".json",
            filetypes=[(self._("json_file"), "*.json"), (self._("all_files"), "*.*")]
        )
        
        if filename:
            try:
                self.template_manager.export_templates(filename)
                messagebox.showinfo(self._("success"), self._("export_success"))
            except Exception as e:
                messagebox.showerror(
                    self._("error"), 
                    self._("export_error").format(error=str(e))
                )
    
    def _backup_database(self):
        """备份数据库文件"""
        from tkinter import filedialog
        import shutil
        import time
        
        # 获取数据库文件路径
        db_file = self.template_manager.db_manager.db_file
        
        # 打开文件对话框选择保存位置
        backup_dir = filedialog.askdirectory(
            title=self._("backup_folder")
        )
        
        if backup_dir:
            # 生成备份文件名，包含时间戳
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            backup_file = os.path.join(backup_dir, f"app_backup_{timestamp}.db")
            
            try:
                # 复制数据库文件
                shutil.copy2(db_file, backup_file)
                messagebox.showinfo(
                    self._("success"), 
                    f"{self._('backup_success')}\n{backup_file}"
                )
            except Exception as e:
                messagebox.showerror(
                    self._("error"), 
                    f"{self._('export_error').format(error=str(e))}"
                )
    
    def _change_language(self):
        """即時更改應用程序語言無需重啟"""
        # 獲取選擇的語言
        lang_code = self.language_var.get()
        
        # 檢查是否與當前語言相同
        if lang_code == self.language_manager.current_language:
            return  # 語言未變更，不需要更新
        
        # 更新語言
        self.language_manager.set_language(lang_code)
        
        # 保存用戶偏好
        self.language_manager.save_user_preference(lang_code)
        
        # 更新界面文本
        self._update_ui_language()
        
        # 移除彈出提示，僅在狀態欄更新訊息
        self.status_var.set(self._("language_changed"))
    
    def _update_ui_language(self):
        """更新界面上所有文本元素的語言"""
        # 更新窗口標題
        self.root.title(self._("app_title"))
        
        # 更新菜單
        self._recreate_menu()
        
        # 更新右键菜单
        self._recreate_context_menus()
        
        # 遞歸更新所有按鈕和標籤
        self._update_widget_language(self.root)
        
        # 更新狀態欄
        if self.email_generator.is_outlook_available():
            self.status_var.set(self._("ready"))
        else:
            self.status_var.set(self._("outlook_unavailable"))
        
    def _update_widget_language(self, parent):
        """遞歸更新所有小部件的語言
        
        Args:
            parent: 父級小部件
        """
        # 處理當前層級所有小部件
        for widget in parent.winfo_children():
            # 檢查是否有language_key屬性
            if hasattr(widget, 'language_key'):
                if isinstance(widget, ttk.Label):
                    widget.config(text=self._(widget.language_key))
                elif isinstance(widget, ttk.Button):
                    widget.config(text=self._(widget.language_key))
                elif isinstance(widget, ttk.LabelFrame):
                    widget.config(text=self._(widget.language_key))
                elif isinstance(widget, ttk.Checkbutton):
                    widget.config(text=self._(widget.language_key))
            
            # 特殊處理某些元素
            if widget == self.edit_btn:
                widget.config(text=self._("edit"))
            elif widget == self.delete_btn:
                widget.config(text=self._("delete"))
            
            # 遞歸處理子級小部件
            if widget.winfo_children():
                self._update_widget_language(widget)

    def _recreate_menu(self):
        """重新創建菜單以更新語言"""
        # 刪除現有菜單
        self.root.config(menu="")
        
        # 重新創建菜單
        self._create_menu()
    
    def _recreate_context_menus(self):
        """重新创建右键菜单以更新语言"""
        # 重新创建模板右键菜单
        self.template_context_menu = tk.Menu(self.root, tearoff=0)
        self.template_context_menu.add_command(label=self._("add_template"), command=self._add_template)
        self.template_context_menu.add_command(label=self._("edit_template"), command=self._edit_selected_template)
        self.template_context_menu.add_command(label=self._("delete_template"), command=self._delete_selected_template)
        self.template_context_menu.add_separator()
        self.template_context_menu.add_command(label=self._("move_to_event_type"), command=self._move_template_to_event_type)
        
        # 重新创建事件类型右键菜单
        self.event_context_menu = tk.Menu(self.root, tearoff=0)
        self.event_context_menu.add_command(label=self._("add_event_type"), command=self._add_event_type)
        self.event_context_menu.add_command(label=self._("rename_event_type"), command=self._rename_event_type)
        self.event_context_menu.add_command(label=self._("delete_event_type"), command=self._delete_event_type)
    
    def _show_about(self):
        print("About clicked")

        # Create the About window
        about_window = tk.Toplevel(self.root)
        about_window.title("About")
        about_window.geometry("400x300")
        about_window.resizable(False, False)

        # Center the About window
        self._center_window_for_about(about_window)

        current_language = self.language_manager.current_language
        app_title = self.template_manager.db_manager.get_app_info("app_title", current_language)
        version = self.template_manager.db_manager.get_app_info("version", current_language)
        about_desc = self.template_manager.db_manager.get_app_info("about_desc", current_language)
        github_url = self.template_manager.db_manager.get_app_info("github_url", current_language)
        developer_info = self.template_manager.db_manager.get_app_info("developer_info", current_language)

        # Main frame
        main_frame = ttk.Frame(about_window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # App title
        app_name = ttk.Label(main_frame, text=app_title, font=("Arial", 16, "bold"))
        app_name.pack(pady=(0, 5))

        # Version
        version_label = ttk.Label(main_frame, text=f"Version: {version}", font=("Arial", 10))
        version_label.pack(pady=(0, 15))

        # Separator
        separator = ttk.Separator(main_frame, orient="horizontal")
        separator.pack(fill=tk.X, pady=10)

        # Description Textbox (with link)
        desc_text = tk.Text(main_frame, wrap=tk.WORD, height=8, width=40, font=("Arial", 10), bd=0)
        desc_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # Hyperlink setup
        desc_text.tag_configure("hyperlink", foreground="blue", underline=True)
        desc_text.tag_bind("hyperlink", "<Button-1>", lambda e: self.open_github_link(github_url))
        desc_text.tag_bind("hyperlink", "<Enter>", lambda e: desc_text.config(cursor="hand2"))
        desc_text.tag_bind("hyperlink", "<Leave>", lambda e: desc_text.config(cursor=""))

        # Insert description (no link embedding)
        desc_text.insert(tk.END, about_desc)

        desc_text.config(state=tk.DISABLED, highlightthickness=0)

        # Developer info
        dev_label = ttk.Label(main_frame, text=developer_info, font=("Arial", 9, "italic"))
        dev_label.pack(pady=(5, 0))

        # Close button
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(15, 0))
        close_btn = ttk.Button(btn_frame, text="Close", command=about_window.destroy)
        close_btn.pack(side=tk.RIGHT)


    def _center_window_for_about(self, window):
        """Center the window on the screen"""
        window.update_idletasks()  # Make sure the window is fully loaded and drawn

        # Get window dimensions
        window_width = window.winfo_width()
        window_height = window.winfo_height()

        # Get screen dimensions
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()

        # Calculate position
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # Ensure positive coordinates
        x = max(0, x)
        y = max(0, y)

        # Set the window position
        window.geometry(f"{window_width}x{window_height}+{x}+{y}")
