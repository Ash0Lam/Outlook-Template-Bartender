import tkinter as tk
from tkinter import ttk, messagebox, StringVar
import re
import os
import sys
import tempfile
import threading
import webview
import multiprocessing
from template_manager import TemplateManager

# 確保模塊可以在任何位置執行
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 讓子進程接收 Queue
def run_webview(queue, temp_path, x, y, width, height):
    import webview
    import os
    
    class Api:
        def save_content(api_self, content):
            queue.put(content)
            # 保存后自动关闭窗口
            webview.windows[0].destroy()

    api = Api()
    
    
    # 創建視窗時設置位置和大小
    window = webview.create_window("CKEditor", temp_path, js_api=api, width=width, height=height, x=x, y=y)
    
    # 使用 HTTP 伺服器模式啟動
    webview.start(http_server=True)


class EditTemplateWindow:
    """模板編輯窗口類"""

    def __init__(self, parent, template_manager, event_type, template, is_new=True, language_manager=None, refresh_callback=None, accounts=None):
        self.parent = parent
        self.content_queue = multiprocessing.Queue()
        self.template_manager = template_manager
        self.event_type = event_type
        self.template = template.copy()
        self.is_new = is_new
        self.language_manager = language_manager
        self.refresh_callback = refresh_callback
        self._ = language_manager.get_text if language_manager else lambda x: x
        self.webview_process = None
        self.accounts = accounts if accounts else []
        
        # 創建窗口
        self.window = tk.Toplevel(parent)
        self.window.title(self._("edit_template_title") if not is_new else self._("new_template_title"))
        self.window.geometry("800x700")
        self.window.minsize(800, 700)
        self.window.transient(parent)
        self.window.grab_set()
        self.window.after(500, self._check_queue)

        self.original_name = template.get("name", "")
        self.name_var = StringVar(value=self.original_name)
        self.sender_var = StringVar()  # 只定義變數，不在這裡創建控件

        
        # 獲取螢幕尺寸
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # 計算位置 - 左側位置
        window_width = 800
        window_height = 700
        pos_x = (screen_width // 2) - window_width - 50  # 中間偏左
        pos_y = (screen_height - window_height) // 2      # 垂直居中
        
        # 確保座標為正數
        pos_x = max(0, pos_x)
        pos_y = max(0, pos_y)
        
        # 設置視窗位置
        self.window.geometry(f"{window_width}x{window_height}+{pos_x}+{pos_y}")
        
        # 添加關閉視窗事件處理
        self.window.protocol("WM_DELETE_WINDOW", self._on_window_close)
        
        self.window.after(500, self._check_queue)
        

        # 創建界面元素
        self._create_widgets()

        if language_manager:
            self.parent.bind("<<LanguageChanged>>", self._on_language_changed)

    def _check_queue(self):
        try:
            content = self.content_queue.get_nowait()
            self.template['body'] = content  # 更新到 template
            print("Received content from WebView:", content)

            # 自动提取变量
            pattern = r'\{([^{}]+)\}'
            body_vars = re.findall(pattern, content)
            body_vars = sorted(list(set(body_vars)))  # 去重 & 排序
            self.variables_entry.delete(0, tk.END)
            self.variables_entry.insert(0, ", ".join(body_vars))

            messagebox.showinfo("Info", "Email content saved, variables auto-extracted!")
            
            # 尝试终止webview进程
            if self.webview_process and self.webview_process.is_alive():
                try:
                    self.webview_process.terminate()
                    self.webview_process.join(1)  # 等待最多1秒让进程终止
                except:
                    pass
        except:
            pass  # Queue 为空

        self.window.after(500, self._check_queue)

    def _validate_email(self, emails):
        pattern = r"[^@]+@[^@]+\.[^@]+"
        return all(re.match(pattern, e.strip()) for e in emails if e.strip())

    def _format_email_entry(self, entry_widget, event):
        content = entry_widget.get()
        emails = [e.strip() for e in re.split(r'[;,]', content) if e.strip()]

        # 只處理當前輸入的最後一個
        if emails:
            last_email = emails[-1]
            valid = re.match(r"[^@]+@[^@]+\.[^@]+", last_email)

            # 當使用者按下空格 或 分號
            if event.keysym in ('space', 'semicolon'):
                if valid:
                    if not content.strip().endswith(";"):
                        entry_widget.insert(tk.END, "; ")
            # 顏色提示
            if all(re.match(r"[^@]+@[^@]+\.[^@]+", e) for e in emails):
                entry_widget.config(foreground='black')
            else:
                entry_widget.config(foreground='red')

    def _add_placeholder(self, entry, placeholder_text):
        entry.insert(0, placeholder_text)
        entry.config(foreground='grey')

        def on_focus_in(event):
            if entry.get() == placeholder_text:
                entry.delete(0, tk.END)
                entry.config(foreground='black')

        def on_focus_out(event):
            if entry.get() == "":
                entry.insert(0, placeholder_text)
                entry.config(foreground='grey')

        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)

    def _create_widgets(self):
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)


        # 設定 Grid 欄位寬度
        main_frame.grid_columnconfigure(0, weight=0)  # Label
        main_frame.grid_columnconfigure(1, weight=1)  # Entry

        label_width = 15
        entry_width = 50

        # Template Name
        ttk.Label(main_frame, text=self._("template_name") + ":", width=label_width, anchor='e').grid(row=0, column=0, sticky='e', padx=5, pady=5)
        name_entry = ttk.Entry(main_frame, textvariable=self.name_var, width=entry_width)
        name_entry.grid(row=0, column=1, sticky='we', padx=5, pady=5)

        # Note
        ttk.Label(main_frame, text=self._("Note") + ":", width=label_width, anchor='e').grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.note_entry = ttk.Entry(main_frame, width=entry_width)
        self.note_entry.grid(row=1, column=1, sticky='we', padx=5, pady=5)
        self.note_entry.insert(0, self.template.get("note_en", ""))

        # Tag
        ttk.Label(main_frame, text=self._("Tag") + ":", width=label_width, anchor='e').grid(row=2, column=0, sticky='e', padx=5, pady=5)
        self.tag_entry = ttk.Entry(main_frame, width=entry_width)
        self.tag_entry.grid(row=2, column=1, sticky='we', padx=5, pady=5)
        self.tag_entry.insert(0, self.template.get("tag_en", ""))

        # Sender
        ttk.Label(main_frame, text=self._("sender") + ":", width=label_width, anchor='e').grid(row=3, column=0, sticky='e', padx=5, pady=5)
        account_emails = [acc["email"] for acc in self.accounts] if self.accounts else []
        self.sender_combobox = ttk.Combobox(main_frame, textvariable=self.sender_var, width=entry_width)
        self.sender_combobox['values'] = account_emails
        self.sender_combobox.grid(row=3, column=1, sticky='we', padx=5, pady=5)

        template_sender = self.template.get("sender", "").strip()
        if template_sender and template_sender in account_emails:
            self.sender_combobox.set(template_sender)
        elif account_emails:
            self.sender_combobox.set(account_emails[0])

        # Recipient
        ttk.Label(main_frame, text=self._("recipient") + ":", width=label_width, anchor='e').grid(row=4, column=0, sticky='e', padx=5, pady=5)
        self.to_entry = ttk.Entry(main_frame, width=entry_width)
        self.to_entry.grid(row=4, column=1, sticky='we', padx=5, pady=5)
        self.to_entry.insert(0, self.template.get("to", ""))
        self.to_entry.bind("<KeyRelease>", lambda e: self._format_email_entry(self.to_entry, e))

        # CC
        ttk.Label(main_frame, text=self._("cc") + ":", width=label_width, anchor='e').grid(row=5, column=0, sticky='e', padx=5, pady=5)
        self.cc_entry = ttk.Entry(main_frame, width=entry_width)
        self.cc_entry.grid(row=5, column=1, sticky='we', padx=5, pady=5)
        self.cc_entry.insert(0, self.template.get("cc", ""))
        self.cc_entry.bind("<KeyRelease>", lambda e: self._format_email_entry(self.cc_entry, e))

        # Subject
        ttk.Label(main_frame, text=self._("subject") + ":", width=label_width, anchor='e').grid(row=6, column=0, sticky='e', padx=5, pady=5)
        self.subject_entry = ttk.Entry(main_frame, width=entry_width)
        self.subject_entry.grid(row=6, column=1, sticky='we', padx=5, pady=5)
        self.subject_entry.insert(0, self.template.get("subject", ""))

        # Variables
        ttk.Label(main_frame, text=self._("variables") + ":", width=label_width, anchor='e').grid(row=7, column=0, sticky='ne', padx=5, pady=5)
        var_frame = ttk.Frame(main_frame)
        var_frame.grid(row=7, column=1, sticky='we', padx=5, pady=5)
        self.variables_description = ttk.Label(var_frame, text=self._("variable_description"))
        self.variables_description.pack(fill=tk.X, padx=5, pady=2)
        self.variables_entry = ttk.Entry(var_frame, width=entry_width)
        self.variables_entry.pack(fill=tk.X, padx=5, pady=2)

        # 有變數就插入現有值，否則顯示提示
        variables = ", ".join(self.template.get("variables", []))
        if variables:
            self.variables_entry.insert(0, variables)
        else:
            self._add_placeholder(self.variables_entry, "e.g. receiver, email, date")

        # Email Content
        self.body_frame = ttk.LabelFrame(main_frame, text=self._("email_content"))
        self.body_frame.grid(row=8, column=0, columnspan=2, sticky='nsew', padx=5, pady=5)

        info_label_frame = ttk.Frame(self.body_frame)
        info_label_frame.pack(fill=tk.X, padx=5, pady=5)
        self.content_info = ttk.Label(info_label_frame, text="HTML內容將在編輯時載入，點擊下方CKEditor按鈕編輯郵件內容", font=("TkDefaultFont", 9, "italic"), foreground="blue")
        self.content_info.pack(padx=5, pady=5)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=9, column=0, columnspan=2, pady=10, sticky='e')
        self.edit_html_btn = ttk.Button(button_frame, text="CKEditor", command=self._open_webview_editor)
        self.edit_html_btn.pack(side=tk.LEFT, padx=5)
        self.save_btn = ttk.Button(button_frame, text=self._("save"), command=self._save_template)
        self.save_btn.pack(side=tk.RIGHT, padx=5)

    def _open_webview_editor(self):
        # 獲取按鈕文本 - 多語言支持
        save_button_text = self._("save") if hasattr(self, '_') and callable(self._) else "Save"
        unsaved_warning = self._("unsaved_changes_warning") if hasattr(self, '_') and callable(self._) else "You have unsaved changes. Are you sure you want to leave?"

        with tempfile.NamedTemporaryFile(suffix='.html', delete=False) as f:
            body_content = self.template.get("body", "").strip()
            if not body_content:
                body_content = ""


            editor_html = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <title>HTML Editor</title>
                <script src="https://cdn.ckeditor.com/4.16.0/standard/ckeditor.js"></script>
                <style>
                    html, body {{ height: 100%; margin: 0; padding: 0; overflow: hidden; display: flex; flex-direction: column; }}
                    #editor-container {{ flex: 1; display: flex; flex-direction: column; }}
                    #controls {{ text-align: center; padding: 5px; }}
                    button {{ padding: 10px 20px; }}
                </style>
            </head>
            <body>
                <div id="editor-container">
                    <textarea id="editor">{body_content}</textarea>
                    <div id="controls">
                        <button onclick="saveContent()">{save_button_text}</button>
                    </div>
                </div>
            <script>
                // 初始化 CKEditor
                CKEDITOR.replace('editor');

                // 確認是否有未保存的變更
                let isContentSaved = false;

                // 綁定保存按鈕
                function saveContent() {{
                    var content = CKEDITOR.instances.editor.getData();
                    window.pywebview.api.save_content(content);
                    isContentSaved = true; // 標記為已保存
                }}

                // 窗口縮放時自動調整 CKEditor 大小
                function resizeEditor() {{
                    var editorHeight = window.innerHeight - 60;
                    CKEDITOR.instances.editor.resize('100%', editorHeight);
                }}

                window.addEventListener('resize', resizeEditor);
                window.addEventListener('load', resizeEditor);

                // 當關閉或刷新視窗時提示用戶
                window.addEventListener('beforeunload', function (e) {{
                    if (!isContentSaved) {{
                        e.preventDefault();
                        e.returnValue = '{unsaved_warning}';
                        return '{unsaved_warning}';
                    }}
                }});
            </script>
            </body>
            </html>
            """
            f.write(editor_html.encode('utf-8'))
            temp_path = f.name

        # 使用多語言支持的提示訊息
        starting_msg = self._("html_editor_starting") if hasattr(self, '_') and callable(self._) else "正在啟動 HTML 編輯器..."
        started_msg = self._("html_editor_started") if hasattr(self, '_') and callable(self._) else "HTML 編輯器已啟動，請在編輯完成後點擊「Save」按鈕"
        
        # 更新 UI 提示
        self.content_info.config(text=starting_msg, foreground="orange")
        self.window.update()

        # 終止之前的進程(如果有)
        if self.webview_process and self.webview_process.is_alive():
            self.webview_process.terminate()
            self.webview_process.join(1)
        
        # 獲取螢幕尺寸
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # 計算 CKEditor 位置（右側）
        webview_width = 800
        webview_height = 700
        webview_x = (screen_width // 2) + 50  # 中間偏右
        webview_y = (screen_height - webview_height) // 2
        
        # 確保座標為正數
        webview_x = max(0, webview_x)
        webview_y = max(0, webview_y)
        
        # 修改為使用帶位置參數的 run_webview
        self.webview_process = multiprocessing.Process(
            target=run_webview, 
    args=(self.content_queue, temp_path, webview_x, webview_y, webview_width, webview_height)  # ✅ 多传入4个参数

        )
        self.webview_process.daemon = True
        self.webview_process.start()

        # 更新 UI 提示
        self.content_info.config(text=started_msg, foreground="green")

        # 終止之前的進程(如果有)
        if self.webview_process and self.webview_process.is_alive():
            self.webview_process.terminate()
            self.webview_process.join(1)
            
        # 創建新進程，但不再嘗試設置位置
        self.webview_process = multiprocessing.Process(target=run_webview,  args=(self.content_queue, temp_path, webview_x, webview_y, webview_width, webview_height))
        self.webview_process.daemon = True
        self.webview_process.start()

        # 在進程啟動後，嘗試通過系統命令重新定位窗口
        # 這需要在主線程中等待幾秒，讓 webview 窗口有時間啟動

        # 更新 UI 提示
        self.content_info.config(text=started_msg, foreground="green")


    def _on_language_changed(self, event=None):
        # 更新界面上的語言
        self.window.title(self._("edit_template_title") if not self.is_new else self._("new_template_title"))
        self.name_label.config(text=self._("template_name") + ":")
        self.note_label.config(text=self._("Note") + ":")
        self.tag_label.config(text=self._("Tag") + ":")
        self.to_label.config(text=self._("recipient") + ":")
        self.cc_label.config(text=self._("cc") + ":")
        self.subject_label.config(text=self._("subject") + ":")
        self.variables_frame.config(text=self._("variables"))
        self.variables_description.config(text=self._("variable_description"))
        self.body_frame.config(text=self._("email_content"))
        self.save_btn.config(text=self._("save"))

    # 添加窗口關閉處理方法
    def _on_window_close(self):
        """處理視窗關閉事件"""
        # 無論如何都提示
        if messagebox.askyesno(
            self._("confirm"), 
            self._("unsaved_changes_warning"),
            parent=self.window
        ):
            self._close_webview_if_open()
            self.window.destroy()


    # 檢查是否有未保存的變更
    def has_unsaved_changes(self):
        """檢查是否有未保存的更改"""
        # 實現檢查未保存更改的邏輯
        # 例如比較原始模板與當前表單數據
        name = self.name_var.get().strip()
        to = self.to_entry.get().strip()
        cc = self.cc_entry.get().strip()
        sender = self.sender_var.get().strip()
        subject = self.subject_entry.get().strip()
        body = self.template.get('body', '').strip()
        variables_text = self.variables_entry.get().strip()
        note_en = self.note_entry.get().strip()
        tag_en = self.tag_entry.get().strip()
        
        # 如果是新模板，只要有任何內容就算有未保存的變更
        if self.is_new:
            return (name or to or cc or subject or body or variables_text or note_en or tag_en)
        
        # 如果是編輯現有模板，比較原始值與當前值
        if name != self.original_name:
            return True
        if to != self.template.get("to", ""):
            return True
        if cc != self.template.get("cc", ""):
            return True
        if sender != self.template.get("sender", ""):
            return True
        if subject != self.template.get("subject", ""):
            return True
        if note_en != self.template.get("note_en", ""):
            return True
        if tag_en != self.template.get("tag_en", ""):
            return True
        
        # 檢查變數列表
        original_vars = ", ".join(self.template.get("variables", []))
        if variables_text != original_vars:
            return True
        
        return False

    # 添加關閉 webview 的方法
    def _close_webview_if_open(self):
        """關閉 webview 如果它還開著"""
        if self.webview_process and self.webview_process.is_alive():
            try:
                # 檢查是否有內容已修改但未保存
                # 這個需要在 webview 中實現相關功能
                # 這裡簡單起見，直接發送關閉信號
                self.webview_process.terminate()
                self.webview_process.join(1)  # 等待最多1秒讓進程終止
            except Exception as e:
                print(f"關閉 webview 時出錯: {e}")

    def _save_template(self):
        """保存模板"""
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("Warning", self._("enter_template_name"))
            return
        
        to = self.to_entry.get().strip()
        cc = self.cc_entry.get().strip()
        sender = self.sender_var.get().strip()
        subject = self.subject_entry.get().strip()
        body = self.template.get('body', '').strip()
        variables_text = self.variables_entry.get().strip()
        variables = [var.strip() for var in variables_text.split(",")] if variables_text else []

        if not sender:  # 如果没有选择发送者
            messagebox.showwarning("Warning", self._("enter_sender"))
            return
        
        if not to:
            messagebox.showwarning("Warning", self._("enter_recipient"))
            return
        if not subject:
            messagebox.showwarning("Warning", self._("enter_subject"))
            return
        if not body:
            messagebox.showwarning("Warning", self._("enter_content"))
            return

        # 处理HTML内容中的图片
        try:
            from image_manager import ImageManager
            image_manager = ImageManager()
            processed_body = image_manager.process_html_content(body, name)
            
            # 更新图片路径后的body
            body = processed_body
        except Exception as e:
            print(f"处理图片时出错: {e}")
            messagebox.showwarning("Warning", f"Error processing images: {e}")

        # 更新模板数据
        template_data = {
            "name": name,
            "sender": sender,
            "to": to,
            "cc": cc,
            "subject": subject,
            "body": body,
            "variables": variables,
            "note_en": self.note_entry.get().strip(),
            "tag_en": self.tag_entry.get().strip()
        }

        # 如果是编辑模板，并且修改了模板名称，则检查是否存在同名模板
        if not self.is_new and name != self.original_name:
            existing_templates = self.template_manager.get_template_names_for_event(self.event_type)
            if name in existing_templates:
                if not messagebox.askyesno("Confirm", self._("template_name_exists").format(name=name)):
                    return
                
                # 重命名模板时，处理图片目录
                try:
                    image_manager.rename_template_image_dir(self.original_name, name)
                except Exception as e:
                    print(f"重命名图片目录时出错: {e}")
                
                self.template_manager.remove_template(self.event_type, self.original_name)

        # 保存模板
        self.template_manager.add_template(self.event_type, template_data)

        # 清理未使用的图片
        try:
            image_manager.cleanup_unused_images(body, name)
        except Exception as e:
            print(f"清理未使用图片时出错: {e}")

        # 刷新模板列表（如果有回调）
        if self.refresh_callback:
            self.refresh_callback()

        # 关闭窗口
        self.window.destroy()