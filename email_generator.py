import win32com.client
import subprocess
import psutil
import time
import re
import os
import tkinter.messagebox as msgbox
from typing import Dict, Any, Optional, List
from image_manager import ImageManager

class EmailGenerator:
    """處理 Outlook 電子郵件生成的類"""

    def __init__(self):
        """初始化電子郵件生成器"""
        self.image_manager = ImageManager()

    def is_outlook_running(self) -> bool:
        """檢查 Outlook 是否正在運行"""
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and 'OUTLOOK.EXE' in proc.info['name'].upper():
                return True
        return False

    def start_outlook_if_needed(self, language_manager=None):
        """檢查 Outlook 是否正在運行，如果沒啟動提示用戶"""
        if not self.is_outlook_running():
            # 根據語言
            title = "Outlook Not Running"
            msg_text = "Outlook is not running. Please open Outlook first before generating the email."
            if language_manager:
                title = language_manager.get_text("outlook_not_running_title")
                msg_text = language_manager.get_text("outlook_not_running_msg")
            
            print(msg_text)
            msgbox.showwarning(title, msg_text)
            return False
        return True

    def is_outlook_available(self) -> bool:
        """檢查 Outlook 是否可用"""
        try:
            _ = win32com.client.Dispatch("Outlook.Application")
            return True
        except Exception:
            return False

    def get_outlook_accounts(self) -> List[Dict[str, Any]]:
        """獲取 Outlook 中配置的所有郵件賬戶"""
        accounts = []
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            for account in namespace.Accounts:
                # 保存完整賬戶對象和郵件地址
                accounts.append({
                    "name": account.DisplayName,
                    "email": account.SmtpAddress,
                    "account": account  # 保存完整賬戶對象以便後續使用
                })
                print(f"找到賬戶: {account.DisplayName} <{account.SmtpAddress}>")
        except Exception as e:
            print(f"獲取 Outlook 賬戶時出錯: {e}")
        return accounts

    def generate_email(self, template: Dict[str, Any], variables: Dict[str, str], sender: Optional[str] = None) -> bool:
        """生成並打開 Outlook 郵件"""
        try:
            # 確保 Outlook 已啟動
            if not self.start_outlook_if_needed():
                return False

            # 每次重新建立 COM 連線
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            mail = outlook.CreateItem(0)

            # 設置寄件人（如果指定）
            if sender:
                print(f"嘗試設置寄件人: {sender}")
                # 獲取所有賬戶
                accounts = self.get_outlook_accounts()
                sender_set = False
                
                for account in accounts:
                    print(f"比較賬戶: {account['email']} 與 {sender}")
                    if account['email'].lower() == sender.lower():
                        try:
                            # 使用更可靠的方法設置寄件人
                            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account['account']))
                            print(f"成功設置寄件人為: {sender}")
                            sender_set = True
                            break
                        except Exception as e:
                            print(f"設置寄件人時出錯: {e}")
                            error_msg = f"設置寄件人時出錯: {str(e)}"
                            msgbox.showerror("寄件人設置錯誤", error_msg)
                            return False
                
                if not sender_set:
                    error_msg = f"找不到匹配的寄件人賬戶: {sender}"
                    print(error_msg)
                    msgbox.showerror("寄件人錯誤", error_msg)
                    return False

            # 設置收件人和抄送
            mail.To = template.get("to", "")
            mail.CC = template.get("cc", "")

            # 主題變數替換
            subject = template.get("subject", "")
            for var_name, var_value in variables.items():
                subject = subject.replace(f"{{{var_name}}}", var_value)
            mail.Subject = subject

            # 正文變數替換
            body = template.get("body", "")
            for var_name, var_value in variables.items():
                body = body.replace(f"{{{var_name}}}", var_value)

            # @email 替換
            pattern = r'@{([^{}]+)}'
            body = re.sub(pattern, lambda m: f"@{m.group(1)}", body)

            # 設置收件人和抄送 (需要做變數替換)
            to = template.get("to", "")
            cc = template.get("cc", "")
            for var_name, var_value in variables.items():
                to = to.replace(f"{{{var_name}}}", var_value)
                cc = cc.replace(f"{{{var_name}}}", var_value)
            mail.To = to
            mail.CC = cc


            # 設定正文
            if "<html>" in body.lower():
                mail.HTMLBody = body
            elif any(tag in body.lower() for tag in ["<p>", "<div>", "<span>", "<table>", "<br"]):
                mail.HTMLBody = f"<html><body>{body}</body></html>"
            else:
                mail.Body = body

            # 檢查是否有圖片附件需要添加
            if "cid:" in body:
                self._add_image_attachments(mail, template.get("name", ""))

            # 顯示郵件
            mail.Display(False)
            return True

        except Exception as e:
            error_message = str(e)
            print(f"生成郵件時出錯: {error_message}")
            
            # 處理特定的錯誤類型
            if "dialog box is open" in error_message.lower():
                user_friendly_message = "Outlook 無法生成郵件，因為有對話框已開啟。\n請關閉所有 Outlook 對話框後再試。"
            else:
                user_friendly_message = f"生成郵件時出錯:\n{error_message}"
                
            # 顯示錯誤彈窗
            msgbox.showerror("郵件生成錯誤", user_friendly_message)
            return False

    def _add_image_attachments(self, mail, template_name):
        """添加圖片附件到郵件並設置內容ID
        
        Args:
            mail: Outlook郵件對象
            template_name (str): 模板名稱
        """
        # 獲取模板的所有圖片
        image_paths = self.image_manager.get_image_paths(template_name)
        
        # 添加圖片作為附件
        for image_path in image_paths:
            filename = os.path.basename(image_path)
            try:
                attachment = mail.Attachments.Add(image_path)
                # 設置內容ID，與HTML中的src="cid:filename"對應
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", filename)
            except Exception as e:
                error_msg = f"添加圖片附件時出錯: {e}"
                print(error_msg)
                msgbox.showerror("圖片附件錯誤", error_msg)