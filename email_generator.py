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
        """獲取 Outlook 中配置的所有郵件賬戶，包括企業郵件"""
        accounts = []
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # 嘗試獲取默認帳戶
            try:
                default_account = namespace.Accounts.Item(1)
                accounts.append({
                    "name": default_account.DisplayName,
                    "email": default_account.SmtpAddress,
                    "account": default_account
                })
                print(f"找到默認帳戶: {default_account.DisplayName} <{default_account.SmtpAddress}>")
            except Exception as e:
                print(f"獲取默認帳戶時出錯: {e}")
            
            # 使用循環方式獲取所有帳戶
            for i in range(1, namespace.Accounts.Count + 1):
                try:
                    account = namespace.Accounts.Item(i)
                    
                    # 避免重複添加
                    if not any(acc["email"] == account.SmtpAddress for acc in accounts):
                        accounts.append({
                            "name": account.DisplayName,
                            "email": account.SmtpAddress,
                            "account": account
                        })
                        print(f"找到帳戶: {account.DisplayName} <{account.SmtpAddress}>")
                except Exception as e:
                    print(f"獲取帳戶 {i} 時出錯: {e}")
            
            # 如果無法獲取任何帳戶，嘗試以不同方式獲取
            if not accounts:
                try:
                    # 嘗試直接從當前使用者個人資料獲取
                    profile = namespace.CurrentUser
                    if profile:
                        accounts.append({
                            "name": profile.Name,
                            "email": profile.Address,
                            "account": None  # 這裡無法獲取完整帳戶對象
                        })
                        print(f"從個人資料獲取: {profile.Name} <{profile.Address}>")
                except Exception as e:
                    print(f"獲取當前使用者資料時出錯: {e}")
            
            # 嘗試獲取委派和共享郵箱
            try:
                # 遍歷所有資料夾，嘗試找到共享郵箱
                folders = namespace.Folders
                for i in range(1, folders.Count + 1):
                    try:
                        folder = folders.Item(i)
                        email = folder.Name  # 對於共享郵箱，通常名稱就是郵件地址
                        if "@" in email and not any(acc["email"] == email for acc in accounts):
                            accounts.append({
                                "name": folder.Name,
                                "email": email,
                                "account": None  # 無法獲取完整帳戶對象
                            })
                            print(f"找到共享郵箱: {email}")
                    except Exception as e:
                        print(f"處理資料夾 {i} 時出錯: {e}")
            except Exception as e:
                print(f"獲取共享郵箱時出錯: {e}")
                
        except Exception as e:
            print(f"獲取 Outlook 賬戶時出錯: {e}")
        
        return accounts

    def generate_email(self, template: Dict[str, Any], variables: Dict[str, str], sender: Optional[str] = None, signature_option: str = "<Default>") -> bool:
        """
        生成並打開 Outlook 郵件
        
        Args:
            template: 模板數據
            variables: 變數值
            sender: 寄件人郵件地址
            signature_option: 簽名檔選項 ("<Default>", "<None>", 或特定簽名檔名稱)
        """
        try:
            # 確保 Outlook 已啟動
            if not self.start_outlook_if_needed():
                return False

            # 每次重新建立 COM 連線
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            mail = outlook.CreateItem(0)

            # 變數替換
            variables_dict = {}
            for var_name, var_value in variables.items():
                variables_dict[f"{{{var_name}}}"] = var_value
            
            # 替換函數，處理缺失變數
            def replace_vars(text):
                if not text:
                    return ""
                for var, value in variables_dict.items():
                    text = text.replace(var, value)
                return text
            
            # 替換主題、正文、收件人和抄送中的變數
            subject = replace_vars(template.get("subject", ""))
            body = replace_vars(template.get("body", ""))
            to = replace_vars(template.get("to", ""))
            cc = replace_vars(template.get("cc", ""))
            
            # @email 替換
            pattern = r'@{([^{}]+)}'
            body = re.sub(pattern, lambda m: f"@{m.group(1)}", body)
            
            # 設置郵件屬性
            mail.To = to
            mail.CC = cc
            mail.Subject = subject
            
            # 處理簽名檔設置
            use_signature = True
            if signature_option == "<None>":
                use_signature = False
            elif signature_option != "<Default>" and isinstance(signature_option, str):
                # 嘗試使用指定的簽名檔
                try:
                    import os
                    appdata = os.environ.get('APPDATA', '')
                    signature_path = os.path.join(appdata, 'Microsoft', 'Signatures', signature_option)
                    
                    # 優先嘗試 HTML 格式
                    html_path = f"{signature_path}.htm"
                    if os.path.exists(html_path):
                        with open(html_path, 'r', encoding='utf-8') as f:
                            signature_html = f.read()
                        # 找到簽名檔的 <body> 部分
                        body_match = re.search(r'<body[^>]*>(.*?)</body>', signature_html, re.DOTALL)
                        if body_match:
                            signature_content = body_match.group(1)
                            # 在郵件 HTML 結尾前添加簽名檔
                            if "<html>" in body.lower():
                                body = body.replace('</body>', signature_content + '</body>')
                            else:
                                body = f"<html><body>{body}{signature_content}</body></html>"
                    
                    # 設置 use_signature = False 避免自動添加默認簽名檔
                    use_signature = False
                except Exception as e:
                    print(f"使用指定簽名檔時出錯: {e}")
            
            # 設置郵件正文
            if "<html>" in body.lower():
                mail.HTMLBody = body
                if not use_signature:
                    try:
                        mail._oleobj_.Invoke(*(2381, 0, 8, 0, False))  # Don't use signature
                    except:
                        print("禁用簽名檔失敗")
            elif any(tag in body.lower() for tag in ["<p>", "<div>", "<span>", "<table>", "<br"]):
                mail.HTMLBody = f"<html><body>{body}</body></html>"
                if not use_signature:
                    try:
                        mail._oleobj_.Invoke(*(2381, 0, 8, 0, False))  # Don't use signature
                    except:
                        print("禁用簽名檔失敗")
            else:
                mail.Body = body
                if not use_signature:
                    try:
                        mail._oleobj_.Invoke(*(2381, 0, 8, 0, False))  # Don't use signature
                    except:
                        print("禁用簽名檔失敗")

            # 檢查是否有圖片附件需要添加
            if "cid:" in body:
                self._add_image_attachments(mail, template.get("name", ""))

            # 設置寄件人（如果指定）
            sender_success = False
            if sender:
                print(f"嘗試設置寄件人: {sender}")
                
                # 1. 直接嘗試設置 SendOnBehalfOfName (最常用於團隊郵箱)
                try:
                    mail.SentOnBehalfOfName = sender
                    sender_success = True
                    print(f"成功設置「代表」發送: {sender}")
                except Exception as e:
                    print(f"設置「代表」發送失敗: {e}")
                
                # 2. 如果上述方法失敗，嘗試找到對應帳戶
                if not sender_success:
                    accounts = self.get_outlook_accounts()
                    for account in accounts:
                        if account['email'].lower() == sender.lower() and account['account'] is not None:
                            try:
                                mail._oleobj_.Invoke(*(64209, 0, 8, 0, account['account']))
                                sender_success = True
                                print(f"成功設置寄件人帳戶: {sender}")
                                break
                            except Exception as e:
                                print(f"設置寄件人帳戶失敗: {e}")
                
                # 3. 使用其他方法嘗試
                if not sender_success:
                    try:
                        mail.SendUsingAccount = sender
                        sender_success = True
                        print(f"成功使用SendUsingAccount: {sender}")
                    except Exception as e:
                        print(f"SendUsingAccount設置失敗: {e}")

            # 如果設置寄件人失敗，給用戶提示
            if sender and not sender_success:
                msgbox.showinfo(
                    "寄件人設置信息", 
                    f"無法自動設置寄件人為「{sender}」\n\n郵件將使用預設寄件人開啟，請在郵件窗口中手動選擇寄件人。"
                )

            # 顯示郵件
            mail.Display(False)

            # 即使設置寄件人失敗，我們也認為郵件生成成功
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

    def get_outlook_signatures(self) -> List[str]:
        """獲取 Outlook 中可用的簽名檔列表"""
        signatures = ["<Default>", "<None>"]  # 基本選項
        
        # 嘗試查找簽名檔路徑
        try:
            import os
            import platform
            
            if platform.system() == 'Windows':
                # Windows 上 Outlook 簽名檔的默認位置
                appdata = os.environ.get('APPDATA', '')
                signature_path = os.path.join(appdata, 'Microsoft', 'Signatures')
                
                if os.path.exists(signature_path):
                    # 獲取所有 HTML 和 RTF 簽名檔
                    for file in os.listdir(signature_path):
                        base, ext = os.path.splitext(file)
                        if ext.lower() in ['.htm', '.html', '.rtf']:
                            if base not in signatures:
                                signatures.append(base)
        except Exception as e:
            print(f"獲取簽名檔時出錯: {e}")
        
        return signatures

    def _find_and_set_account(self, mail, namespace, email_address):
        """嘗試查找並設置對應的帳戶"""
        # 遍歷所有資料夾，嘗試找到匹配的共享郵箱
        for i in range(1, namespace.Folders.Count + 1):
            folder = namespace.Folders.Item(i)
            if folder.Name.lower() == email_address.lower():
                mail.SendUsingAccount = folder.Name
                return True
        
        # 如果無法找到匹配資料夾，嘗試設置 SentOnBehalfOfName
        mail.SentOnBehalfOfName = email_address
        return True

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