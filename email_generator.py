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
        try:
            # 確保 Outlook 已啟動
            if not self.start_outlook_if_needed():
                return False

            # 每次重新建立 COM 連線
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            mail = outlook.CreateItem(0)

            # 獲取原始內容
            body = template.get("body", "")
            subject = template.get("subject", "")
            to = template.get("to", "")
            cc = template.get("cc", "")
            
            # 先處理 @{變數Email} 格式，在一般變數替換之前
            pattern = r'@\{([^{}]+)\}'
            
            def email_to_name(match):
                email_var = match.group(1)
                if email_var in variables:
                    email = variables[email_var]
                    # 從電子郵件中提取名字（取@符號前的部分）
                    name = email.split('@')[0] if '@' in email else email
                    # 將名字格式化，例如將"john.doe"轉為"John Doe"
                    name = name.replace('.', ' ').title()
                    return f"@{name}"
                return match.group(0)  # 如果變數不存在，保持原樣
            
            # 先應用 @{變數} 處理
            body = re.sub(pattern, email_to_name, body)
            
            # 準備變數字典
            variables_dict = {}
            for var_name, var_value in variables.items():
                variables_dict[var_name] = var_value
            
            # 定義完整替換函數 - 確保所有變數都被替換
            def replace_all_vars(text):
                if not text:
                    return ""
                
                # 使用正則表達式替換所有 {變數名} 格式的文本
                def replace_var(match):
                    var_full = match.group(0)  # 完整的 {變數名}
                    var_name = match.group(1)  # 提取變數名
                    
                    # 檢查是否為圖片變數
                    if var_name.startswith('inserted_photo') or var_name.startswith('inserted photo'):
                        # 如果是圖片變數但在主題或普通文本中，直接返回空字符串或描述性文本
                        return "[圖片]"
                        
                    # 非圖片變數，嘗試從變數字典中獲取值
                    if var_name in variables_dict:
                        return variables_dict[var_name]
                    return var_full  # 如果找不到對應的值，保持原樣
                    
                # 使用正則替換
                pattern = r'\{([^{}]+)\}'
                return re.sub(pattern, replace_var, text)
            
            # 對主題和收件人信息進行完整變數替換
            subject = replace_all_vars(subject)
            to = replace_all_vars(to)
            cc = replace_all_vars(cc)
            
            # 處理內文中的變數替換 - 圖片變數在主窗口中已經處理
            # 但一般文本變數需要在這裡處理
            if "<html>" in body.lower():
                # HTML 內容需要特殊處理，保留圖片變數的處理結果
                pattern = r'\{([^{}]+)\}'
                
                def replace_body_var(match):
                    var_name = match.group(1)
                    
                    # 檢查是否為圖片變數 - 保留原樣，因為可能已經被處理
                    if var_name.startswith('inserted_photo') or var_name.startswith('inserted photo'):
                        return match.group(0)
                    
                    # 非圖片變數
                    if var_name in variables_dict:
                        return variables_dict[var_name]
                    return match.group(0)
                
                # 替換內文中的變數
                body = re.sub(pattern, replace_body_var, body)
            else:
                # 純文本內容，直接替換所有變數
                body = replace_all_vars(body)
            
            # 設置郵件屬性 - 確保主題被完全設置
            mail.To = to
            mail.CC = cc
            
            # 特別處理主題行 - 確保完整顯示
            try:
                # 先嘗試直接設置主題
                mail.Subject = subject
                
                # 額外嘗試使用Property訪問器設置主題，以確保完整顯示
                mail._oleobj_.SetProperty(0x0037, subject)  # 0x0037是Subject屬性的MAPI標識符
            except Exception as e:
                print(f"設置主題時遇到錯誤: {e}")
                # 如果高級方法失敗，回退到基本方法
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
            elif any(tag in body.lower() for tag in ["<p>", "<div>", "<span>", "<table>", "<br", "<img"]):
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
                
                # 2. 如果上述方法失敗，嘗試找到對應賬戶
                if not sender_success:
                    accounts = self.get_outlook_accounts()
                    for account in accounts:
                        if account['email'].lower() == sender.lower() and account['account'] is not None:
                            try:
                                mail._oleobj_.Invoke(*(64209, 0, 8, 0, account['account']))
                                sender_success = True
                                print(f"成功設置寄件人賬戶: {sender}")
                                break
                            except Exception as e:
                                print(f"設置寄件人賬戶失敗: {e}")
                
                # 3. 使用其他方法嘗試
                if not sender_success:
                    try:
                        mail.SendUsingAccount = sender
                        sender_success = True
                        print(f"成功使用SendUsingAccount: {sender}")
                    except Exception as e:
                        print(f"SendUsingAccount設置失敗: {e}")

            # 額外驗證主題是否正確設置
            if mail.Subject != subject:
                try:
                    print(f"檢測到主題不匹配，重新設置主題: {subject}")
                    mail.Subject = subject
                except:
                    pass

            # 如果設置寄件人失敗，給用戶提示
            if sender and not sender_success:
                msgbox.showinfo(
                    "寄件人設置信息", 
                    f"無法自動設置寄件人為「{sender}」\n\n郵件將使用預設寄件人開啟，請在郵件窗口中手動選擇寄件人。"
                )

            # 創建並顯示郵件前，再次驗證主題設置
            try:
                # 設置一些額外的屬性，以確保主題完整顯示
                mail._oleobj_.SetProperty(0x0037, subject)  # 0x0037是Subject屬性的MAPI標識符
                mail._oleobj_.SetProperty(0x0070, subject)  # 0x0070是ConversationTopic屬性的MAPI標識符
            except:
                pass

            # 顯示郵件
            mail.Display(False)
            
            # 生成後立即檢查主題並提示用戶
            if len(subject) > 60:  # 主題較長時提示用戶
                msgbox.showinfo(
                    "主題完整性提示", 
                    f"您的郵件主題較長，在Outlook中可能不會完整顯示。\n\n完整主題是:\n{subject}\n\n如果您看到主題顯示不全，可以在郵件窗口中手動修改。"
                )

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
