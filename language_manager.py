from typing import Dict, Any, List
import os
import json
from db_manager import DatabaseManager

class LanguageManager:
    """管理應用程序多語言支持的類"""
    
    def __init__(self, default_language='zh_TW', db_manager: DatabaseManager = None):
        self.db_manager = db_manager if db_manager else DatabaseManager()
        self.default_language = default_language
        self.current_language = default_language

        # 初始化语言和翻译
        self._init_languages_and_translations()

        # 加载翻译（包含app_info）
        self.translations = self._load_translations()

        # 加载用户偏好设置
        preferred_language = self.load_user_preference()
        if preferred_language:
            self.current_language = preferred_language

    def reload_translations(self):
        """重新載入翻譯"""
        self.translations = self._load_translations()
        print("🔄 Translations reloaded!")

    def _init_languages_and_translations(self):
        """初始化语言和翻译数据"""
        self._ensure_language('zh_TW', '繁體中文')
        self._ensure_language('en_US', 'English')
        self._ensure_default_translations()

    def _ensure_language(self, code: str, description: str):
        languages = self.db_manager.get_languages()
        if not any(lang['code'] == code for lang in languages):
            self.db_manager.add_language(code, description)

    def _ensure_default_translations(self):
        zh_translations = self.db_manager.get_translations('zh_TW')
        en_translations = self.db_manager.get_translations('en_US')
        if len(zh_translations) < 10 or len(en_translations) < 10:
            self._add_default_translations()

    
    def _add_default_translations(self):
        """添加默认翻译数据"""
        # 创建默认翻译数据
        default_translations = {
            'zh_TW': {
                # 通用
                'app_title': 'Outlook 模板助手',
                'close': '關閉',
                'save': '保存',
                'cancel': '取消',
                'confirm': '確認',
                'warning': '警告',
                'error': '錯誤',
                'success': '成功',
                'information': '訊息',
                'language': '語言',
                'english': 'English',
                'traditional_chinese': '繁體中文',
                'language_changed': '語言設置已更改為繁體中文。',
                'language_changed_en': '語言設置已更改為英文。',
                
                # 主窗口
                'file': '文件',
                'import_templates': '匯入模板',
                'export_templates': '匯出模板',
                'exit': '退出',
                'template': '模板',
                'template_management': '模板管理',
                'help': '幫助',
                'about': '關於',
                'select_event_type': '選擇事件類型',
                'select_template': '選擇模板',
                'fill_variables': '填寫變數',
                'generate_email': '生成 Email',
                'ready': '就緒',
                'outlook_unavailable': '警告: 無法連接到 Outlook',
                'outlook_unavailable_msg': '無法連接到 Outlook。請確保 Outlook 已安裝並正確配置。',
                'sender': '寄件人',
                
                # 模板管理窗口
                'template_management_title': '模板管理',
                'event_type': '事件類型',
                'add': '新增',
                'delete': '刪除',
                'search': '搜索',
                'clear': '清除',
                'template_list': '模板列表',
                'add_template': '新增模板',
                'edit_template': '編輯模板',
                'delete_template': '刪除模板',
                'template_preview': '模板預覽',
                'name': '名稱',
                'recipient': '收件人',
                'cc': '抄送',
                'subject': '主題',
                'variables': '變數',
                'content': '內容',
                
                # 编辑模板窗口
                'edit_template_title': '編輯模板',
                'new_template_title': '新增模板',
                'template_name': '模板名稱',
                'variable_description': '輸入變數名稱，用逗號分隔。變數在模板中使用 {變數名} 格式。',
                'email_content': '電子郵件內容',
                'extract_variables': '提取變數',
                'html_template': 'HTML模板',
                'html_format_support': '支持 HTML 格式，可使用上方按鈕插入常用標籤',
                
                # 消息
                'confirm_replace_content': '這將替換當前內容，確定要繼續嗎？',
                'enter_event_type': '請輸入新的事件類型名稱:',
                'event_type_exists': "事件類型 '{name}' 已存在",
                'confirm_delete_event': "確定要刪除事件類型 '{name}' 及其所有模板嗎？",
                'please_select_event_type': '請選擇事件類型',
                'please_select_template': '請選擇模板',
                'template_not_found': '無法找到所選模板',
                'confirm_delete_template': "確定要刪除模板 '{name}' 嗎？",
                'enter_template_name': '請輸入模板名稱',
                'enter_recipient': '請輸入收件人',
                'enter_subject': '請輸入主題',
                'enter_content': '請輸入郵件內容',
                'template_name_exists': "模板名稱 '{name}' 已存在，是否要覆蓋？",
                'import_success': '模板導入成功',
                'import_error': '導入模板失敗，請確保文件格式正確',
                'export_success': '模板導出成功',
                'export_error': '導出模板失敗: {error}',
                'please_fill_variables': '請填寫以下變數: {vars}',
                'email_generated': '已生成郵件: {name}',
                'email_generation_failed': '生成郵件失敗',
                'language_change_restart': '語言設置已更改。請重啟應用程序以完全應用更改。',
                "outlook_not_running_title": "Outlook 未啟動",
                "outlook_not_running_msg": "Outlook 未啟動，請先手動開啟 Outlook 再進行發信。",
                'refresh': '刷新',
                'refreshed': '已刷新',
                'backup_reminder': '數據庫備份提醒',
                'backup_message': '請定期備份您的數據庫文件，以防數據丢失。\n\n數據庫文件位置:\n{db_path}\n\n建議每週備份一次，複製此文件到安全位置。',
                'backup_now': '立即備份',
                'backup_export': '導出爲JSON',
                'backup_success': '備份成功',
                'backup_folder': '備份文件夾',
                
                # HTML编辑器
                'font': '字型',
                'size': '大小',
                'color': '顏色',
                'background': '背景',
                'bold': '粗體',
                'italic': '斜體',
                'underline': '底線',
                'align_left': '靠左對齊',
                'align_center': '置中對齊',
                'align_right': '靠右對齊',
                'bullet_list': '項目符號列表',
                'number_list': '編號列表',
                'insert_table': '插入表格',
                'insert_image': '插入圖片',
                'insert_link': '插入連結',
                'paragraph': '段落',
                'link': '超連結',
                'list': '列表',
                'table': '表格',
                'rows': '行數',
                'columns': '列數',
                'insert': '插入',
                'preview': '預覽',
                'image': '圖片',
                'file': '檔案',
                'all': '所有',
                'cut': '剪切',
                'copy': '複製',
                'paste': '貼上',
                'select_all': '全選',
                'ready': '就緒',
                'preview_not_available': '預覽功能不可用，請安裝 tkhtmlview 模組',
                'html_pasted': 'HTML 內容已貼上',
                'paste_error': '貼上內容時出錯',
                'image_insert_error': '插入圖片時出錯',
                'rich_text': '富文本',
                'edit': '編輯',
                'all': '所有',
                'file': '檔案',
                'image': '圖片',
                
                # 版本和说明
                'version': '版本',
                'about_description': '這是一個幫助用戶快速生成 Outlook 郵件的工具，\n可以通過模板管理來創建和編輯不同類型的郵件模板。\n開發者：XXX\n聯繫方式：[Github](https://github.com/Ash0Lam)',
                
                # 文件类型
                'json_file': 'JSON 文件',
                'all_files': '所有文件'
            },
            'en_US': {
                # General
                'app_title': 'Outlook Template Assistant',
                'close': 'Close',
                'save': 'Save',
                'cancel': 'Cancel',
                'confirm': 'Confirm',
                'warning': 'Warning',
                'error': 'Error',
                'success': 'Success',
                'information': 'Information',
                'language': 'Language',
                'english': 'English',
                'traditional_chinese': 'Traditional Chinese',
                'language_changed': 'Language has been changed to English.',
                'language_changed_zh': 'Language has been changed to Traditional Chinese.',
                
                # Main Window
                'file': 'File',
                'import_templates': 'Import Templates',
                'export_templates': 'Export Templates',
                'exit': 'Exit',
                'template': 'Template',
                'template_management': 'Template Management',
                'help': 'Help',
                'about': 'About',
                'select_event_type': 'Select Event Type',
                'select_template': 'Select Template',
                'fill_variables': 'Fill Variables',
                'generate_email': 'Generate Email',
                'ready': 'Ready',
                'outlook_unavailable': 'Warning: Unable to connect to Outlook',
                'outlook_unavailable_msg': 'Unable to connect to Outlook. Please ensure Outlook is installed and properly configured.',
                'sender': 'Sender',
                
                # Template Management Window
                'template_management_title': 'Template Management',
                'event_type': 'Event Type',
                'add': 'Add',
                'delete': 'Delete',
                'search': 'Search',
                'clear': 'Clear',
                'template_list': 'Template List',
                'add_template': 'Add Template',
                'edit_template': 'Edit Template',
                'delete_template': 'Delete Template',
                'template_preview': 'Template Preview',
                'name': 'Name',
                'recipient': 'Recipient',
                'cc': 'CC',
                'subject': 'Subject',
                'variables': 'Variables',
                'content': 'Content',
                'cut': 'Cut',
                'copy': 'Copy',
                'paste': 'Paste',
                'select_all': 'Select All',
                'ready': 'Ready',
                'preview_not_available': 'Preview not available, please install tkhtmlview module',
                'html_pasted': 'HTML content pasted',
                'paste_error': 'Error pasting content',
                'image_insert_error': 'Error inserting image',
                'rich_text': 'Rich Text',
                'edit': 'Edit',
                'all': 'All',
                'file': 'File',
                'image': 'Image',
                
                # Edit Template Window
                'edit_template_title': 'Edit Template',
                'new_template_title': 'New Template',
                'template_name': 'Template Name',
                'variable_description': 'Enter variable names, separated by commas. Variables are used in the template in the format {variable_name}.',
                'email_content': 'Email Content',
                'extract_variables': 'Extract Variables',
                'html_template': 'HTML Template',
                'html_format_support': 'HTML format is supported. Use the buttons above to insert common tags.',
                
                # Messages
                'confirm_replace_content': 'This will replace the current content. Are you sure you want to continue?',
                'enter_event_type': 'Please enter a new event type name:',
                'event_type_exists': "Event type '{name}' already exists",
                'confirm_delete_event': "Are you sure you want to delete event type '{name}' and all its templates?",
                'please_select_event_type': 'Please select an event type',
                'please_select_template': 'Please select a template',
                'template_not_found': 'Cannot find the selected template',
                'confirm_delete_template': "Are you sure you want to delete template '{name}'?",
                'enter_template_name': 'Please enter a template name',
                'enter_recipient': 'Please enter a recipient',
                'enter_subject': 'Please enter a subject',
                'enter_content': 'Please enter email content',
                'template_name_exists': "Template name '{name}' already exists. Do you want to override it?",
                'import_success': 'Templates imported successfully',
                'import_error': 'Failed to import templates. Please ensure the file format is correct.',
                'export_success': 'Templates exported successfully',
                'export_error': 'Failed to export templates: {error}',
                'please_fill_variables': 'Please fill in the following variables: {vars}',
                'email_generated': 'Email generated: {name}',
                'email_generation_failed': 'Failed to generate email',
                'language_change_restart': 'Language setting has been changed. Please restart the application to fully apply the changes.',
                "outlook_not_running_title": "Outlook Not Running",
                "outlook_not_running_msg": "Outlook is not running. Please open Outlook first before generating the email.",
                'refresh': 'Refresh',
                'refreshed': 'Refreshed',
                'backup_reminder': 'Database Backup Reminder',
                'backup_message': 'Please regularly backup your database file to prevent data loss.\n\nDatabase file location:\n{db_path}\n\nWe recommend backing up weekly by copying this file to a safe location.',
                'backup_now': 'Backup Now',
                'backup_export': 'Export as JSON',
                'backup_success': 'Backup Successful',
                'backup_folder': 'Backup Folder',
                
                # HTML Editor
                'font': 'Font',
                'size': 'Size',
                'color': 'Color',
                'background': 'Background',
                'bold': 'Bold',
                'italic': 'Italic',
                'underline': 'Underline',
                'align_left': 'Align Left',
                'align_center': 'Center',
                'align_right': 'Align Right',
                'bullet_list': 'Bullet List',
                'number_list': 'Numbered List',
                'insert_table': 'Insert Table',
                'insert_image': 'Insert Image',
                'insert_link': 'Insert Link',
                'paragraph': 'Paragraph',
                'link': 'Link',
                'list': 'List',
                'table': 'Table',
                'rows': 'Rows',
                'columns': 'Columns',
                'insert': 'Insert',
                'preview': 'Preview',
                'image': 'Image',
                'file': 'File',
                'all': 'All',
                
                # Version and Description
                'version': 'Version',
                'about_description': 'This tool helps users quickly generate Outlook emails.\nWith the template management feature, you can easily create and edit various types of email templates.\nDeveloper: Ash\nContact: [Github](https://github.com/Ash0Lam)',
                
                # File Types
                'json_file': 'JSON File',
                'all_files': 'All Files'
            }
        }
        
        # 添加翻译数据
        for lang_code, translations in default_translations.items():
            for key, text in translations.items():
                self.db_manager.add_translation(lang_code, key, text)
    
    def _load_translations(self) -> Dict[str, Dict[str, str]]:
        translations = {}
        languages = self.db_manager.get_languages()
        for lang in languages:
            lang_code = lang['code']
            translations[lang_code] = {}
            translations[lang_code].update(self.db_manager.get_translations(lang_code))
            translations[lang_code].update(self.db_manager.get_app_info_translations(lang_code))  # 加载 app_info
        return translations
    
    def set_language(self, language_code: str) -> None:
        """設置當前語言
        
        Args:
            language_code (str): 語言代碼 ('zh_TW' 或 'en_US')
        """
        if language_code in self.translations:
            self.current_language = language_code
    
    def get_text(self, key: str) -> str:
        """獲取指定鍵的翻譯文本
        
        Args:
            key (str): 翻譯鍵
            
        Returns:
            str: 翻譯後的文本，如果未找到則返回鍵本身
        """
        # 嘗試從當前語言獲取翻譯
        if key in self.translations.get(self.current_language, {}):
            return self.translations[self.current_language][key]
        
        # 如果當前語言沒有該翻譯，嘗試從默認語言獲取
        if key in self.translations.get(self.default_language, {}):
            return self.translations[self.default_language][key]
        
        # 如果都沒有找到，返回鍵本身
        return key
    
    def get_available_languages(self) -> Dict[str, str]:
        """獲取可用語言列表
        
        Returns:
            Dict[str, str]: 語言代碼到語言名稱的映射
        """
        return {
            'zh_TW': self.get_text('traditional_chinese'),
            'en_US': self.get_text('english')
        }
    
    def save_user_preference(self, language_code: str) -> None:
        """保存用戶語言偏好設置
        
        Args:
            language_code (str): 語言代碼
        """
        self.db_manager.save_setting('language', language_code)
    
    def load_user_preference(self) -> str:
        """加載用戶語言偏好設置
        
        Returns:
            str: 語言代碼，如果未找到則返回默認語言
        """
        language = self.db_manager.get_setting('language')
        if language and language in self.translations:
            return language
        return self.default_language
    
    def migrate_from_json(self, language_dir: str = 'data/language') -> bool:
        """从JSON文件迁移语言数据
        
        Args:
            language_dir (str): 语言JSON文件目录
            
        Returns:
            bool: 是否迁移成功
        """
        try:
            # 添加默认语言
            self._ensure_language('zh_TW', '繁體中文')
            self._ensure_language('en_US', 'English')
            
            # 迁移翻译数据
            success = False
            for lang_code in ['zh_TW', 'en_US']:
                lang_file = os.path.join(language_dir, f'{lang_code}.json')
                if os.path.exists(lang_file):
                    with open(lang_file, 'r', encoding='utf-8') as f:
                        translations = json.load(f)
                    for key, text in translations.items():
                        self.db_manager.add_translation(lang_code, key, text)
                    success = True
            
            # 如果没有成功迁移任何文件，则添加默认翻译
            if not success:
                self._add_default_translations()
            
            # 重新加载翻译
            self.translations = self._load_translations()
            
            return True
        except Exception as e:
            print(f"迁移失败: {e}")
            return False