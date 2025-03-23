import sqlite3
import os
import json
from typing import Dict, List, Any, Optional

class DatabaseManager:
    """数据库管理类，提供SQLite数据库操作功能"""
    
    def __init__(self, db_file: str = 'data/app.db'):
        """初始化数据库管理器
        
        Args:
            db_file (str): 数据库文件路径
        """
        # 确保数据目录存在
        os.makedirs(os.path.dirname(db_file), exist_ok=True)
        
        self.db_file = db_file
        self.conn = None
        self.init_database()
    
    def get_connection(self):
        """获取数据库连接，每个线程使用自己的连接"""
        conn = sqlite3.connect(self.db_file, check_same_thread=False)  # Enable multi-threading support for SQLite
        conn.row_factory = sqlite3.Row  # Return rows as dictionaries
        conn.execute("PRAGMA foreign_keys = ON")  # Enable foreign key support
        return conn
    
    def close_connection(self):
        """关闭数据库连接"""
        if self.conn:
            self.conn.close()
            self.conn = None
    
    def init_database(self):
        """初始化数据库表结构"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # 創建設置表
        cursor.execute('''CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )''')
        
        # 創建語言表
        cursor.execute('''CREATE TABLE IF NOT EXISTS languages (
            code TEXT PRIMARY KEY,
            description TEXT NOT NULL
        )''')
        
        # 創建翻譯表
        cursor.execute('''CREATE TABLE IF NOT EXISTS translations (
            language_code TEXT,
            key TEXT,
            text TEXT NOT NULL,
            PRIMARY KEY (language_code, key),
            FOREIGN KEY (language_code) REFERENCES languages(code)
        )''')
        
        # 創建事件類型表
        cursor.execute('''CREATE TABLE IF NOT EXISTS event_types (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE
        )''')
        
        # 創建 app_info_translations 表
        cursor.execute('''CREATE TABLE IF NOT EXISTS app_info_translations (
            key TEXT,
            language_code TEXT,
            value TEXT,
            PRIMARY KEY (key, language_code),
            FOREIGN KEY (language_code) REFERENCES languages(code)
        )''')

        # 檢查 templates 表是否已存在 sender 欄位
        cursor.execute("PRAGMA table_info(templates)")
        table_info = cursor.fetchall()
        columns = [column['name'] for column in table_info]
        
        if not table_info:  # 如果表不存在，創建帶有 sender 欄位的表
            cursor.execute('''CREATE TABLE IF NOT EXISTS templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                event_type_id INTEGER,
                name TEXT NOT NULL,
                recipient TEXT,
                cc TEXT,
                subject TEXT,
                body TEXT,
                note_en TEXT,
                tag_en TEXT,
                sender TEXT,
                FOREIGN KEY (event_type_id) REFERENCES event_types(id) ON DELETE CASCADE,
                UNIQUE (event_type_id, name)
            )''')
        elif 'sender' not in columns:  # 如果表已存在但缺少 sender 欄位，添加它
            cursor.execute('ALTER TABLE templates ADD COLUMN sender TEXT')
        
        # 創建模板變量表
        cursor.execute('''CREATE TABLE IF NOT EXISTS template_variables (
            template_id INTEGER,
            variable_name TEXT,
            PRIMARY KEY (template_id, variable_name),
            FOREIGN KEY (template_id) REFERENCES templates(id) ON DELETE CASCADE
        )''')
        
        conn.commit()

    # 设置相关方法
    def save_setting(self, key: str, value: str):
        """保存设置
        
        Args:
            key (str): 设置键
            value (str): 设置值
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute(
            "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
            (key, value)
        )
        
        conn.commit()
        
    def add_app_info_translation(self, language_code: str, key: str, value: str):
        """添加或更新 app_info 多語言資訊"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT OR REPLACE INTO app_info_translations (key, language_code, value) VALUES (?, ?, ?)",
            (key, language_code, value)
        )
        conn.commit()

    def get_app_info_translations(self, language_code: str) -> Dict[str, str]:
        """獲取指定語言所有 app_info 翻譯"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT key, value FROM app_info_translations WHERE language_code = ?",
            (language_code,)
        )
        return {row['key']: row['value'] for row in cursor.fetchall()}
    
    def _load_translations(self) -> Dict[str, Dict[str, str]]:
        translations = {}
        languages = self.db_manager.get_languages()
        for lang in languages:
            lang_code = lang['code']
            # 通用翻譯
            translations[lang_code] = self.db_manager.get_translations(lang_code)
            # 加入 app_info
            app_info_trans = self.db_manager.get_app_info_translations(lang_code)
            translations[lang_code].update(app_info_trans)
        return translations

    def get_setting(self, key: str, default: str = None) -> Optional[str]:
        """获取设置
        
        Args:
            key (str): 设置键
            default (str, optional): 默认值
        
        Returns:
            Optional[str]: 设置值，如果不存在则返回默认值
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT value FROM settings WHERE key = ?", (key,))
        result = cursor.fetchone()
        
        return result[0] if result else default
    
    # 语言相关方法
    def add_language(self, code: str, description: str):
        """添加语言
        
        Args:
            code (str): 语言代码
            description (str): 语言描述
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute(
            "INSERT OR REPLACE INTO languages (code, description) VALUES (?, ?)",
            (code, description)
        )
        
        conn.commit()

    
    def get_languages(self) -> List[Dict]:
        """获取所有语言
        
        Returns:
            List[Dict]: 语言列表，每项包含code和description
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT code, description FROM languages")
        
        return [dict(row) for row in cursor.fetchall()]
    
    def add_translation(self, language_code: str, key: str, text: str):
        """添加翻译
        
        Args:
            language_code (str): 语言代码
            key (str): 翻译键
            text (str): 翻译文本
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute(
            "INSERT OR REPLACE INTO translations (language_code, key, text) VALUES (?, ?, ?)",
            (language_code, key, text)
        )
        
        conn.commit()
    
    def get_translations(self, language_code: str) -> Dict[str, str]:
        """获取指定语言的所有翻译
        
        Args:
            language_code (str): 语言代码
        
        Returns:
            Dict[str, str]: 翻译字典，键为翻译键，值为翻译文本
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute(
            "SELECT key, text FROM translations WHERE language_code = ?",
            (language_code,)
        )
        
        return {row['key']: row['text'] for row in cursor.fetchall()}
    
    # 事件类型相关方法
    def add_event_type(self, name: str) -> int:
        """添加事件类型
        
        Args:
            name (str): 事件类型名称
        
        Returns:
            int: 新事件类型的ID
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute(
            "INSERT OR IGNORE INTO event_types (name) VALUES (?)",
            (name,)
        )
        
        conn.commit()
        
        # 获取新插入的ID或已存在的ID
        cursor.execute("SELECT id FROM event_types WHERE name = ?", (name,))
        result = cursor.fetchone()
        
        return result[0] if result else None
    
    def get_event_type_id(self, name: str) -> Optional[int]:
        """根据名称获取事件类型ID
        
        Args:
            name (str): 事件类型名称
        
        Returns:
            Optional[int]: 事件类型ID，如果不存在则返回None
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT id FROM event_types WHERE name = ?", (name,))
        result = cursor.fetchone()
        
        return result[0] if result else None
    
    def get_event_types(self) -> List[Dict]:
        """获取所有事件类型
        
        Returns:
            List[Dict]: 事件类型列表，每项包含id和name
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT id, name FROM event_types")
        
        return [dict(row) for row in cursor.fetchall()]
    
    def delete_event_type(self, event_type_id: int) -> bool:
        """删除事件类型
        
        Args:
            event_type_id (int): 事件类型ID
        
        Returns:
            bool: 是否成功删除
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            cursor.execute("DELETE FROM event_types WHERE id = ?", (event_type_id,))
            conn.commit()
            return cursor.rowcount > 0
        except sqlite3.Error:
            conn.rollback()
            return False
    
    # 模板相关方法
    def add_template(self, event_type_id: int, name: str, recipient: str, cc: str, 
                     subject: str, body: str, variables: List[str], note_en: str, tag_en: str, sender: str) -> int:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            conn.execute("BEGIN")
            cursor.execute("SELECT id FROM templates WHERE event_type_id = ? AND name = ?", (event_type_id, name))
            existing = cursor.fetchone()
            if existing:
                cursor.execute("""UPDATE templates 
                               SET recipient = ?, cc = ?, subject = ?, body = ?, note_en = ?, tag_en = ?, sender = ?
                               WHERE id = ?""",
                               (recipient, cc, subject, body, note_en, tag_en, sender, existing[0]))
                template_id = existing[0]
                cursor.execute("DELETE FROM template_variables WHERE template_id = ?", (template_id,))
            else:
                cursor.execute("""INSERT INTO templates 
                                (event_type_id, name, recipient, cc, subject, body, note_en, tag_en, sender) 
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                               (event_type_id, name, recipient, cc, subject, body, note_en, tag_en, sender))
                template_id = cursor.lastrowid
            for variable in variables:
                cursor.execute("INSERT INTO template_variables (template_id, variable_name) VALUES (?, ?)",
                               (template_id, variable))
            conn.commit()
            return template_id
        except sqlite3.Error as e:
            conn.rollback()
            print(f"DB error: {e}")
            return None
    
    def get_template(self, template_id: int) -> Optional[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("""SELECT t.id, t.event_type_id, t.name, t.recipient, t.cc, t.subject, t.body, t.note_en, t.tag_en, t.sender,
                               et.name as event_type_name
                          FROM templates t
                          JOIN event_types et ON t.event_type_id = et.id
                          WHERE t.id = ?""", (template_id,))
        template = cursor.fetchone()
        if not template:
            return None
        template_dict = dict(template)
        cursor.execute("SELECT variable_name FROM template_variables WHERE template_id = ?", (template_id,))
        template_dict['variables'] = [row[0] for row in cursor.fetchall()]
        return template_dict
    
    def get_template_by_name(self, event_type: str, template_name: str) -> Optional[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("""SELECT t.id, t.event_type_id, t.name, t.recipient as recipient, t.cc, t.subject, t.body, t.note_en, t.tag_en, t.sender
                          FROM templates t
                          JOIN event_types et ON t.event_type_id = et.id
                          WHERE et.name = ? AND t.name = ?""", (event_type, template_name))
        template = cursor.fetchone()
        if not template:
            return None
        template_dict = dict(template)
        if 'recipient' in template_dict:
            template_dict['to'] = template_dict['recipient']
            del template_dict['recipient']
        cursor.execute("SELECT variable_name FROM template_variables WHERE template_id = ?", (template_dict['id'],))
        template_dict['variables'] = [row[0] for row in cursor.fetchall()]
        return template_dict
        
    def get_templates_for_event(self, event_type: str) -> List[Dict]:
        """獲取指定事件類型的所有模板
        
        Args:
            event_type (str): 事件類型名稱
        
        Returns:
            List[Dict]: 模板列表
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute(
            """SELECT t.id, t.name, t.recipient, t.cc, t.subject, t.body, t.note_en, t.tag_en, t.sender
            FROM templates t
            JOIN event_types et ON t.event_type_id = et.id
            WHERE et.name = ?""",
            (event_type,)
        )
        
        templates = []
        for template in cursor.fetchall():
            template_dict = dict(template)
            
            # 修改鍵名以匹配預期輸出
            if 'recipient' in template_dict:
                template_dict['to'] = template_dict['recipient']
                del template_dict['recipient']
            
            # 獲取變量
            cursor.execute(
                "SELECT variable_name FROM template_variables WHERE template_id = ?",
                (template_dict['id'],)
            )
            
            template_dict['variables'] = [row[0] for row in cursor.fetchall()]
            templates.append(template_dict)
        
        return templates
    
    def get_template_names_for_event(self, event_type: str) -> List[str]:
        """获取指定事件类型的所有模板名称
        
        Args:
            event_type (str): 事件类型名称
        
        Returns:
            List[str]: 模板名称列表
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute(
            """SELECT t.name
              FROM templates t
              JOIN event_types et ON t.event_type_id = et.id
              WHERE et.name = ?""",
            (event_type,)
        )
        
        return [row[0] for row in cursor.fetchall()]
    
    def delete_template(self, template_id: int) -> bool:
        """删除模板
        
        Args:
            template_id (int): 模板ID
        
        Returns:
            bool: 是否成功删除
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            cursor.execute("DELETE FROM templates WHERE id = ?", (template_id,))
            conn.commit()
            return cursor.rowcount > 0
        except sqlite3.Error:
            conn.rollback()
            return False
    
    def delete_template_by_name(self, event_type: str, template_name: str) -> bool:
        """根据事件类型和模板名称删除模板
        
        Args:
            event_type (str): 事件类型名称
            template_name (str): 模板名称
        
        Returns:
            bool: 是否成功删除
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            cursor.execute(
                """DELETE FROM templates 
                   WHERE event_type_id = (SELECT id FROM event_types WHERE name = ?)
                   AND name = ?""",
                (event_type, template_name)
            )
            conn.commit()
            return cursor.rowcount > 0
        except sqlite3.Error:
            conn.rollback()
            return False
    
    def search_templates(self, keyword: str) -> List[Dict]:
        """搜索包含關鍵字的模板
        
        Args:
            keyword (str): 搜索關鍵字
        
        Returns:
            List[Dict]: 模板列表
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # 使用LIKE進行模糊匹配
        keyword = f"%{keyword}%"
        
        cursor.execute(
            """SELECT t.id, t.name, t.recipient, t.cc, t.subject, t.body, t.note_en, t.tag_en, t.sender, et.name as event_type
            FROM templates t
            JOIN event_types et ON t.event_type_id = et.id
            WHERE t.name LIKE ? OR t.subject LIKE ? OR t.body LIKE ?""",
            (keyword, keyword, keyword)
        )
        
        results = []
        for template in cursor.fetchall():
            template_dict = dict(template)
            
            # 修改鍵名以匹配預期輸出
            if 'recipient' in template_dict:
                template_dict['to'] = template_dict['recipient']
                del template_dict['recipient']
            
            # 獲取變量
            cursor.execute(
                "SELECT variable_name FROM template_variables WHERE template_id = ?",
                (template_dict['id'],)
            )
            
            template_dict['variables'] = [row[0] for row in cursor.fetchall()]
            # 構造返回格式與JSON版一致
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
                    "sender": template_dict['sender'],  # 添加 sender 欄位
                }
            })
        
        return results
    
    # 导入导出方法
    def export_templates(self) -> Dict:
        """導出所有模板
        
        Returns:
            Dict: 包含所有事件類型和模板的字典
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # 獲取所有事件類型
        cursor.execute("SELECT id, name FROM event_types")
        event_types = cursor.fetchall()
        
        result = {"event_types": []}
        
        for event_type in event_types:
            event_type_id = event_type['id']
            event_type_name = event_type['name']
            
            # 獲取該事件類型下的所有模板
            cursor.execute(
                """SELECT id, name, recipient as to, cc, subject, body, note_en, tag_en, sender
                FROM templates
                WHERE event_type_id = ?""",
                (event_type_id,)
            )
            
            templates = []
            for template in cursor.fetchall():
                template_dict = dict(template)
                
                # 獲取變量
                cursor.execute(
                    "SELECT variable_name FROM template_variables WHERE template_id = ?",
                    (template_dict['id'],)
                )
                
                variables = [row[0] for row in cursor.fetchall()]
                
                # 移除id欄位
                del template_dict['id']
                
                templates.append({
                    **template_dict,
                    "variables": variables
                })
            
            result["event_types"].append({
                "name": event_type_name,
                "templates": templates
            })
        
        return result
    
    def import_templates(self, data: Dict) -> bool:
        """導入模板
        
        Args:
            data (Dict): 包含事件類型和模板的字典
        
        Returns:
            bool: 是否成功導入
        """
        if "event_types" not in data:
            return False
        
        conn = self.get_connection()
        cursor = conn.cursor()
        
        try:
            # 開始事務
            conn.execute("BEGIN")
            
            # 清空現有數據
            cursor.execute("DELETE FROM template_variables")
            cursor.execute("DELETE FROM templates")
            cursor.execute("DELETE FROM event_types")
            
            # 導入事件類型和模板
            for event_type in data["event_types"]:
                event_type_name = event_type["name"]
                
                # 添加事件類型
                cursor.execute(
                    "INSERT INTO event_types (name) VALUES (?)",
                    (event_type_name,)
                )
                event_type_id = cursor.lastrowid
                
                # 添加模板
                for template in event_type["templates"]:
                    cursor.execute(
                        """INSERT INTO templates 
                        (event_type_id, name, recipient, cc, subject, body, note_en, tag_en, sender) 
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (
                            event_type_id, 
                            template["name"], 
                            template.get("to", ""), 
                            template.get("cc", ""), 
                            template.get("subject", ""), 
                            template.get("body", ""),
                            template.get("note_en", ""),
                            template.get("tag_en", ""),
                            template.get("sender", "")
                        )
                    )
                    template_id = cursor.lastrowid
                    
                    # 添加變量
                    for variable in template.get("variables", []):
                        cursor.execute(
                            "INSERT INTO template_variables (template_id, variable_name) VALUES (?, ?)",
                            (template_id, variable)
                        )
            
            conn.commit()
            return True
            
        except (sqlite3.Error, KeyError) as e:
            conn.rollback()
            print(f"導入失敗: {e}")
            return False
        
    def get_app_info(self, key, language="en_US"):
        """獲取特定鍵和語言的應用程序信息
        
        Args:
            key (str): 信息鍵
            language (str, optional): 語言代碼. 默認為 "en_US".
            
        Returns:
            str: 應用程序信息值，若找不到則返回鍵本身
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # 先嘗試從多語言表獲取
        cursor.execute(
            "SELECT value FROM app_info_translations WHERE key = ? AND language_code = ?", 
            (key, language)
        )
        result = cursor.fetchone()
        
        if result:
            return result[0]
        
        # 如果找不到指定語言的翻譯，嘗試獲取英文翻譯
        if language != "en_US":
            cursor.execute(
                "SELECT value FROM app_info_translations WHERE key = ? AND language_code = 'en_US'", 
                (key, )
            )
            result = cursor.fetchone()
            
            if result:
                return result[0]
        
        # 如果在多語言表中找不到，嘗試從基本表獲取
        cursor.execute("SELECT value FROM app_info WHERE key = ?", (key,))
        result = cursor.fetchone()
        
        return result[0] if result else key
        
    def migrate_from_json(self, template_file: str = 'data/templates.json', 
                        settings_file: str = 'data/settings.json',
                        language_dir: str = 'data/language') -> bool:
        """从JSON文件迁移数据
        
        Args:
            template_file (str): 模板文件路径
            settings_file (str): 设置文件路径
            language_dir (str): 语言文件目录
        
        Returns:
            bool: 是否成功迁移
        """
        try:
            # 迁移模板
            if os.path.exists(template_file):
                with open(template_file, 'r', encoding='utf-8') as f:
                    templates_data = json.load(f)
                self.import_templates(templates_data)
            
            # 迁移设置
            if os.path.exists(settings_file):
                with open(settings_file, 'r', encoding='utf-8') as f:
                    settings_data = json.load(f)
                for key, value in settings_data.items():
                    self.save_setting(key, str(value))
            
            # 迁移语言
            if os.path.exists(language_dir):
                # 添加默认语言
                self.add_language('zh_TW', '繁體中文')
                self.add_language('en_US', 'English')
                
                # 迁移翻译
                for lang_code in ['zh_TW', 'en_US']:
                    lang_file = os.path.join(language_dir, f'{lang_code}.json')
                    if os.path.exists(lang_file):
                        with open(lang_file, 'r', encoding='utf-8') as f:
                            translations = json.load(f)
                        for key, text in translations.items():
                            self.add_translation(lang_code, key, text)
            
            return True
        except Exception as e:
            print(f"迁移失败: {e}")
            return False
