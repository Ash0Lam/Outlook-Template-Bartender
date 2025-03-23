from typing import Dict, List, Any, Optional
from db_manager import DatabaseManager

class TemplateManager:
    """管理電子郵件模板的類"""
    
    def __init__(self, db_manager: DatabaseManager = None):
        """初始化模板管理器
        
        Args:
            db_manager (DatabaseManager, optional): 数据库管理器实例，如果为None则创建新实例
        """
        self.db_manager = db_manager if db_manager else DatabaseManager()
        self._init_default_templates()
    
    def _init_default_templates(self):
        """初始化默认模板"""
        # 检查是否已有模板，如果没有则添加默认模板
        event_types = self.get_event_types()
        if not event_types:
            # 添加默认事件类型
            event_type_id = self.db_manager.add_event_type("Storage Event")
            
            # 添加默认模板
            self.db_manager.add_template(
                event_type_id,
                "Storage Outage",
                "recipient1@example.com",
                "cc1@example.com, cc2@example.com",
                "Storage Outage Notification - {ID}",
                "<html><body><p>Dear Team,</p><p>We are experiencing a storage outage at {Location}. The issue has been identified with ID {ID}.</p><p>The affected company is {Company}.</p><p>We are working to resolve this issue and will provide updates as necessary.</p><p>Best regards,</p></body></html>",
                ["ID", "Location", "Company"],
                "This is a storage outage notification.",  # English note
                "这是一个存储中断通知。"  # Chinese note
            )
    
    def get_template_by_name(self, event_type: str, template_name: str) -> Optional[Dict]:
        """根据事件类型和模板名称获取模板
        
        Args:
            event_type (str): 事件类型名称
            template_name (str): 模板名称
        
        Returns:
            Optional[Dict]: 模板信息
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute(
            """SELECT t.id, t.event_type_id, t.name, t.recipient as recipient, t.cc, t.subject, t.body, t.note_en
            FROM templates t
            JOIN event_types et ON t.event_type_id = et.id
            WHERE et.name = ? AND t.name = ?""",
            (event_type, template_name)
        )
        
        template = cursor.fetchone()
        if not template:
            return None
            
        template_dict = dict(template)
        
        # 修改键名以匹配预期输出
        if 'recipient' in template_dict:
            template_dict['to'] = template_dict['recipient']
            del template_dict['recipient']
        
        # 获取变量
        cursor.execute(
            "SELECT variable_name FROM template_variables WHERE template_id = ?",
            (template_dict['id'],)
        )
        
        template_dict['variables'] = [row[0] for row in cursor.fetchall()]
        
        # 添加 note_en 和 note_zh 到返回结果
        template_dict['note_en'] = template_dict.get('note_en', '')  # 默认值为空字符串
        
        return template_dict


    def get_event_types(self) -> List[str]:
        """獲取所有事件類型名稱
        
        Returns:
            List[str]: 事件類型名稱列表
        """
        event_types = self.db_manager.get_event_types()
        return [event_type['name'] for event_type in event_types]
    
    def get_templates_for_event(self, event_type: str) -> List[Dict]:
        """獲取特定事件類型的所有模板
        
        Args:
            event_type (str): 事件類型名稱
            
        Returns:
            List[Dict]: 模板列表
        """
        return self.db_manager.get_templates_for_event(event_type)
    
    def get_template_names_for_event(self, event_type: str) -> List[str]:
        """獲取特定事件類型的所有模板名稱
        
        Args:
            event_type (str): 事件類型名稱
            
        Returns:
            List[str]: 模板名稱列表
        """
        return self.db_manager.get_template_names_for_event(event_type)
    
    def get_template(self, event_type: str, template_name: str) -> Optional[Dict]:
        """獲取特定事件類型和模板名稱的模板"""
        template_data = self.db_manager.get_template_by_name(event_type, template_name)
        if template_data:
            # 将数据库中的字段映射到返回的数据结构
            template = {
                "name": template_data['name'],
                "to": template_data['to'],
                "cc": template_data['cc'],
                "subject": template_data['subject'],
                "body": template_data['body'],
                "variables": template_data['variables'],
                "note_en": template_data['note_en'],
                "tag_en": template_data['tag_en'],
                "sender": template_data.get('sender', '')  # 添加 sender 欄位
            }
            return template
        return None
        
    def add_event_type(self, event_type: str) -> None:
        """添加新的事件類型
        
        Args:
            event_type (str): 事件類型名稱
        """
        self.db_manager.add_event_type(event_type)
    
    def remove_event_type(self, event_type: str) -> None:
        """刪除事件類型
        
        Args:
            event_type (str): 事件類型名稱
        """
        event_type_id = self.db_manager.get_event_type_id(event_type)
        if event_type_id:
            self.db_manager.delete_event_type(event_type_id)
    
    def add_template(self, event_type: str, template: Dict) -> None:
        """添加新模板到指定事件類型"""
        event_type_id = self.db_manager.get_event_type_id(event_type)
        if not event_type_id:
            event_type_id = self.db_manager.add_event_type(event_type)
        
        self.db_manager.add_template(
            event_type_id,
            template.get("name", ""),
            template.get("to", ""),
            template.get("cc", ""),
            template.get("subject", ""),
            template.get("body", ""),
            template.get("variables", []),
            template.get("note_en", ""),
            template.get("tag_en", ""),
            template.get("sender", "")   # ⭐️ 這行要補上！
        )

    
    def remove_template(self, event_type: str, template_name: str) -> None:
        """從指定事件類型中刪除模板
        
        Args:
            event_type (str): 事件類型名稱
            template_name (str): 模板名稱
        """
        self.db_manager.delete_template_by_name(event_type, template_name)
    
    def search_templates(self, keyword: str) -> List[Dict]:
        """搜索包含特定關鍵字的模板
        
        Args:
            keyword (str): 搜索關鍵字
            
        Returns:
            List[Dict]: 匹配的模板列表，每個字典包含事件類型和模板
        """
        return self.db_manager.search_templates(keyword)
    
    def export_templates(self, filename: str) -> None:
        """將模板導出到JSON文件
        
        Args:
            filename (str): 導出文件路徑
        """
        import json
        templates_data = self.db_manager.export_templates()
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(templates_data, f, ensure_ascii=False, indent=2)
    
    def import_templates(self, filename: str) -> bool:
        """從JSON文件導入模板
        
        Args:
            filename (str): 導入文件路徑
            
        Returns:
            bool: 是否成功導入
        """
        import json
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                templates_data = json.load(f)
            return self.db_manager.import_templates(templates_data)
        except (json.JSONDecodeError, FileNotFoundError) as e:
            print(f"导入失败: {e}")
            return False
    
    def migrate_from_json(self, template_file: str = 'data/templates.json') -> bool:
        """从JSON文件迁移模板数据
        
        Args:
            template_file (str): 模板JSON文件路径
            
        Returns:
            bool: 是否迁移成功
        """
        return self.db_manager.migrate_from_json(template_file=template_file)
    
    def update_template(self, event_type: str, template_name: str, updated_template: Dict) -> None:
        """更新模板"""
        event_type_id = self.db_manager.get_event_type_id(event_type)
        if not event_type_id:
            raise ValueError(f"Event type {event_type} does not exist")
        
        # 更新模板数据
        self.db_manager.update_template(
            event_type_id,
            template_name,
            updated_template.get("name", ""),
            updated_template.get("to", ""),
            updated_template.get("cc", ""),
            updated_template.get("subject", ""),
            updated_template.get("body", ""),
            updated_template.get("variables", []),
            updated_template.get("note_en", ""),
            updated_template.get("tag_en", "")
        )
    def get_event_type_id(self, event_type_name: str) -> Optional[int]:
        """
        根據事件類型名稱獲取ID
        
        Args:
            event_type_name (str): 事件類型名稱
            
        Returns:
            Optional[int]: 事件類型ID，若不存在則返回None
        """
        return self.db_manager.get_event_type_id(event_type_name)

    def copy_template_to_event_type(self, source_event_type: str, target_event_type: str, template_name: str) -> bool:
        """
        將模板從一個事件類型複製到另一個事件類型
        
        Args:
            source_event_type (str): 源事件類型名稱
            target_event_type (str): 目標事件類型名稱
            template_name (str): 模板名稱
            
        Returns:
            bool: 操作是否成功
        """
        # 獲取模板數據
        template = self.get_template(source_event_type, template_name)
        if not template:
            return False
            
        # 檢查目標事件類型是否存在
        target_event_type_id = self.get_event_type_id(target_event_type)
        if not target_event_type_id:
            target_event_type_id = self.db_manager.add_event_type(target_event_type)
            
        # 添加模板到目標事件類型
        return self.add_template(target_event_type, template) is not None

    def move_template_to_event_type(self, source_event_type: str, target_event_type: str, template_name: str) -> bool:
        """
        將模板從一個事件類型移動到另一個事件類型
        
        Args:
            source_event_type (str): 源事件類型名稱
            target_event_type (str): 目標事件類型名稱
            template_name (str): 模板名稱
            
        Returns:
            bool: 操作是否成功
        """
        # 首先複製模板
        if not self.copy_template_to_event_type(source_event_type, target_event_type, template_name):
            return False
            
        # 然後刪除原模板
        return self.remove_template(source_event_type, template_name)

    def rename_event_type(self, old_name: str, new_name: str) -> bool:
        """
        重命名事件類型
        
        Args:
            old_name (str): 舊事件類型名稱
            new_name (str): 新事件類型名稱
            
        Returns:
            bool: 操作是否成功
        """
        # 檢查新名稱是否已存在
        if self.get_event_type_id(new_name):
            return False
            
        # 獲取所有模板
        templates = self.get_templates_for_event(old_name)
        
        # 創建新事件類型
        self.add_event_type(new_name)
        
        # 將所有模板複製到新事件類型
        for template in templates:
            self.add_template(new_name, template)
        
        # 刪除舊事件類型（將自動刪除相關模板）
        self.remove_event_type(old_name)
        
        return True
