import os
import re
import uuid
import base64
import shutil
from pathlib import Path

class ImageManager:
    """管理電子郵件模板中的圖片"""
    
    def __init__(self, app_dir=None):
        """初始化圖片管理器
        
        Args:
            app_dir (str, optional): 應用程序目錄路徑，如果為None，則使用當前目錄
        """
        self.app_dir = app_dir or os.path.dirname(os.path.abspath(__file__))
        self.image_dir = os.path.join(self.app_dir, 'images')
        
        # 確保圖片目錄存在
        os.makedirs(self.image_dir, exist_ok=True)
    
    def process_html_content(self, html_content, template_name):
        """處理HTML內容，提取並保存圖片
        
        Args:
            html_content (str): HTML內容
            template_name (str): 模板名稱，用於建立唯一的圖片目錄
            
        Returns:
            str: 處理後的HTML內容，圖片引用已替換為本地路徑
        """
        # 為模板建立唯一目錄
        template_dir = self._get_template_image_dir(template_name)
        
        # 先檢查是否存在base64圖片
        if "data:image" not in html_content:
            return html_content
        
        # 提取所有base64圖片
        img_pattern = r'<img [^>]*?src="(data:image/([^;]+);base64,([^"]+))"[^>]*?>'
        matches = re.finditer(img_pattern, html_content)
        
        # 替換圖片路徑
        for match in matches:
            base64_uri = match.group(1)
            img_format = match.group(2)
            base64_data = match.group(3)
            
            # 生成唯一文件名
            img_filename = f"{uuid.uuid4()}.{img_format}"
            img_path = os.path.join(template_dir, img_filename)
            
            # 保存圖片
            try:
                image_data = base64.b64decode(base64_data)
                with open(img_path, 'wb') as f:
                    f.write(image_data)
                
                # 替換圖片引用為Content-ID引用
                html_content = html_content.replace(base64_uri, f"cid:{img_filename}")
            except Exception as e:
                print(f"保存圖片時出錯: {e}")
        
        return html_content
    
    def _get_template_image_dir(self, template_name):
        """獲取模板的圖片目錄，確保存在
        
        Args:
            template_name (str): 模板名稱
            
        Returns:
            str: 模板圖片目錄路徑
        """
        # 生成目錄安全的名稱
        safe_name = re.sub(r'[^\w\-_]', '_', template_name)
        template_dir = os.path.join(self.image_dir, safe_name)
        
        # 確保目錄存在
        os.makedirs(template_dir, exist_ok=True)
        
        return template_dir
    
    def get_image_paths(self, template_name):
        """獲取模板的所有圖片路徑
        
        Args:
            template_name (str): 模板名稱
            
        Returns:
            list: 圖片路徑列表
        """
        template_dir = self._get_template_image_dir(template_name)
        if not os.path.exists(template_dir):
            return []
        
        return [os.path.join(template_dir, f) for f in os.listdir(template_dir) 
                if os.path.isfile(os.path.join(template_dir, f)) 
                and f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
    
    def cleanup_unused_images(self, html_content, template_name):
        """清理未使用的圖片
        
        Args:
            html_content (str): HTML內容
            template_name (str): 模板名稱
        """
        template_dir = self._get_template_image_dir(template_name)
        if not os.path.exists(template_dir):
            return
        
        # 提取HTML中引用的所有圖片
        referenced_images = set()
        img_pattern = r'<img [^>]*?src="cid:([^"]+)"[^>]*?>'
        for match in re.finditer(img_pattern, html_content):
            referenced_images.add(match.group(1))
        
        # 刪除未引用的圖片
        for filename in os.listdir(template_dir):
            if filename not in referenced_images and os.path.isfile(os.path.join(template_dir, filename)):
                try:
                    os.remove(os.path.join(template_dir, filename))
                except Exception as e:
                    print(f"刪除未使用圖片時出錯: {e}")
                    
    def rename_template_image_dir(self, old_name, new_name):
        """當模板重命名時更新圖片目錄
        
        Args:
            old_name (str): 舊模板名稱
            new_name (str): 新模板名稱
            
        Returns:
            bool: 是否成功重命名
        """
        old_dir = self._get_template_image_dir(old_name)
        new_dir = self._get_template_image_dir(new_name)
        
        if not os.path.exists(old_dir):
            return True  # 舊目錄不存在，視為成功
            
        try:
            # 如果新目錄已存在，先清空
            if os.path.exists(new_dir) and old_dir != new_dir:
                shutil.rmtree(new_dir)
            # 重命名目錄
            if old_dir != new_dir:
                shutil.move(old_dir, new_dir)
            return True
        except Exception as e:
            print(f"重命名圖片目錄時出錯: {e}")
            return False