# update_translations.py
import sqlite3

def update_translations():
    """更新數據庫中的翻譯，添加簽名檔相關文本"""
    conn = sqlite3.connect('data/app.db')
    cursor = conn.cursor()
    
    # 簽名檔相關翻譯
    new_translations = {
        "signature": {
            "en_US": "Signature",
            "zh_TW": "簽名檔"
        },
        "signature_options": {
            "en_US": "Signature Options",
            "zh_TW": "簽名檔選項"
        },
        "use_default_signature": {
            "en_US": "Use default signature",
            "zh_TW": "使用默認簽名檔"
        },
        "no_signature": {
            "en_US": "No signature",
            "zh_TW": "不使用簽名檔"
        },
        "custom_signature": {
            "en_US": "Custom signature",
            "zh_TW": "自定義簽名檔"
        },
        "select_signature": {
            "en_US": "Select signature",
            "zh_TW": "選擇簽名檔"
        },
        "default_signature": {
            "en_US": "Default",
            "zh_TW": "默認"
        },
        "none_signature": {
            "en_US": "None",
            "zh_TW": "無"
        }
    }
    
    # 獲取所有支持的語言
    cursor.execute("SELECT code FROM languages")
    languages = [row[0] for row in cursor.fetchall()]
    
    if not languages:
        print("No languages found in database")
        return
    
    # 為每種語言添加翻譯
    translations_added = 0
    for key, translations in new_translations.items():
        for lang_code in languages:
            # 檢查翻譯是否已存在
            cursor.execute("SELECT 1 FROM translations WHERE language_code = ? AND key = ?", 
                           (lang_code, key))
            if cursor.fetchone():
                continue  # 已存在，跳過
            
            # 獲取該語言的翻譯，如果沒有使用英文
            text = translations.get(lang_code, translations.get("en_US", key))
            
            # 添加翻譯
            cursor.execute("INSERT INTO translations (language_code, key, text) VALUES (?, ?, ?)",
                           (lang_code, key, text))
            translations_added += 1
    
    conn.commit()
    print(f"Added {translations_added} new translations")
    conn.close()

if __name__ == "__main__":
    update_translations()