import subprocess
import psutil
import time

def is_outlook_running():
    """檢查 Outlook 是否正在運行"""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and 'OUTLOOK.EXE' in proc.info['name'].upper():
            return True
    return False

def start_outlook():
    """啟動 Outlook"""
    try:
        subprocess.Popen(['start', 'outlook'], shell=True)
        time.sleep(5)  # 等待 Outlook 啟動
        print("Outlook 已啟動")
    except Exception as e:
        print(f"啟動 Outlook 時出錯: {e}")

if __name__ == "__main__":
    if not is_outlook_running():
        print("Outlook 未運行，正在啟動...")
        start_outlook()
    else:
        print("Outlook 已經在運行。")