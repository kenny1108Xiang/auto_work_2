import json
import logging
import re
import requests

URLS_FILE = "forms_url.txt"

# 設定 logging
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s.%(msecs)03d - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

class FormClosedException(Exception):
    """自訂異常：用於表示表單已關閉、已滿或不接受回應"""
    pass


def resolve_short_url(day_number, mode=1):
    """
    根據星期數字和模式，從 URL 檔案中讀取短網址並解析為完整的表單 URL。
    
    參數:
        day_number: 星期數字 (1=星期一, 7=星期日)
        mode: 0 使用 forms_url_test.txt, 1 使用 forms_url.txt (預設為 1)
    
    返回:
        str: 完整的表單 URL (viewform)，失敗則返回 None
    """
    # 根據 mode 選擇檔案
    urls_file_path = "form/forms_url_test.txt" if mode == 0 else "form/forms_url.txt"
    
    try:
        day_index = int(day_number)
    except (TypeError, ValueError):
        logging.error(f"提供的數值 {day_number!r} 無法轉換成整數。")
        return None
    
    if day_index < 1 or day_index > 7:
        logging.error(f"星期數字必須在 1-7 之間，收到：{day_index}")
        return None
    
    # 讀取 URL 檔案
    try:
        with open(urls_file_path, "r", encoding="utf-8") as url_file:
            urls = [line.strip() for line in url_file if line.strip()]
    except OSError as e:
        logging.error(f"無法讀取 URL 清單檔案 {urls_file_path}。 {e}")
        return None
    
    # 檢查是否有足夠的 URL
    if day_index > len(urls):
        logging.error(f"在 {urls_file_path} 中找不到第 {day_index} 天的短網址。")
        return None
    
    short_url = urls[day_index - 1]
    logging.info(f"正在從第 {day_index} 天的短網址 {short_url} 解析正式表單網址...")
    
    # 解析短網址
    try:
        resolve_response = requests.get(short_url, allow_redirects=False)
        resolve_response.raise_for_status()
        form_url = resolve_response.headers.get("Location")
        if not form_url:
            logging.error("短網址回應缺少 Location 標頭。")
            return None
        logging.info(f"成功解析表單 URL: {form_url}")
        return form_url
    except requests.exceptions.RequestException as e:
        logging.error(f"解析短網址時失敗。 {e}")
        return None


def fetch_form_entry_ids_for_day(form_url, day_number):
    """
    從 Google 表單中抓取欄位的 Entry ID。
    
    星期一到五 (1-5)：只抓取「姓名」和「選項」
    星期六到日 (6-7)：抓取「姓名」、「選項」和「原因」
    
    參數:
        form_url: 完整的表單 URL
        day_number: 星期數字 (1-7)，用於判斷是否需要抓取原因欄位
    
    返回:
        tuple: (name_entry, option_entry, reason_entry)
               星期 1-5 時 reason_entry 為 None
    """
    name_entry = None
    option_entry = None
    reason_entry = None
    
    # 判斷是否需要抓取原因欄位（只有星期六、日需要）
    need_reason = int(day_number) >= 6
    
    if not form_url:
        logging.error("表單 URL 為空，無法抓取欄位 ID。")
        return name_entry, option_entry, reason_entry
    
    logging.info(f"正在從 {form_url} 抓取表單欄位資訊...")
    
    try:
        response = requests.get(form_url)
        response.raise_for_status()
        
        # Google 表單會將其結構資訊存在名為 FB_PUBLIC_LOAD_DATA_ 的 JS 變數中
        form_data_match = re.search(r'var FB_PUBLIC_LOAD_DATA_ = (.*?);', response.text)
        if not form_data_match:
            logging.error("在頁面原始碼中找不到表單結構資料 (FB_PUBLIC_LOAD_DATA_)")
            return name_entry, option_entry, reason_entry
        
        # 解析這個 JS 變數的內容 (它是一個 JSON 格式的陣列)
        form_data = json.loads(form_data_match.group(1))
        
        entry_map = {}
        # 問題列表通常儲存在這個巢狀結構中
        questions = form_data[1][1]
        for question in questions:
            question_text = question[1]  # 取得問題的文字描述
            entry_id = question[4][0][0]  # 取得問題對應的 entry ID
            entry_map[question_text] = f"entry.{entry_id}"
        
        # 根據問題的關鍵字，找到我們需要的 entry ID
        name_entry = next((v for k, v in entry_map.items() if "姓名" in k), None)
        option_entry = next((v for k, v in entry_map.items() if "排休" in k), None)
        
        # 只在星期六、日才抓取原因欄位
        if need_reason:
            reason_entry = next((v for k, v in entry_map.items() if "原因" in k), None)
            logging.info(f"抓取到欄位 ID: 姓名={name_entry}, 選項={option_entry}, 原因={reason_entry}")
        else:
            logging.info(f"抓取到欄位 ID: 姓名={name_entry}, 選項={option_entry} (星期 {day_number} 不需要原因)")
        
    except requests.exceptions.RequestException as e:
        logging.error(f"無法訪問表單頁面。 {e}")
    except (AttributeError, IndexError, json.JSONDecodeError) as e:
        logging.error(f"解析表單結構時失敗，可能是表單格式有變。 {e}")
    
    return name_entry, option_entry, reason_entry