# vinh_long_all_in_one.py
# CHẠY 1 LẦN: Cập nhật VINH_LONG + SO_NGANH + PHUONG_XA
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from google.oauth2 import service_account
from googleapiclient.discovery import build
import sys
import os
from datetime import datetime
import time
import string

# === FIX TIẾNG VIỆT ===
if sys.platform.startswith('win'):
    import codecs
    sys.stdout = codecs.getwriter("utf-8")(sys.stdout.detach())

# === CẤU HÌNH CHUNG ===
URL = "https://dichvucong.gov.vn/p/home/dvc-index-tinhthanhpho-tonghop.html"
SPREADSHEET_ID = "1aaHMEIPuifMauJEmyeUtdwr3HhJs9FcqBbBpzHAE-nc"
SERVICE_ACCOUNT_FILE = "service_account.json"
TINH_MUC_TIEU = "Vĩnh Long"

# Sheet names
SHEET_VL = "VINH_LONG"
SHEET_SO = "SO_NGANH"
SHEET_XA = "PHUONG_XA"

# VINH_LONG: Cột cần tính Δ
DELTA_COLUMNS = [
    "Công khai, minh bạch",
    "Tiến độ giải quyết",
    "Dịch vụ trực tuyến",
    "Mức độ hài lòng",
    "Số hóa hồ sơ",
    "Tổng điểm"
]

# MÀU SẮC
GREEN_FILL = {"red": 0.78, "green": 0.94, "blue": 0.81}
RED_FILL = {"red": 1.0, "green": 0.78, "blue": 0.78}
YELLOW_FILL = {"red": 1.0, "green": 0.92, "blue": 0.6}
GREEN_FONT = {"red": 0.0, "green": 0.39, "blue": 0.0}
RED_FONT = {"red": 0.61, "green": 0.0, "blue": 0.02}
YELLOW_FONT = {"red": 0.61, "green": 0.4, "blue": 0.0}
BOLD_FONT = {"bold": True}

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
driver = None

# === HÀM HỖ TRỢ ===
def connect_google_sheets():
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return build('sheets', 'v4', credentials=credentials)

def get_sheet_id(service, sheet_name):
    spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for sheet in spreadsheet.get('sheets', []):
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    print(f"Sheet '{sheet_name}' chưa tồn tại → Tạo mới...")
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]}
    ).execute()
    return get_sheet_id(service, sheet_name)  # Gọi lại

def get_column_letter(n):
    return string.ascii_uppercase[n - 1] if n <= 26 else 'Z'

def safe_float(val):
    if not val: return 0.0
    try: return float(str(val).replace(',', '.').strip())
    except: return 0.0

# === 1. CẬP NHẬT VINH_LONG (Δ + DỮ LIỆU MỚI) ===
def update_vinh_long_sheet(service, headers_web, vinh_long_row, thoi_gian_moi):
    print(f"\nCập nhật sheet: {SHEET_VL}...")
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_VL}!A:ZZ"
    ).execute()
    values = result.get('values', [])
    sheet_id = get_sheet_id(service, SHEET_VL)

    # === TÌM DÒNG DỮ LIỆU CUỐI CÙNG (BỎ QUA Δ) ===
    last_data_row = 0  # Dòng số (1-based)
    old_values = {col: 0.0 for col in DELTA_COLUMNS}
    if values:
        headers_sheet = values[0]
        for r_idx in range(len(values) - 1, 0, -1):
            row = values[r_idx]
            if len(row) > 0 and row[0] not in ["Δ", ""]:
                last_data_row = r_idx + 1
                for col_name in DELTA_COLUMNS:
                    try:
                        sheet_col_idx = headers_sheet.index(col_name)
                        old_val = safe_float(row[sheet_col_idx] if sheet_col_idx < len(row) else "")
                        old_values[col_name] = old_val
                    except:
                        old_values[col_name] = 0.0
                break

    # === XÁC ĐỊNH VỊ TRÍ GHI MỚI ===
    delta_row = last_data_row + 1  # Δ nằm ngay dưới dòng dữ liệu cũ
    new_data_row = delta_row + 1  # Dữ liệu mới nằm dưới Δ

    # === TÍNH Δ ===
    delta_dict = {}
    for col_name in DELTA_COLUMNS:
        try:
            web_idx = headers_web.index(col_name)
            new_val = safe_float(vinh_long_row[web_idx])
        except:
            new_val = 0.0
        old_val = old_values[col_name]
        delta = new_val - old_val
        delta_dict[col_name] = delta

    requests = []

    # === TẠO HEADER NẾU CHƯA CÓ ===
    if not values:
        full_header = ["THỜI GIAN LẤY DỮ LIỆU"] + headers_web
        header_cells = [{"userEnteredValue": {"stringValue": h}, "userEnteredFormat": {"textFormat": BOLD_FONT}} for h in full_header]
        requests.append({
            "updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
                "rows": [{"values": header_cells}],
                "fields": "userEnteredValue,userEnteredFormat.textFormat"
            }
        })
        delta_row = 2
        new_data_row = 3

    # === GHI DÒNG Δ (CHỈ KHI CÓ DỮ LIỆU CŨ) ===
    if last_data_row > 0:
        delta_cells = [{"userEnteredValue": {"stringValue": "Δ"}}]
        for h in headers_web:
            if h in DELTA_COLUMNS:
                delta = delta_dict[h]
                cell = {"userEnteredValue": {"numberValue": delta}}
                fill = GREEN_FILL if delta > 0 else RED_FILL if delta < 0 else None
                if fill:
                    cell["userEnteredFormat"] = {"backgroundColor": fill}
                delta_cells.append(cell)
            else:
                delta_cells.append({"userEnteredValue": {"stringValue": ""}})
        requests.append({
            "updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": delta_row - 1, "endRowIndex": delta_row},
                "rows": [{"values": delta_cells}],
                "fields": "userEnteredValue,userEnteredFormat.backgroundColor"
            }
        })

    # === GHI DỮ LIỆU MỚI ===
    new_cells = [{"userEnteredValue": {"stringValue": thoi_gian_moi}}]
    for val in vinh_long_row:
        try:
            num = float(val.replace(',', '.'))
            new_cells.append({"userEnteredValue": {"numberValue": num}})
        except:
            new_cells.append({"userEnteredValue": {"stringValue": val}})

    requests.append({
        "updateCells": {
            "range": {"sheetId": sheet_id, "startRowIndex": new_data_row - 1, "endRowIndex": new_data_row},
            "rows": [{"values": new_cells}],
            "fields": "userEnteredValue"
        }
    })

    if requests:
        service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": requests}).execute()

    print(f"→ {SHEET_VL}: Ghi Δ (dòng {delta_row}) + Dữ liệu mới (dòng {new_data_row}) → KHÔNG GHI ĐÈ")

# === 2. CẬP NHẬT SO_NGANH / PHUONG_XA (NGANG) ===
def update_horizontal_sheet(service, sheet_name, data_rows, full_time):
    print(f"\nCập nhật sheet: {sheet_name}...")
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A:ZZ"
    ).execute()
    values = result.get('values', [])
    sheet_id = get_sheet_id(service, sheet_name)

    last_diem_col = None
    if values and len(values) > 0:
        header = values[0]
        diem_cols = [i for i, h in enumerate(header) if h == "Điểm"]
        if diem_cols:
            last_diem_col = diem_cols[-1]

    start_col = len(values[0]) if values else 2
    diff_col = start_col
    new_diem_col = start_col + 1

    # Header
    if not values:
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            body={"values": [["STT", "Tên đơn vị"]]}
        ).execute()

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!{get_column_letter(diff_col + 1)}1",
        valueInputOption="USER_ENTERED",
        body={"values": [["Chênh lệch"]]}
    ).execute()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!{get_column_letter(new_diem_col + 1)}1",
        valueInputOption="USER_ENTERED",
        body={"values": [["Điểm"]]}
    ).execute()

    # Dữ liệu + Δ
    diff_values = []
    new_diem_values = []
    requests = []
    for idx, row_data in enumerate(data_rows):
        ten_don_vi = row_data[1]
        diem_moi = safe_float(row_data[2])
        old_val = 0.0
        if last_diem_col is not None and len(values) > idx + 1:
            old_cell = values[idx + 1][last_diem_col] if last_diem_col < len(values[idx + 1]) else ""
            old_val = safe_float(old_cell)
        diff = round(diem_moi - old_val, 2)

        row_idx = idx + 2
        diff_values.append([diff])
        new_diem_values.append([diem_moi])

        fill = GREEN_FILL if diff > 0 else RED_FILL if diff < 0 else YELLOW_FILL
        font_color = GREEN_FONT if diff > 0 else RED_FONT if diff < 0 else YELLOW_FONT
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row_idx - 1,
                    "endRowIndex": row_idx,
                    "startColumnIndex": diff_col,
                    "endColumnIndex": diff_col + 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": fill,
                        "textFormat": {"foregroundColor": font_color, "bold": True}
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)"
            }
        })

    if diff_values:
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!{get_column_letter(diff_col + 1)}2",
            valueInputOption="USER_ENTERED",
            body={"values": diff_values}
        ).execute()
    if new_diem_values:
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!{get_column_letter(new_diem_col + 1)}2",
            valueInputOption="USER_ENTERED",
            body={"values": new_diem_values}
        ).execute()

    if requests:
        service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": requests}).execute()

    # GHI THỜI GIAN DƯỚI CỘT ĐIỂM MỚI
    last_data_row = len(data_rows) + 1
    time_row = last_data_row + 1
    time_cell = f"{sheet_name}!{get_column_letter(new_diem_col + 1)}{time_row}"
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=time_cell,
        valueInputOption="USER_ENTERED",
        body={"values": [[full_time]]}
    ).execute()

    print(f"→ {sheet_name}: +Chênh lệch + Điểm | Thời gian: {time_cell}")

# === MAIN ===
try:
    print("Khởi động Chrome...")
    options = Options()
    # options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    driver.get(URL)
    wait = WebDriverWait(driver, 30)

    # === LẤY DỮ LIỆU VINH_LONG ===
    print("Lấy dữ liệu tổng hợp Vĩnh Long...")
    container = wait.until(EC.presence_of_element_located((By.ID, "table-container")))
    table = wait.until(EC.presence_of_element_located((By.XPATH, ".//div[@id='table-container']//table")))
    headers_web = [th.text.strip() for th in table.find_elements(By.XPATH, ".//tr[th]/th")]
    vinh_long_row = None
    for row in table.find_elements(By.XPATH, ".//tr[td]"):
        cells = row.find_elements(By.TAG_NAME, "td")
        row_data = [cell.text.strip() for cell in cells]
        if len(row_data) > 1 and TINH_MUC_TIEU in row_data[1]:
            vinh_long_row = row_data
            break
    if not vinh_long_row:
        raise Exception("Không tìm thấy Vĩnh Long")

    # === CHỌN VĨNH LONG ĐỂ LẤY 2 BẢNG ===
    print("Chọn Vĩnh Long để lấy bảng chi tiết...")
    dropdown = wait.until(EC.element_to_be_clickable((By.ID, "select2-tinhtp-container")))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown)
    time.sleep(0.5)
    ActionChains(driver).move_to_element(dropdown).click().perform()
    search_box = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "select2-search__field")))
    search_box.clear()
    search_box.send_keys(TINH_MUC_TIEU)
    time.sleep(1)
    option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[contains(text(), '{TINH_MUC_TIEU}')]")))
    ActionChains(driver).move_to_element(option).click().perform()
    time.sleep(5)

    # === LẤY 2 BẢNG ===
    container_detail = driver.find_element(By.CLASS_NAME, "province-table-container")
    tables = container_detail.find_elements(By.TAG_NAME, "table")
    tables_data = []
    for i, table in enumerate(tables, 1):
        data_rows = []
        for row in table.find_elements(By.XPATH, ".//tr[td]"):
            cells = row.find_elements(By.TAG_NAME, "td")
            row_data = []
            for idx, cell in enumerate(cells):
                text = cell.text.strip()
                if idx == 2:
                    span = cell.find_elements(By.TAG_NAME, "span")
                    if span: text = span[0].text.strip()
                row_data.append(text)
            if row_data:
                data_rows.append(row_data)
        tables_data.append({
            "data": data_rows,
            "sheet": SHEET_SO if i == 1 else SHEET_XA
        })

    full_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    print(f"Thời gian lấy: {full_time}")

    driver.quit()
    driver = None

    # === GHI GOOGLE SHEETS ===
    service = connect_google_sheets()

    # 1. VINH_LONG
    update_vinh_long_sheet(service, headers_web, vinh_long_row, full_time)

    # 2. SO_NGANH & PHUONG_XA
    for table in tables_data:
        update_horizontal_sheet(service, table['sheet'], table['data'], full_time)

    print(f"\nHOÀN TẤT! ĐÃ CẬP NHẬT 3 SHEET:")
    print(f"→ {SHEET_VL}")
    print(f"→ {SHEET_SO}")
    print(f"→ {SHEET_XA}")
    print(f"→ Link: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")

except Exception as e:
    print(f"LỖI: {e}")
    import traceback
    traceback.print_exc()
finally:
    if driver:
        try: driver.quit()
        except: pass