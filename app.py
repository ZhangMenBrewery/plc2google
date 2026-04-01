from flask import Flask, render_template, jsonify, request
from flask_socketio import SocketIO, emit
import psycopg2
from psycopg2 import OperationalError
import pygsheets
import time
import datetime
import snap7
import json
import string
import os
import requests
from threading import Lock

app = Flask(__name__)
app.config['SECRET_KEY'] = 'zmb-brewery-plc-monitor'
socketio = SocketIO(app, cors_allowed_origins="*", manage_session=False)

# 全域變數
plc_client = None
plc_tag = {}
plc_data_cache = {}
data_lock = Lock()
running = True

# 設定檔案路徑
SETTINGS_FILE = 'settings.json'

def load_settings():
    """載入設定檔案"""
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"載入設定失敗：{e}")
    return {}

def save_settings(settings):
    """儲存設定檔案"""
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        print(f"儲存設定失敗：{e}")
        return False

def apply_settings_to_runtime(settings):
    """將設定即時套用到執行中的 plc_reader（可免重啟）"""
    global plc_reader
    if not plc_reader:
        return

    general_config = settings.get('general', {})
    plc_reader.write_interval = general_config.get('write_interval', 300)

    db_config = settings.get('database', {})
    plc_reader.sql_host = db_config.get('db_host', '192.168.60.12')
    plc_reader.sql_user = db_config.get('db_user', 'zhangmen')
    plc_reader.sql_password = db_config.get('db_password', '54685508')
    plc_reader.sql_database = db_config.get('db_database', 'zmb')
    plc_reader.sql_enabled = db_config.get('enabled', False)

    gs_config = settings.get('google_sheets', {})
    plc_reader.gs_key = gs_config.get('gs_key', 'zmb54685508-c88132768091.json')
    plc_reader.gs_enabled = gs_config.get('enabled', False)

    beername_config = settings.get('beername', {})
    plc_reader.beername_enabled = beername_config.get('enabled', False)
    plc_reader.beername_sheet_url = beername_config.get('sheet_url', 'https://docs.google.com/spreadsheets/d/1QUh-ZHlJSFG0RhkmP0JT7WBxRUqC8EgCMwGn7lZVQas/edit#gid=877936540')
    plc_reader.beername_worksheet_title = beername_config.get('worksheet_title', '發酵桶現況')
    plc_reader.beername_column_index = beername_config.get('column_index', 3)
    plc_reader.beername_start_row = beername_config.get('start_row', 2)
    plc_reader.beername_end_row = beername_config.get('end_row', 23)
    plc_reader.beername_db_number = beername_config.get('db_number', 142)
    plc_reader.beername_string_size = beername_config.get('string_size', 256)
    plc_reader.beername_update_interval = beername_config.get('update_interval', 300)

    ranges = beername_config.get('ranges', [])
    if not ranges:
        ranges = [{
            'worksheet_title': plc_reader.beername_worksheet_title,
            'column_index': plc_reader.beername_column_index,
            'start_row': plc_reader.beername_start_row,
            'end_row': plc_reader.beername_end_row
        }]
    plc_reader.beername_ranges = ranges

    plc_config = settings.get('plc', {})
    new_ip = plc_config.get('plc_ip', "192.168.60.201")
    new_rack = plc_config.get('plc_rack', 0)
    new_slot = plc_config.get('plc_slot', 2)

    should_reconnect = (
        plc_reader.plc_ip != new_ip or
        plc_reader.plc_rack != new_rack or
        plc_reader.plc_slot != new_slot
    )

    plc_reader.plc_ip = new_ip
    plc_reader.plc_rack = new_rack
    plc_reader.plc_slot = new_slot

    if should_reconnect:
        try:
            plc_reader.client.disconnect()
        except Exception:
            pass
        try:
            plc_reader.client.connect(plc_reader.plc_ip, plc_reader.plc_rack, plc_reader.plc_slot)
            print(f"已重新連接到 PLC: {plc_reader.plc_ip}")
        except Exception as e:
            print(f"重新連接 PLC 失敗：{e}")

class ZMBPlcReader:
    """PLC 數據讀取類別"""
    
    def __init__(self):
        # 從設定檔案載入配置
        settings = load_settings()
        
        # 一般設定
        general_config = settings.get('general', {})
        self.write_interval = general_config.get('write_interval', 300)
        
        # Google Sheets 設定
        gs_config = settings.get('google_sheets', {})
        self.gs_key = gs_config.get('gs_key', 'zmb54685508-c88132768091.json')
        self.gs_title = 'ZMB-' + str(datetime.date.today())[:-3]
        self.gs_enabled = gs_config.get('enabled', False)
        
        # 酒款名稱寫入設定（更新模式 - 合併範圍 1+ 範圍 2）
        beername_config = settings.get('beername', {})
        self.beername_enabled = beername_config.get('enabled', False)
        self.beername_sheet_url = beername_config.get('sheet_url', 'https://docs.google.com/spreadsheets/d/1QUh-ZHlJSFG0RhkmP0JT7WBxRUqC8EgCMwGn7lZVQas/edit#gid=877936540')
        self.beername_worksheet_title = beername_config.get('worksheet_title', '發酵桶現況')
        self.beername_column_index = beername_config.get('column_index', 3)  # 預設第 3 列 (C 欄)
        self.beername_start_row = beername_config.get('start_row', 2)  # 預設從第 2 行開始
        self.beername_end_row = beername_config.get('end_row', 23)  # 預設到第 23 行
        self.beername_db_number = beername_config.get('db_number', 142)  # PLC DB 編號
        self.beername_string_size = beername_config.get('string_size', 256)  # 每個酒款字串大小
        self.beername_update_interval = beername_config.get('update_interval', 300)  # 更新間隔（秒）
        
        # 範圍設定（合併範圍 1+ 範圍 2）
        self.beername_ranges = beername_config.get('ranges', [])
        if not self.beername_ranges:
            self.beername_ranges = [{
                'worksheet_title': self.beername_worksheet_title,
                'column_index': self.beername_column_index,
                'start_row': self.beername_start_row,
                'end_row': self.beername_end_row
            }]
        
        # 資料庫設定
        db_config = settings.get('database', {})
        self.sql_host = db_config.get('db_host', '192.168.60.12')
        self.sql_user = db_config.get('db_user', 'zhangmen')
        self.sql_password = db_config.get('db_password', '54685508')
        self.sql_database = db_config.get('db_database', 'zmb')
        self.sql_enabled = db_config.get('enabled', False)
        
        # PLC 設定
        plc_config = settings.get('plc', {})
        self.plc_ip = plc_config.get('plc_ip', "192.168.60.201")
        self.plc_rack = plc_config.get('plc_rack', 0)
        self.plc_slot = plc_config.get('plc_slot', 2)
        
        # 載入 PLC 標籤配置
        self.load_plc_tags()
        
        # 載入 SQL 配置
        self.load_sql_config()
        
        # 初始化 PLC 客戶端
        self.client = snap7.client.Client()
        try:
            self.client.connect(self.plc_ip, self.plc_rack, self.plc_slot)
            print(f"成功連接到 PLC: {self.plc_ip}")
        except Exception as e:
            print(f"連接 PLC 失敗：{e}")
    
    def load_sql_config(self):
        """載入 SQL 配置"""
        try:
            # 載入 SQL 建立命令
            path_create = 'sql_create.json'
            if os.path.exists(path_create):
                with open(path_create, encoding='utf-8', errors='ignore') as f:
                    self.sql_create_commands = json.load(f, strict=False)
            else:
                self.sql_create_commands = {}
            
            # 載入 SQL 插入命令
            path = 'sql_insert.json'
            if os.path.exists(path):
                with open(path, encoding='utf-8', errors='ignore') as f:
                    self.sql_insert_tag = json.load(f, strict=False)
            else:
                self.sql_insert_tag = {}
        except Exception as e:
            print(f"載入 SQL 配置失敗：{e}")
            self.sql_create_commands = {}
            self.sql_insert_tag = {}
    
    def write_to_sql(self, data):
        """將數據寫入資料庫"""
        if not self.sql_enabled:
            return
        
        try:
            conn = psycopg2.connect(
                host=self.sql_host,
                user=self.sql_user,
                password=self.sql_password,
                dbname=self.sql_database
            )
            cur = conn.cursor()
            success_count = 0
            
            for region, region_data in data.items():
                try:
                    # 建立資料表（如果不存在），DDL 單獨 commit
                    if region in self.sql_create_commands:
                        cur.execute(self.sql_create_commands[region])
                        conn.commit()
                    
                    # 動態建立 INSERT 語句（欄位名稱移除空格並轉小寫）
                    timestamp = region_data.get('timestamp', datetime.datetime.now())
                    tags = self.plc_tag.get(region, {})
                    table_name = region_to_table(region)
                    if not table_name:
                        print(f"找不到區域 {region} 對應的資料表，略過寫入")
                        continue
                    
                    columns = ['timestamp']
                    values = [timestamp]
                    
                    for tag_name in tags.keys():
                        if tag_name in region_data and tag_name != 'timestamp':
                            columns.append(tag_name.lower().replace(' ', ''))
                            values.append(region_data[tag_name])
                    
                    placeholders = ','.join(['%s'] * len(values))
                    col_names = ','.join(columns)
                    sql = f"INSERT INTO {table_name} ({col_names}) VALUES ({placeholders}) ON CONFLICT DO NOTHING"
                    
                    cur.execute(sql, values)
                    conn.commit()
                    success_count += 1
                except Exception as e:
                    print(f"寫入 {region} 到資料庫失敗：{e}")
                    conn.rollback()
            
            cur.close()
            conn.close()
            print(f"資料庫寫入成功：{success_count}/{len(data)} 個區域")
        except Exception as e:
            print(f"寫入資料庫失敗：{e}")
    
    def write_to_google_sheet(self, data):
        """將數據寫入 Google Sheets"""
        if not self.gs_enabled:
            return
        
        # 動態生成 Google Sheet 標題，確保換日時使用正確的日期
        gs_title = 'ZMB-' + str(datetime.date.today())[:-3]
        print(f"[DEBUG] write_to_google_sheet - 使用標題：{gs_title}")
        
        try:
            gc = pygsheets.authorize(service_account_file=self.gs_key)
            
            try:
                ss = gc.open(gs_title)
                print(f"[DEBUG] 找到現有 Google Sheet: {gs_title}")
                # 確保現有表格也有被分享給指定人員 (處理之前的遺漏)
                zmb_group = ['zhangmenbrewery@gmail.com', 'chunkai721@gmail.com']
                for member in zmb_group:
                    try:
                        ss.share(member, role='writer', type='user')
                    except Exception as e:
                        print(f"[DEBUG] 補分享給 {member} 失敗：{e}")
            except pygsheets.SpreadsheetNotFound:
                print(f"[DEBUG] 創建新的 Google Sheet: {gs_title}")
                # gc.create 返回 pygsheets.Spreadsheet 對象
                ss = gc.create(gs_title)
                # 分享給指定人員
                zmb_group = ['zhangmenbrewery@gmail.com', 'chunkai721@gmail.com']
                for member in zmb_group:
                    try:
                        ss.share(member, role='writer', type='user')
                        print(f"[DEBUG] 已分享給 {member}")
                    except Exception as e:
                        print(f"[DEBUG] 分享給 {member} 失敗：{e}")
                
                try:
                    ss.share('', role='reader', type='anyone')
                except Exception as e:
                    print(f"[DEBUG] 設置公開檢視失敗：{e}")
                
                # 刪除預設的 sheet1
                try:
                    ss.del_worksheet(ss.sheet1)
                except Exception as e:
                    print(f"[DEBUG] 刪除預設工作表失敗：{e}")
            
            success_count = 0
            for region, region_data in data.items():
                try:
                    # worksheet_by_title 找不到時會拋出例外，需用 try/except 處理
                    try:
                        wks = ss.worksheet_by_title(region)
                    except pygsheets.WorksheetNotFound:
                        tags = self.plc_tag.get(region, {})
                        num_cols = len(tags) + 1
                        # 處理欄位英文字母 (A=1, Z=26)
                        if num_cols <= 26:
                            letter = string.ascii_uppercase[num_cols - 1]
                        else:
                            letter = 'Z'
                            
                        header = ['Time'] + list(tags.keys())
                        wks = ss.add_worksheet(title=region, rows=10000, cols=len(header))
                        wks.update_values(f'A1:{letter}1', [header])
                        wks.frozen_rows = 1
                    
                    # 準備數據 - 使用標籤名稱對應 region_data 的 key
                    rows = len(wks.get_col(1, include_tailing_empty=False)) + 1
                    data_row = [datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
                    
                    for tag_name in self.plc_tag.get(region, {}).keys():
                        data_row.append(region_data.get(tag_name, ''))
                    
                    num_cols = len(self.plc_tag.get(region, {})) + 1
                    if num_cols <= 26:
                        last_col = string.ascii_uppercase[num_cols - 1]
                    else:
                        last_col = 'Z'
                        
                    wks.update_values(f'A{rows}:{last_col}{rows}', [data_row])
                    success_count += 1
                except Exception as e:
                    print(f"寫入 {region} 到 Google Sheet 失敗：{e}")
            
            print(f"Google Sheets 寫入成功：{success_count}/{len(data)} 個區域")
        except Exception as e:
            print(f"寫入 Google Sheet 失敗：{e}")
    
    def load_plc_tags(self):
        """載入 PLC 標籤配置"""
        try:
            with open('plc_tag.json', encoding='utf-8', errors='ignore') as f:
                self.plc_tag = json.load(f, strict=False)
        except Exception as e:
            print(f"載入 PLC 標籤失敗：{e}")
            self.plc_tag = {}
        
        # 載入區域順序配置
        try:
            with open('region_order.json', encoding='utf-8', errors='ignore') as f:
                order_data = json.load(f, strict=False)
                self.region_order = order_data.get('order', [])
        except Exception as e:
            print(f"載入區域順序失敗：{e}")
            self.region_order = []
    
    def ensure_plc_connected(self):
        """確保 PLC 連線正常，如果斷線則重新連接"""
        try:
            # 檢查連線狀態
            if not self.client or not self.client.get_connected():
                print("PLC 未連接，嘗試連接...")
                try:
                    self.client.connect(self.plc_ip, self.plc_rack, self.plc_slot)
                    print(f"成功連接到 PLC: {self.plc_ip}")
                    return True
                except Exception as e:
                    print(f"連接 PLC 失敗：{e}")
                    return False
            return True
        except Exception as e:
            print(f"檢查 PLC 連線狀態失敗：{e}")
            return False
    
    def read_plc_value(self, tag):
        """從 PLC 讀取單一標籤的值"""
        # 確保連線正常
        if not self.ensure_plc_connected():
            return None
            
        try:
            tag_parts = tag.split('.')
            
            if tag_parts[0].startswith('DB'):
                # DB 區塊讀取
                db_num = int(tag_parts[0][2:])
                
                if tag_parts[1].startswith('DBDDI'):
                    # DBDDI - 整數
                    offset = int(tag_parts[1][5:])
                    data = self.client.db_read(db_num, offset, 4)
                    return snap7.util.get_dint(data, 0)
                elif tag_parts[1].startswith('DBD'):
                    # DBD - 浮點數
                    offset = int(tag_parts[1][3:])
                    data = self.client.db_read(db_num, offset, 4)
                    return snap7.util.get_real(data, 0)
                elif tag_parts[1].startswith('DBX'):
                    # DBX - 布林值
                    offset = int(tag_parts[1][3:])
                    bit = int(tag_parts[2])
                    data = self.client.db_read(db_num, offset, 1)
                    return 1 if snap7.util.get_bool(data, 0, bit) else 0
                else:
                    return None
                    
            elif tag_parts[0].startswith('Q'):
                # 輸出區域
                byte_num = int(tag_parts[0][1:])
                bit = int(tag_parts[1])
                data = self.client.ab_read(0, 11)
                return 1 if snap7.util.get_bool(data, byte_num, bit) else 0
            else:
                return None
                
        except Exception as e:
            error_msg = repr(e)  # 使用 repr() 來正確顯示二進制錯誤訊息
            # 如果是 Socket error 或 Connection closed，表示連接已斷開，需要重新連接
            if "Socket error" in error_msg or "Connection closed" in error_msg or "Connection refused" in error_msg or "broken pipe" in error_msg.lower():
                print(f"PLC 連接已斷開，嘗試重新連接...")
                try:
                    self.client.disconnect()
                    time.sleep(1)  # 等待 1 秒後重連
                    self.client.connect(self.plc_ip, self.plc_rack, self.plc_slot)
                    print(f"成功重新連接到 PLC: {self.plc_ip}")
                except Exception as reconnect_error:
                    print(f"重新連接 PLC 失敗：{reconnect_error}")
                    return None
            else:
                print(f"讀取標籤 {tag} 失敗：{e}")
            return None
    
    def get_all_plc_data(self):
        """讀取所有 PLC 數據"""
        global plc_data_cache
        
        from collections import OrderedDict
        
        with data_lock:
            # 使用 OrderedDict 保持順序
            data = OrderedDict()
            
            # 按照 region_order 中的順序排列
            regions_to_process = self.region_order if self.region_order else sorted(self.plc_tag.keys())
            
            for region in regions_to_process:
                if region not in self.plc_tag:
                    continue
                tags = self.plc_tag[region]
                region_data = OrderedDict()
                region_data["timestamp"] = datetime.datetime.now().isoformat()
                
                # 按照標籤在 plc_tag.json 中的順序
                for tag_name, tag_address in tags.items():
                    value = self.read_plc_value(tag_address)
                    if value is not None:
                        region_data[tag_name] = value
                
                data[region] = region_data
            
            # 處理未在 region_order 中但存在的區域
            for region in self.plc_tag:
                if region not in data:
                    region_data = OrderedDict()
                    region_data["timestamp"] = datetime.datetime.now().isoformat()
                    for tag_name, tag_address in self.plc_tag[region].items():
                        value = self.read_plc_value(tag_address)
                        if value is not None:
                            region_data[tag_name] = value
                    data[region] = region_data
            
            plc_data_cache = data
            
            # 讀取完成後斷線，不持續占用連線
            try:
                if self.client:
                    self.client.disconnect()
                    print("PLC 讀取完成，已斷開連線")
            except Exception as e:
                print(f"斷開 PLC 連線失敗：{e}")
            
            return data
    
    def get_plc_status(self):
        """檢查 PLC 連接狀態"""
        try:
            # 嘗試讀取一個簡單的值來檢查連接
            self.client.get_cpu_info()
            return {"connected": True, "ip": self.plc_ip}
        except Exception as e:
            return {"connected": False, "error": str(e)}
    
    def close(self):
        """關閉 PLC 連接"""
        if self.client:
            try:
                self.client.disconnect()
            except:
                pass
    
    def write_beername_to_plc(self):
        """從 Google Sheet 讀取酒款名稱並寫入 PLC DB142（使用 Big5 編碼）- 合併範圍 1+ 範圍 2"""
        try:
            return self._write_beername_merged()
        except Exception as e:
            print(f"寫入酒款名稱到 PLC 失敗：{e}")
            return False
    
    def _write_beername_merged(self):
        """合併模式寫入 - 將範圍 1 和範圍 2 的對應行號酒款名稱合併後寫入 PLC（B2+E2、B3+E3、...）"""
        if not self.beername_ranges:
            print("範圍設定為空，略過酒款名稱寫入")
            return False

        try:
            gc = pygsheets.authorize(service_account_file=self.gs_key)
            ss = gc.open_by_url(self.beername_sheet_url)
            
            # 獲取第一個範圍的設定
            range1 = self.beername_ranges[0]
            range2 = self.beername_ranges[1] if len(self.beername_ranges) > 1 else None
            
            # 獲取工作表
            if 'worksheet_title' in range1:
                wks1 = ss.worksheet_by_title(range1['worksheet_title'])
            else:
                wks1 = ss.worksheet_by_title(self.beername_worksheet_title)
            
            if range2:
                if 'worksheet_title' in range2:
                    wks2 = ss.worksheet_by_title(range2['worksheet_title'])
                else:
                    wks2 = ss.worksheet_by_title(self.beername_worksheet_title)
            else:
                wks2 = None
            
            # 讀取範圍 1 和範圍 2 的酒款名稱
            start_row = range1.get('start_row', self.beername_start_row)
            end_row = range1.get('end_row', self.beername_end_row)
            column_index1 = range1.get('column_index', self.beername_column_index)
            
            col_data1 = wks1.get_col(column_index1)
            col_data2 = wks2.get_col(range2.get('column_index', self.beername_column_index)) if wks2 else None
            
            # 讀取第 2-23 行（共 22 個）
            beername1 = col_data1[start_row-1:23]
            beername2 = col_data2[start_row-1:23] if col_data2 else None
            
            # 將對應行號的酒款名稱合併成一個字串
            all_beernames = []
            for i in range(len(beername1)):
                if beername2 and i < len(beername2):
                    # 合併範圍 1 和範圍 2 的酒款名稱
                    combined = beername1[i] + beername2[i]
                    all_beernames.append(combined)
                else:
                    all_beernames.append(beername1[i])
            
            print(f"範圍 1: 讀取 {len(beername1)} 個酒款名稱 (column_index={column_index1}, 實際讀取：{beername1})")
            if beername2:
                print(f"範圍 2: 讀取 {len(beername2)} 個酒款名稱 (column_index={range2.get('column_index')}, 實際讀取：{beername2})")
            print(f"合併後總共 {len(all_beernames)} 個酒款名稱")
            print(f"酒款名稱列表：{all_beernames}")
            
            # 寫入合併後的資料到 PLC
            result = self._write_beername_to_plc(all_beernames)
            
            return result
        except Exception as e:
            print(f"合併寫入失敗：{e}")
            return False
    
    def _write_beername_to_plc(self, beername):
        """將酒款名稱寫入 PLC（共用方法）"""
        try:
            # 連接 PLC
            client = snap7.client.Client()
            client.connect(self.plc_ip, self.plc_rack, self.plc_slot)
            
            # 計算最大可寫入的酒款數量（DB 大小限制）
            max_beer_count = 22  # DB142 通常大小為 5632 位元組，5632/256 = 22
            actual_count = min(len(beername), max_beer_count)
            
            print(f"準備寫入 {len(beername)} 個酒款名稱，實際寫入 {actual_count} 個（DB 大小限制）")
            
            # 將每個酒款名稱寫入 PLC 的 DB 區塊（使用 Big5 編碼）
            for i in range(actual_count):
                data = bytearray(self.beername_string_size)
                
                # 使用 Big5 編碼轉換繁體中文
                try:
                    beername_bytes = beername[i].encode('big5', errors='ignore')
                except:
                    beername_bytes = beername[i].encode('utf-8', errors='ignore')
                
                # S7-300 String 格式：
                # 第一個字節：最大長度
                # 第二個字節：實際長度
                # 其餘：字串內容
                max_len = self.beername_string_size - 2
                actual_len = min(len(beername_bytes), max_len)
                
                data[0] = max_len  # 最大長度
                data[1] = actual_len  # 實際長度
                
                # 複製字節數據到 DB
                for j in range(actual_len):
                    data[2 + j] = beername_bytes[j]
                
                address = i * self.beername_string_size
                print(f"寫入酒款 {i+1}/{actual_count}: {beername[i]} (地址：{address})")
                client.db_write(db_number=self.beername_db_number, start=address, data=data)
            
            client.disconnect()
            print(f"成功寫入 {actual_count} 個酒款名稱到 PLC DB{self.beername_db_number}（Big5 編碼）")
            return True
        except Exception as e:
            print(f"寫入酒款名稱到 PLC 失敗：{e}")
            return False

# 全域 PLC 讀取器
plc_reader = None

def init_plc_reader():
    """初始化 PLC 讀取器"""
    global plc_reader
    plc_reader = ZMBPlcReader()
    return plc_reader

def plc_data_updater():
    """背景任務：定期更新 PLC 數據"""
    global running, plc_data_cache
    
    last_write_time = 0  # 上次寫入時間戳記
    last_beername_write = 0  # 上次酒款名稱寫入時間戳記
    
    while running:
        try:
            if plc_reader:
                data = plc_reader.get_all_plc_data()
                # 每 5 秒更新即時顯示
                socketio.emit('plc_update', data)
                
                # 依 write_interval 頻率寫入資料庫與 Google Sheets
                current_time = time.time()
                write_interval = plc_reader.write_interval if plc_reader else 300
                if current_time - last_write_time >= write_interval:
                    if plc_reader.sql_enabled:
                        plc_reader.write_to_sql(data)
                    if plc_reader.gs_enabled:
                        plc_reader.write_to_google_sheet(data)
                    last_write_time = current_time

                # 酒款更新獨立觸發，不再綁定 write_interval（5 分鐘）
                if plc_reader and plc_reader.beername_enabled:
                    update_interval = plc_reader.beername_update_interval if plc_reader.beername_update_interval > 0 else 300
                    if current_time - last_beername_write >= update_interval:
                        plc_reader.write_beername_to_plc()
                        last_beername_write = current_time
        except Exception as e:
            print(f"更新 PLC 數據時發生錯誤：{e}")
        
        time.sleep(5)  # 每 5 秒讀取 PLC 更新即時顯示

@app.route('/')
def index():
    """首頁"""
    # 按照 region_order 或字母順序排列區域
    if plc_reader:
        if plc_reader.region_order:
            regions = plc_reader.region_order
        else:
            # 按照字母順序排列
            regions = sorted(plc_reader.plc_tag.keys())
    else:
        regions = []
    return render_template('index.html', regions=regions)

@app.route('/api/plc/status')
def api_plc_status():
    """API: 獲取 PLC 連接狀態"""
    if plc_reader:
        return jsonify(plc_reader.get_plc_status())
    return jsonify({"connected": False, "error": "PLC 讀取器未初始化"})

@app.route('/api/plc/data')
def api_plc_data():
    """API: 獲取所有 PLC 數據"""
    global plc_data_cache
    with data_lock:
        # 確保返回的是列表格式，保持順序
        return jsonify(list(plc_data_cache.items()))

@app.route('/api/plc/region/<region_name>')
def api_plc_region(region_name):
    """API: 獲取特定區域的 PLC 數據"""
    global plc_data_cache
    with data_lock:
        if region_name in plc_data_cache:
            return jsonify(plc_data_cache[region_name])
        return jsonify({"error": f"區域 {region_name} 不存在"})

@app.route('/api/plc-tags')
def api_plc_tags():
    """API: 獲取 PLC 標籤配置"""
    global plc_tag
    if plc_reader:
        return jsonify(plc_reader.plc_tag)
    return jsonify({})

@app.route('/api/plc-tags', methods=['POST'])
def api_save_plc_tags():
    """API: 儲存 PLC 標籤配置"""
    try:
        new_tags = request.get_json()
        if isinstance(new_tags, dict):
            # 更新 plc_tag.json
            with open('plc_tag.json', 'w', encoding='utf-8') as f:
                json.dump(new_tags, f, ensure_ascii=False, indent=4)
            # 更新 plc_reader 的標籤
            if plc_reader:
                plc_reader.plc_tag = new_tags
                # 重新載入區域順序
                plc_reader.load_plc_tags()
            return jsonify({'status': 'success', 'message': 'PLC 標籤已更新'})
        return jsonify({'status': 'error', 'message': '無效的請求格式'}), 400
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/regions')
def api_regions():
    """API: 獲取所有可用區域列表"""
    global plc_tag
    if plc_reader:
        # 返回按照順序的區域列表
        if plc_reader.region_order:
            return jsonify(plc_reader.region_order)
        return jsonify(list(plc_reader.plc_tag.keys()))
    return jsonify([])

@app.route('/api/region-order', methods=['GET', 'POST'])
def api_region_order():
    """API: 獲取或更新區域順序"""
    if request.method == 'POST':
        try:
            order = request.get_json()
            if isinstance(order, list):
                # 更新 region_order.json
                with open('region_order.json', 'w', encoding='utf-8') as f:
                    json.dump({'order': order}, f, ensure_ascii=False, indent=4)
                # 更新 plc_reader 的順序
                if plc_reader:
                    plc_reader.region_order = order
                return jsonify({'status': 'success', 'message': '區域順序已更新'})
            return jsonify({'status': 'error', 'message': '無效的請求格式'}), 400
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)}), 500
    else:
        # 返回當前順序
        if plc_reader and plc_reader.region_order:
            return jsonify(plc_reader.region_order)
        return jsonify(list(plc_reader.plc_tag.keys()) if plc_reader else [])

@app.route('/api/history/<region_name>')
def api_history(region_name):
    """API: 獲取歷史數據"""
    try:
        conn = psycopg2.connect(
            host='192.168.60.12',
            user='zhangmen',
            password='54685508',
            dbname='zmb'
        )
        cur = conn.cursor()
        
        # 根據區域名稱映射到對應的資料表
        table_mapping = {}
        
        # 從 plc_tag 動態建立映射
        if plc_reader:
            for region in plc_reader.plc_tag.keys():
                table_name = region_to_table(region)
                if table_name:
                    table_mapping[region] = table_name
        
        table_name = table_mapping.get(region_name)
        
        if not table_name:
            return jsonify({"error": f"區域 {region_name} 沒有對應的資料表"}), 404
        
        # 查詢最近 100 筆記錄
        cur.execute(f"""
            SELECT * FROM {table_name}
            ORDER BY timestamp DESC
            LIMIT 100
        """)
        
        columns = [desc[0] for desc in cur.description]
        rows = cur.fetchall()
        
        result = []
        for row in rows:
            record = dict(zip(columns, row))
            # 轉換 timestamp 為字串
            if 'timestamp' in record and record['timestamp']:
                record['timestamp'] = record['timestamp'].isoformat()
            result.append(record)
        
        cur.close()
        conn.close()
        
        return jsonify({
            "region": region_name,
            "data": result,
            "count": len(result)
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def region_to_table(region):
    """將區域名稱轉換為資料表名稱"""
    # 特殊處理已知區域
    special_mapping = {
        'Hot Water': 'plc_hotwater',
        'Mash/Lauter': 'plc_mashlauter',
        'Wort Kettle': 'plc_wortkettle',
        'Ice Water': 'plc_icewater',
        'Glycol#1': 'plc_glycol1',
        'Glycol#2': 'plc_glycol2',
    }
    
    if region in special_mapping:
        return special_mapping[region]
    
    # FV 區域映射
    if region.startswith('FV#'):
        fv_num = region.replace('FV#', '')
        return f'plc_fv{fv_num}'
    
    return None

@app.route('/alerts')
def alerts():
    """警報頁面"""
    regions = list(plc_reader.plc_tag.keys()) if plc_reader else []
    return render_template('alerts.html', regions=regions)

@app.route('/settings')
def settings():
    """設定頁面"""
    settings = load_settings()
    return render_template('settings.html',
                         general_config=settings.get('general', {}),
                         plc_config=settings.get('plc', {}),
                         db_config=settings.get('database', {}),
                         gs_config=settings.get('google_sheets', {}),
                         beername_config=settings.get('beername', {}))

@app.route('/api/settings/general', methods=['POST'])
def api_settings_general():
    """API: 儲存一般設定"""
    try:
        settings = load_settings()
        settings['general'] = request.get_json()
        if save_settings(settings):
            apply_settings_to_runtime(settings)
            return jsonify({'status': 'success', 'message': '一般設定已儲存'})
        return jsonify({'status': 'error', 'message': '儲存失敗'}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/settings/plc', methods=['POST'])
def api_settings_plc():
    """API: 儲存 PLC 設定"""
    try:
        settings = load_settings()
        settings['plc'] = request.get_json()
        if save_settings(settings):
            apply_settings_to_runtime(settings)
            return jsonify({'status': 'success', 'message': 'PLC 設定已儲存，已嘗試即時套用'})
        return jsonify({'status': 'error', 'message': '儲存失敗'}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/settings/database', methods=['POST'])
def api_settings_database():
    """API: 儲存資料庫設定"""
    try:
        settings = load_settings()
        settings['database'] = request.get_json()
        if save_settings(settings):
            apply_settings_to_runtime(settings)
            return jsonify({'status': 'success', 'message': '資料庫設定已儲存，已即時套用'})
        return jsonify({'status': 'error', 'message': '儲存失敗'}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/settings/google-sheets', methods=['POST'])
def api_settings_google_sheets():
    """API: 儲存 Google Sheets 設定"""
    try:
        settings = load_settings()
        settings['google_sheets'] = request.get_json()
        if save_settings(settings):
            apply_settings_to_runtime(settings)
            return jsonify({'status': 'success', 'message': 'Google Sheets 設定已儲存，已即時套用'})
        return jsonify({'status': 'error', 'message': '儲存失敗'}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/settings/beername', methods=['POST'])
def api_settings_beername():
    """API: 儲存酒款名稱設定"""
    try:
        settings = load_settings()
        beername_config = request.get_json()
        if not isinstance(beername_config, dict):
            return jsonify({'status': 'error', 'message': '請求格式錯誤'}), 400

        ranges = beername_config.get('ranges', [])
        if beername_config.get('enabled', False) and (not isinstance(ranges, list) or len(ranges) == 0):
            return jsonify({'status': 'error', 'message': '啟用酒款輪播時，至少需要一個啟用的範圍'}), 400

        settings['beername'] = beername_config
        if save_settings(settings):
            apply_settings_to_runtime(settings)
            return jsonify({'status': 'success', 'message': '酒款名稱設定已儲存，已即時套用'})
        return jsonify({'status': 'error', 'message': '儲存失敗'}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/settings/test-plc', methods=['GET', 'POST'])
def api_test_plc():
    """API: 測試 PLC 連接"""
    try:
        if request.method == 'POST':
            plc_config = request.get_json() or {}
        else:
            settings = load_settings()
            plc_config = settings.get('plc', {})
        plc_ip = plc_config.get('plc_ip', '192.168.60.201')
        plc_rack = plc_config.get('plc_rack', 0)
        plc_slot = plc_config.get('plc_slot', 2)
        
        test_client = snap7.client.Client()
        test_client.connect(plc_ip, plc_rack, plc_slot)
        
        # 嘗試讀取 CPU 狀態
        cpu_info = test_client.get_cpu_info()
        test_client.disconnect()
        
        return jsonify({'status': 'success', 'message': f'成功連接到 {plc_ip}'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/settings/test-database', methods=['GET', 'POST'])
def api_test_database():
    """API: 測試資料庫連接"""
    try:
        if request.method == 'POST':
            db_config = request.get_json() or {}
        else:
            settings = load_settings()
            db_config = settings.get('database', {})
        db_host = db_config.get('db_host', '192.168.60.12')
        db_user = db_config.get('db_user', 'zhangmen')
        db_password = db_config.get('db_password', '54685508')
        db_database = db_config.get('db_database', 'zmb')
        
        conn = psycopg2.connect(
            host=db_host,
            user=db_user,
            password=db_password,
            dbname=db_database
        )
        cur = conn.cursor()
        cur.execute('SELECT version();')
        result = cur.fetchone()
        cur.close()
        conn.close()
        
        return jsonify({'status': 'success', 'message': f'成功連接到 {db_host}'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@socketio.on('connect')
def handle_connect():
    """SocketIO 連接事件"""
    print('客戶端已連接')
    emit('connected', {'status': 'connected'})

@socketio.on('disconnect')
def handle_disconnect():
    """SocketIO 斷開事件"""
    print('客戶端已斷開')

def start_background_updater():
    """啟動背景數據更新器"""
    import threading
    thread = threading.Thread(target=plc_data_updater)
    thread.daemon = True
    thread.start()

if __name__ == '__main__':
    # 初始化 PLC 讀取器
    init_plc_reader()
    
    # 啟動背景更新器
    start_background_updater()

    # 透過環境變數控制埠號，容器預設使用 8001
    port = int(os.environ.get('PORT', '8001'))

    # 啟動 Flask 伺服器（use_reloader=False 防止 Werkzeug 啟動第二個進程重複寫入）
    socketio.run(app, host='0.0.0.0', port=port, debug=True, use_reloader=False, allow_unsafe_werkzeug=True)
