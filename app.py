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
socketio = SocketIO(app, cors_allowed_origins="*", async_mode='threading')

# 全域變數
plc_client = None
plc_tag = {}
plc_data_cache = {}
data_lock = Lock()
running = True

class ZMBPlcReader:
    """PLC 數據讀取類別"""
    
    def __init__(self):
        self.gs_key = 'zmb54685508-c88132768091.json'
        self.sql_host = '192.168.60.12'
        self.sql_user = 'zhangmen'
        self.sql_password = '54685508'
        self.sql_database = 'zmb'
        self.plc_ip = "192.168.60.201"
        self.plc_rack = 0
        self.plc_slot = 2
        
        # 載入 PLC 標籤配置
        self.load_plc_tags()
        
        # 初始化 PLC 客戶端
        self.client = snap7.client.Client()
        try:
            self.client.connect(self.plc_ip, self.plc_rack, self.plc_slot)
            print(f"成功連接到 PLC: {self.plc_ip}")
        except Exception as e:
            print(f"連接 PLC 失敗：{e}")
    
    def load_plc_tags(self):
        """載入 PLC 標籤配置"""
        try:
            with open('plc_tag.json', encoding='utf-8', errors='ignore') as f:
                self.plc_tag = json.load(f, strict=False)
        except Exception as e:
            print(f"載入 PLC 標籤失敗：{e}")
            self.plc_tag = {}
    
    def read_plc_value(self, tag):
        """從 PLC 讀取單一標籤的值"""
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
            print(f"讀取標籤 {tag} 失敗：{e}")
            return None
    
    def get_all_plc_data(self):
        """讀取所有 PLC 數據"""
        global plc_data_cache
        
        with data_lock:
            data = {}
            
            for region, tags in self.plc_tag.items():
                region_data = {"timestamp": datetime.datetime.now().isoformat()}
                
                for tag_name, tag_address in tags.items():
                    value = self.read_plc_value(tag_address)
                    if value is not None:
                        region_data[tag_name] = value
                
                data[region] = region_data
            
            plc_data_cache = data
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
    
    while running:
        try:
            if plc_reader:
                data = plc_reader.get_all_plc_data()
                # 發送 SocketIO 事件
                socketio.emit('plc_update', data)
        except Exception as e:
            print(f"更新 PLC 數據時發生錯誤：{e}")
        
        time.sleep(5)  # 每 5 秒更新一次

@app.route('/')
def index():
    """首頁"""
    return render_template('index.html')

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
        return jsonify(plc_data_cache)

@app.route('/api/plc/region/<region_name>')
def api_plc_region(region_name):
    """API: 獲取特定區域的 PLC 數據"""
    global plc_data_cache
    with data_lock:
        if region_name in plc_data_cache:
            return jsonify(plc_data_cache[region_name])
        return jsonify({"error": f"區域 {region_name} 不存在"}), 404

@app.route('/api/regions')
def api_regions():
    """API: 獲取所有可用區域列表"""
    global plc_tag
    if plc_reader:
        return jsonify(list(plc_reader.plc_tag.keys()))
    return jsonify([])

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
        table_mapping = {
            'Hot Water': 'plc_hotwater',
            'Mash/Lauter': 'plc_mashlauter',
            'Wort Kettle': 'plc_wortkettle',
            'Ice Water': 'plc_icewater',
            'Glycol#1': 'plc_glycol1',
            'Glycol#2': 'plc_glycol2',
        }
        
        # FV 區域映射
        for i in range(1, 23):
            table_mapping[f'FV#{i}'] = f'plc_fv{i}'
        
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

@app.route('/alerts')
def alerts():
    """警報頁面"""
    regions = list(plc_reader.plc_tag.keys()) if plc_reader else []
    return render_template('alerts.html', regions=regions)

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
    
    # 啟動 Flask 伺服器
    socketio.run(app, host='0.0.0.0', port=8001, debug=True)
