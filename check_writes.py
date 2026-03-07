import psycopg2, json, pygsheets, datetime

with open('/app/settings.json') as f:
    s = json.load(f)

# 檢查資料庫最新寫入
print("=== 資料庫最新寫入 ===")
db = s['database']
conn = psycopg2.connect(host=db['db_host'], user=db['db_user'], password=db['db_password'], dbname=db['db_database'])
cur = conn.cursor()
tables = [('Hot Water','plc_hotwater'), ('FV#1','plc_fv1'), ('Glycol#1','plc_glycol1'), ('Wort Kettle','plc_wortkettle')]
for name, table in tables:
    cur.execute('SELECT timestamp FROM ' + table + ' ORDER BY timestamp DESC LIMIT 1')
    row = cur.fetchone()
    ts = str(row[0]) if row else '無資料'
    print(name + ': ' + ts)
cur.close()
conn.close()

# 檢查 Google Sheets
print("\n=== Google Sheets 最新寫入 ===")
gs = s['google_sheets']
gc = pygsheets.authorize(service_account_file=gs['gs_key'])
title = 'ZMB-' + str(datetime.date.today())[:-3]
try:
    ss = gc.open(title)
    wks = ss.worksheet_by_title('Hot Water')
    rows = len(wks.get_col(1, include_tailing_empty=False))
    last = wks.get_values('A' + str(rows), 'B' + str(rows))
    print('Hot Water 最後一列 (' + str(rows) + '): ' + str(last))
    print('Google Sheets 正常')
except Exception as e:
    print('錯誤: ' + str(e))
