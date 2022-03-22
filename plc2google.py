from urllib import request
import mysql.connector
from mysql.connector import Error
import pygsheets
import time
import datetime
import snap7
import json
import string
import os
import requests


class zmb_plc():
    def __init__(self) -> None:
        self.__gs_key = 'zmb54685508-c88132768091.json'
        self.__gs_title = 'ZMB-' + str(datetime.date.today())[:-3]
        self.__zmb_group = [
            'zhangmenbrewery@gmail.com', 'chunkai721@gmail.com']
        self.__plc_tag = {}

        self.__sql_host = 'localhost'
        self.__sql_user = 'zhangmen'
        self.__sql_password = '54685508'
        path = 'sql_insert.json'
        with open(path,encoding='utf-8', errors='ignore') as f:
            self.__sql_insert_tag = json.load(f, strict=False)

    # ss mean spreadsheet
    def open_ss(self):
        # gc means google client
        gc = pygsheets.authorize(service_account_file=self.__gs_key)
        try:
            ss = gc.open(self.__gs_title)
            print(f"Opened spreadsheet with url:{ss.url}")
            path = 'plc_tag.json'
            with open(path,encoding='utf-8', errors='ignore') as f:
                self.__plc_tag = json.load(f, strict=False)
        except pygsheets.SpreadsheetNotFound as error:
            res = gc.sheet.create(self.__gs_title)  # Please set the new Spreadsheet name.
            print(f"Created spreadsheet with url:{res['spreadsheetUrl']}")
            ss = gc.open_by_key(res['spreadsheetId'])
            for member in self.__zmb_group:
                ss.share(member, role='writer', type='user')
            ss.share('', role='reader', type='anyone')
            path = 'plc_tag.json'
            with open(path,encoding='utf-8', errors='ignore') as f:
                self.__plc_tag = json.load(f, strict=False)
            # create worksheet
            for item in self.__plc_tag.keys():
                wks = ss.add_worksheet(title=item, rows=10000, cols=10)
                letter = dict(zip(range(1,27),string.ascii_uppercase))[len(self.__plc_tag[item])+1]
                header = ['Time'] + list(self.__plc_tag[item].keys())
                wks.update_values(f"A1:{letter}1", [header,])
                wks.frozen_rows = 1
            ss.del_worksheet(ss.sheet1)
        return ss

    def delete_all(self):
        gc = pygsheets.authorize(service_account_file=self.__gs_key)
        sss = gc.open_all()
        for sheet in sss:
            sheet.delete()
        print('All sheetes have been deleted.')


    def write_to_sheet(self,ss):
        client = snap7.client.Client()
        client.connect("192.168.60.201", 0, 2)
    #    plc_di = client.eb_read(0, 7)
    #    print(snap7.util.get_bool(plc_di,6,1))
    #    plc_do = client.ab_read(0 ,11)
    #    print(snap7.util.get_bool(plc_do, 10, 5))
        for item in self.__plc_tag.keys():
            print(f"{time.strftime('%m/%d/%Y %H:%M:%S', time.localtime())} > {item}")
            wks = ss.worksheet_by_title(item)
            rows = len(wks.get_col(1, include_tailing_empty=False))+1
            data = []
            data.append(time.strftime('%m/%d/%Y %H:%M:%S', time.localtime()))
            for value in self.__plc_tag[item].values():
                tag = value.split('.')
                if tag[0].startswith('DB'):
                    if tag[1].startswith('DBDDI'):#db client.db_read(int(tag[0][2:]), int(tag[1][3:]), 4)
                        data.append(snap7.util.get_dint(client.db_read(int(tag[0][2:]), int(tag[1][5:]), 4), 0))  # DB40.DBD8
                    elif tag[1].startswith('DBD'):#db client.db_read(int(tag[0][2:]), int(tag[1][3:]), 4)
                        data.append(snap7.util.get_real(client.db_read(int(tag[0][2:]), int(tag[1][3:]), 4), 0))  # DB40.DBD8
                    elif tag[1].startswith('DBX'):
                        data.append(1 if snap7.util.get_bool(client.db_read(int(tag[0][2:]), int(tag[1][3:]), 1), 0,int(tag[2])) == True else 0)  # DB40.DBX3.3
                    else:
                        pass
                elif tag[0].startswith('Q'):
                    data.append(1 if snap7.util.get_bool(client.ab_read(0, 11), int(tag[0][1:]), int(tag[1])) == True else 0)  # Q4.3
                else: # for digital input
                    pass
            wks.update_values(f"A{rows}:{dict(zip(range(1,27),string.ascii_uppercase))[len(self.__plc_tag[item])+1]}{rows}", [data, ])    

    #建立sql連線
    def connect_db(self, db):
        connection = None
        try:
            connection = mysql.connector.connect(
                host=self.__sql_host,
                user=self.__sql_user,
                passwd=self.__sql_password,
                database=db
            )
            print("Connection to MySQL DB successful")
        except Error as e:
            print(f"The error '{e}' occurred")
        return connection

    def write_to_sheet_sql(self,ss):
        client = snap7.client.Client()
        client.connect("192.168.60.201", 0, 2)
        token = 'ypWJFnAK7qffYVaHFsDZgXNfhc1RZs5DekSkKVB73kO'
    #    plc_di = client.eb_read(0, 7)
    #    print(snap7.util.get_bool(plc_di,6,1))
    #    plc_do = client.ab_read(0 ,11)
    #    print(snap7.util.get_bool(plc_do, 10, 5))

        conn = self.connect_db("zmb")
        cur = conn.cursor()

        for item in self.__plc_tag.keys():
            print(f"{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())} > {item}")
            wks = ss.worksheet_by_title(item)
            rows = len(wks.get_col(1, include_tailing_empty=False))+1
            data = []
            data.append(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
            for value in self.__plc_tag[item].values():
                tag = value.split('.')
                if tag[0].startswith('DB'):
                    if tag[1].startswith('DBDDI'):#db client.db_read(int(tag[0][2:]), int(tag[1][3:]), 4)
                        data.append(snap7.util.get_dint(client.db_read(int(tag[0][2:]), int(tag[1][5:]), 4), 0))  # DB40.DBD8
                    elif tag[1].startswith('DBD'):#db client.db_read(int(tag[0][2:]), int(tag[1][3:]), 4)
                        data.append(snap7.util.get_real(client.db_read(int(tag[0][2:]), int(tag[1][3:]), 4), 0))  # DB40.DBD8
                    elif tag[1].startswith('DBX'):
                        data.append(1 if snap7.util.get_bool(client.db_read(int(tag[0][2:]), int(tag[1][3:]), 1), 0,int(tag[2])) == True else 0)  # DB40.DBX3.3
                    else:
                        pass
                elif tag[0].startswith('Q'):
                    data.append(1 if snap7.util.get_bool(client.ab_read(0, 11), int(tag[0][1:]), int(tag[1])) == True else 0)  # Q4.3
                else: # for digital input
                    pass
            wks.update_values(f"A{rows}:{dict(zip(range(1,27),string.ascii_uppercase))[len(self.__plc_tag[item])+1]}{rows}", [data, ])
            sql_insert = self.__sql_insert_tag[item]
            # data.insert(0,'')s
            # print(data)
            # data[0] = datetime.datetime.strptime(data[0], '%m/%d/%Y %H:%M:%S')
            # data[0] = data[0].strftime('%Y-%m-%d %H:%M:%S')
            if str(item).startswith('FV'):
                data = data[:-1]
                if (data[3] or [4] == 1) and (abs(data[1]-data[2])>1) and datetime.datetime.now().minute%2==0:
                    msg = f'Waring: {item} Temperature: {data[2]} Setpoint:{data[1]}'
                    lineNotifyMessage(token=token, msg=msg)
            if str(item).startswith('Glycol#2'):
                if (data[3] == 0) and datetime.datetime.now().minute%2==0:
                    msg = f'Waring: {item} power id off'
                    lineNotifyMessage(token=token, msg=msg)
            if str(item).startswith('Glycol'):
                if (data[4] == 1) and (abs(data[1]-data[2])>1) and datetime.datetime.now().minute%2==0:
                    msg = f'Waring: {item} Temperature: {data[2]} Setpoint:{data[1]}'
                    lineNotifyMessage(token=token, msg=msg)
            cur.execute(sql_insert, tuple(data))
            conn.commit()
        cur.close()
        conn.close()

def mask(rawBytes):
    data = [ord(i) for i in rawBytes]
    length = len(rawBytes) + 128 if len(rawBytes) + 128 <= 254 else 254
    Bytes = [0x81, length]
    index = 2
    mask = os.urandom(4)
    for i in range(len(mask)):
        Bytes.insert(i + index, mask[i])        
    for i in range(len(data)):
        data[i] ^= mask[i % 4]
        Bytes.insert(i + index + 4, data[i])
    return bytes(Bytes)

def beername():
    gc = pygsheets.authorize(service_account_file='zmb54685508-c88132768091.json')
    ss_url = 'https://docs.google.com/spreadsheets/d/1QUh-ZHlJSFG0RhkmP0JT7WBxRUqC8EgCMwGn7lZVQas/edit#gid=877936540'
    #ss means spreadsheet
    ss = gc.open_by_url(ss_url)
    #wks means worksheet
    wks = ss.worksheet_by_title('發酵桶現況')
    #讀取資料至pd
    beername = wks.get_col(3)[1:23]
    client = snap7.client.Client()
    client.connect("192.168.60.201", 0, 2)
    for i in range(len(beername)):
        data = bytearray(256)
        snap7.util.set_string(data, 0, beername[i], 255)
        client.db_write(db_number=142, start=i*256, data=data)

def lineNotifyMessage(token, msg, img=None):
    """Send a LINE Notify message (with or without an image)."""
    URL = 'https://notify-api.line.me/api/notify'
    headers = {'Authorization': 'Bearer ' + token}
    payload = {'message': msg}
    files = {'imageFile': open(img, 'rb')} if img else None
    r = requests.post(URL, headers=headers, params=payload, files=files)
    if files:
        files['imageFile'].close()
    return r.status_code



def main():
        while True:
            try:
                plc_obj = zmb_plc()
                ss_obj = plc_obj.open_ss()
                plc_obj.write_to_sheet_sql(ss_obj)
                beername()
            except Exception as error:
                print(f"{time.strftime('%m/%d/%Y %H:%M:%S', time.localtime())} > {error} ")
                time.sleep(60)
            print(f"{time.strftime('%m/%d/%Y %H:%M:%S', time.localtime())} > Sleeping 300sec")
            time.sleep(300)

if __name__ == '__main__':
    main()
    #zmb_plc().delete_all()
