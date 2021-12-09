import pygsheets
import time
import datetime
import numpy
import pytz
import snap7
import json
import string


class zmb_plc():
    def __init__(self) -> None:
        self.__gs_key = 'zmb54685508-c88132768091.json'
        self.__gs_title = 'ZMB-' + str(datetime.date.today())[:-3]
        self.__zmb_group = [
            'zhangmenbrewery@gmail.com', 'chunkai721@gmail.com']
        self.__plc_tag = {}

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
                    if tag[1].startswith('DBD'):#db client.db_read(int(tag[0][2:]), int(tag[1][3:]), 4)
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

def main():
        while True:
            try:
                plc_obj = zmb_plc()
                ss_obj = plc_obj.open_ss()
                plc_obj.write_to_sheet(ss_obj)
            except Exception as error:
                print(f"{time.strftime('%m/%d/%Y %H:%M:%S', time.localtime())} > {error} ")
                time.sleep(60)
            print(f"{time.strftime('%m/%d/%Y %H:%M:%S', time.localtime())} > Sleeping 300sec")
            time.sleep(300)

if __name__ == '__main__':
    main()
    #zmb_plc().delete_all()
