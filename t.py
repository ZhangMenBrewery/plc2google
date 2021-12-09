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
    
    def tag_value(self, tag):
        if tag.startswich('DB'):
            pass
        elif tag.startswich('Q'):
            pass
        else:
            pass

    # ss mean spreadsheet
    def open_ss(self):
        # gc means google client
        gc = pygsheets.authorize(service_account_file=self.__gs_key)
        try:
            ss = gc.open(self.__gs_title)
            print(f"Opened spreadsheet with url:{ss.url}")
            for member in self.__zmb_group:
                ss.share(member, role='writer', type='user')
            ss.share('', role='reader', type='anyone')
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
            wks = ss.worksheet_by_title(item)
            rows = len(wks.get_col(1, include_tailing_empty=False))+1
            data = []
            data.append(time.strftime('%m/%d/%Y %H:%M:%S', time.localtime()))
            data.append(snap7.util.get_real(client.db_read(40, 8, 4), 0))  # DB40.DBD8
            data.append(snap7.util.get_real(
                client.db_read(112, 0, 4), 0))  # DB112.DBD0
        data.append(1 if snap7.util.get_bool(
            client.ab_read(0, 11), 4, 3) == 'TRUE' else 0)  # Q4.3
        data.append(snap7.util.get_real(
            client.db_read(40, 88, 4), 0))  # DB40.DBD88
        data.append(1 if snap7.util.get_bool(
            client.ab_read(0, 11), 6, 0) == 'TRUE' else 0)  # Q6.0
        data.append(snap7.util.get_dint(client.db_read(40, 0, 4), 0))  # DB40.DBD0
        wks_hw.update_values(f"A{rows}:G{rows}", [data, ])

def main():
        plc_obj = zmb_plc()
        ss_obj = plc_obj.open_ss()
        plc_obj.write_to_sheet(ss_obj)
        #plc_obj.delete_all()

if __name__ == '__main__':
    main()


'''


            wks_hw = ss.add_worksheet(title="Hot Water", rows=10000, cols=10)
            wks_hw.update_values('A1:G1', [
                                ['Time', 'HW_Temp', 'HW_SP', 'HW_Valve', 'HW_Vol', 'HW_Pump', 'HW_Flow'], ])
            wks_hw.frozen_rows = 1
            wks_ml = ss.add_worksheet(title="Mash/Lauter", rows=10000, cols=10)
            wks_ml.update_values('A1:I1', [['Time', 'ML_Temp', 'MT_SP', 'ML_Valve',
                                'ML_Pump', 'ML_Pump_Sp', 'ML_Agitator', 'ML_Agi_Sp', ' ML_Flow'], ])
            wks_ml.frozen_rows = 1
            wks_wk = ss.add_worksheet(title="Wort Kettle", rows=10000, cols=10)
            wks_wk.update_values('A1:J1', [['Time', 'WK_Temp', 'WK_SP', 'WK_Valve1', 'WK_Valve2',
                                'WK_Chimney', 'WK_Pump', 'WK_Pump_Sp', ' WK_Flow', ' WK_Flow_Vol'], ])
            wks_wk.frozen_rows = 1
            wks_iwt = ss.add_worksheet(title="Ice Water", rows=10000, cols=10)
            wks_iwt.update_values('A1:C1', [['Time', 'IWT_Vol', 'IWT_Pump'], ])
            wks_iwt.frozen_rows = 1
            wks_gy1 = ss.add_worksheet(title="Gloycol#1", rows=10000, cols=10)
            wks_gy1.update_values(
                'A1:E1', [['Time', 'GY1_Temp', 'GY1_SP', 'GY1_Power', 'GY1_Pump'], ])
            wks_gy1.frozen_rows = 1
            wks_gy2 = ss.add_worksheet(title="Gloycol#2", rows=10000, cols=10)
            wks_gy2.update_values(
                'A1:E1', [['Time', 'GY2_Temp', 'GY2_SP', 'GY2_Power', 'GY2_Pump'], ])
            wks_gy2.frozen_rows = 1
            for num in range(1, 23):
                wks_fv = ss.add_worksheet(
                    title="FV#" + str(num), rows=10000, cols=10)
                wks_fv.update_values('A1:F1', [['Time', 'FV'+str(num)+'_Temp', 'FV'+str(
                    num)+'_SP', 'FV'+str(num)+'_Valve1', 'FV'+str(num) + ' GY_Temp'], ])
                wks_fv.frozen_rows = 1
            ss.del_worksheet(ss.sheet1)
        return ss


# Try to open the Google sheet based on its title and if it fails, create it                                                                                                                                                                                                                                                          
try:
    sheet = gc.open(sheet_title)
    print(f"Opened spreadsheet with id:{sheet.id} and url:{sheet.url}")
except pygsheets.SpreadsheetNotFound as error:
    # Can't find it and so create it                                                                                                                                                                                                                                                                                                  
    res = gc.sheet.create(sheet_title)
    sheet_id = res['spreadsheetId']
    sheet = gc.open_by_key(sheet_id)
    print(f"Created spreadsheet with id:{sheet.id} and url:{sheet.url}")

    # Share with self to allow to write to it                                                                                                                                                                                                                                                                                         
    sheet.share('zhangmenbrewery@gmail.com', role='writer', type='user')

    # Share to all for reading                                                                                                                                                                                                                                                                                                        
    sheet.share('', role='reader', type='anyone')

# Write something into it                                                                                                                                                                                                                                                                                                             
wks = sheet.sheet1
wks.update_value('A1', "something")
'''


'''
order_url = 'https://docs.google.com/spreadsheets/d/1gDwQCBAg8441lnzl-ywy984rugpRmgLzyFHyJc0pZFE/edit#gid=642205089'

#ss means spreadsheet
ss_order = gc.open_by_url(order_url)

#wks means worksheet
wks_order = ss_order.worksheet_by_title('店面訂購頁面')


#讀取資料至pd
while True:
    try:
        pd_orders = wks_order.get_as_df(start = 'N2', end=(100,30), index_colum=0, empty_value='', include_tailing_empty=False)[2:]
        stock = pd_orders.iloc[:,1]
        for item in stock:
            if int(item) < 0:
                print(datetime.datetime.now(pytz.timezone('Asia/Taipei')),': out of stock')
                wks_order.set_dataframe(old_orders.iloc[:,3:].replace(0, numpy.nan), 'Q5', copy_index=False, copy_head=False, nan='')
                break
        old_orders = pd_orders.replace(0, numpy.nan)
        time.sleep(1)
    except:
        time.sleep(1)

'''
