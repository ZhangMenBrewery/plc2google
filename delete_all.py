import pygsheets, time, datetime, numpy, pytz

#gc means google client
gc = pygsheets.authorize(service_account_file='zmb54685508-c88132768091.json')

sheets = gc.open_all()
for sheet in sheets:
    print(sheet.url)
    sheet.delete()
