import pygsheets
import time
import datetime
import numpy
import pytz
import snap7
import json
import string

client = snap7.client.Client()
client.connect("192.168.60.201", 0, 2)
print((snap7.util.get_bool(client.ab_read(0, 11), 4, 0))) #Q4.4