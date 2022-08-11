#!/usr/bin/python


import Adafruit_DHT
import time
from datetime import datetime
from datetime import date
from openpyxl import Workbook

wb = Workbook()
sensor = Adafruit_DHT.DHT11

ws = wb.active
pin = 23

# Try to grab a sensor reading.  Use the read_retry method which will retry up
# to 15 times to get a sensor reading (waiting 2 seconds between each retry).
humidity, temperature = Adafruit_DHT.read_retry(sensor, pin)

for i in range(1, 6):
    if humidity is not None and temperature is not None:
            print('Temp={0:0.1f}*C  Humidity={1:0.1f}%'.format(temperature, humidity))
            now = datetime.now()
            ws['A'+str(i)] = str(now.strftime("%d"))
            ws['B'+str(i)] = str(now.strftime("%m"))
            ws['C'+str(i)] = str(now.strftime("%Y"))
            ws['D'+str(i)] = str(now.strftime("%H:%M:%S"))
            ws['E'+str(i)] = str(temperature+"C")
            ws['F'+str(i)] = str(humidity"%")            
            wb.save("temperature.xlsx")
            time.sleep(1800)
    else:
        print('Failed to get reading. Try again!')
