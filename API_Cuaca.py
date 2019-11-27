import requests
key = '&appid=459d9e04c30f301451da1b16999bfc5b'
kota = input('Ketik kota: ')
url = f'http://api.openweathermap.org/data/2.5/weather?q={kota}{key}'
data = requests.get(url)
weather = data.json()['weather']
main = data.json()['main']
wind = data.json()['wind']
sys = data.json()['sys']


import xlsxwriter
file = xlsxwriter.Workbook(f'cuaca di {kota}.xlsx')
sheet = file.add_worksheet(f'cuaca di {kota}')

sheet.write(0, 0, 'No')
sheet.write(0, 1, 'Cuaca')
sheet.write(0, 2, 'Suhu')
sheet.write(0, 3, 'Tekanan Udara')
sheet.write(0, 4, 'Kecepatan Angin')
sheet.write(0, 5, 'Sunrise')
sheet.write(0, 6, 'Sunset')

list = []

no = 1
list.append(no)

for i in weather:
    list.append(i['main'])

main1 = []
main1.append(main)
for j in main1:
    list.append(str(round(j['temp']-273.15)))
    list.append(str(j['pressure']))

wind1 = []
wind1.append(wind)
for k in wind1:
    list.append(str(round(k['speed'])))

sunrise = sys['sunrise']
sunset = sys['sunset']
from datetime import datetime
sunrise1 = datetime.utcfromtimestamp(int(sunrise))
sunset1 = datetime.utcfromtimestamp(int(sunset))
sunrise2 = int(sunrise1.strftime('%H'))+7-24
sunset2 = int(sunset1.strftime('%H'))+7-12
list.append(str(sunrise2))
list.append(str(sunset2))

myList = []
myList.append(list)

row = 1
for a,b,c,d,e,f,g in myList:
    sheet.write(row, 0, a)
    sheet.write(row, 1, b)
    sheet.write(row, 2, c)
    sheet.write(row, 3, d)
    sheet.write(row, 4, e)
    sheet.write(row, 5, f)
    sheet.write(row, 6, g)
 
file.close()


judul = ['No', 'Cuaca', 'Suhu', 'Tekanan Udara', 'Kecepatan Angin', 'Sunrise', 'Sunset']
hasil = []
for i in myList:
    isi_csv = dict(zip(judul, i))
    hasil.append(isi_csv)

# CONVERT TO CSV
import csv
with open(f'Cuaca di kota {kota}.csv', 'w', newline='') as y:
    kolom = judul
    tulis = csv.DictWriter(y, fieldnames=kolom)
    tulis.writeheader()
    tulis.writerows(hasil)

# CONVERT TO JSON
import json
with open(f'Cuaca di kota {kota}.json', 'w') as x:
    json.dump(hasil, x)