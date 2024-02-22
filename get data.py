

import time
import requests
from openpyxl import Workbook
from pandas import read_excel ,DataFrame ,concat 
import datetime


inscode = '65576885779918210'


def get_date_format(date):
    year_str=str(date.year)
    month_str=str(date.month)
    if len(month_str)==1:
        month_str=f"0{date.month}"
    day_str=str(date.day)
    if len(day_str)==1:
        day_str=f"0{date.day}"

    return year_str+month_str+day_str

list_of_dates=[]
start_date=datetime.date(year=2024, month=1, day=29)
day_periods=2
for i in range(0, day_periods):
    cur_delta=datetime.timedelta(days=i)
    cur_date= start_date + cur_delta
    date_format= get_date_format(cur_date)
    list_of_dates.append(date_format)
#######################################################
    

def get_data(inscode,date):

    
    adress = f'https://cdn.tsetmc.com/api/Trade/GetTradeHistory/{inscode}/{date}/false'
    header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36'}
    sotun=['ردیف','زمان','حجم','قیمت']

    while True:
        try:
            req1 = requests.get(url=adress , headers=header , timeout=15).text.split('},{')
            if len(req1) > 1 :
                break
            else:
                time.sleep(2)
                print('Error10')
        except:
                time.sleep(2)
                print('Error11')


    dictriz = {}

    for a in req1:

        radif = a.split('nTran":')[1].split(',')[0]

        zaman = a.split('hEven":')[1].split(',')[0]

        hajm = a.split('qTitTran":')[1].split(',')[0]

        gheymat = a.split('pTran":')[1].split('.')[0]

        dictriz[int(radif)] = [int(radif), int(zaman), int(hajm), int(gheymat)]
        

    ########################################################




    sarsotun =  ['زمان', 'ردیف', 'تعداد ف', 'حجم ف', 'قیمت ف', 'قیمت خ', 'حجم خ', 'تعداد خ' ]

    address = f'https://cdn.tsetmc.com/api/BestLimits/{inscode}/{date}'
    header = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36'}
    while True:
        try:
            req2 = requests.get(url=address , headers=header, timeout=15).text.split('},{')
            if len(req2) > 1 :
                break
            else:
                time.sleep(2)
                print('Error20')
        except:
            time.sleep(2)
            print('Error21')

    order_dict = {}
    shomarande = 0
    for a in req2:
        radif = int(a.split('number":')[1].split(',')[0])
        if radif != 1 :
         continue

        shomarande +=1

        zaman = int(a.split('hEven":')[1].split(',')[0])
        kharidtedad = int(a.split('zOrdMeDem":')[1].split(',')[0])
        kharidhajm = int(a.split('qTitMeDem":')[1].split(',')[0])
        kharidgheymat = int(a.split('pMeDem":')[1].split('.')[0])
        forushtedad = int(a.split('zOrdMeOf":')[1].split(',')[0])
        forushhajm = int(a.split('qTitMeOf":')[1].split(',')[0])
        forushgheymat = int(a.split('pMeOf":')[1].split('.')[0])
        order_dict[shomarande] = [zaman, radif, forushtedad, forushhajm, forushgheymat, kharidgheymat, kharidhajm, kharidtedad ]
        
       

    ########################################################

    for a in sorted(dictriz, reverse = False):
         for b in sorted(order_dict, reverse = False):
          if dictriz[a][1] <= order_dict[b][0]:
               dictriz[a] = dictriz[a] + [safghabl[4],safghabl[5]]
               break
               
          else:
               safghabl = order_dict[b]

    noemoamelat = ['','','']
    timemoamele = 0
    for c in sorted(dictriz , reverse = False):
        
        if timemoamele == dictriz[c][1] :
         if type(noemoamelat[0]) == int :
              noemoamelat = [dictriz[c][2] ,'','']
         elif type(noemoamelat[1]) == int :
              noemoamelat = ['', dictriz[c][2] ,'']
         elif type(noemoamelat[2]) == int :
              noemoamelat = ['','', dictriz[c][2]]

         
         dictriz[c] = dictriz[c] + noemoamelat + [int(date)]
         continue
        
        elif dictriz[c][3] != dictriz[c][4] and dictriz[c][3] != dictriz[c][5] :
         noemoamelat = ['','', dictriz[c][2]]	 
        elif dictriz[c][4] == dictriz[c][5]:
         noemoamelat = ['','',dictriz[c][2]]

        elif dictriz[c][3] == dictriz[c][4]:
         noemoamelat = [ dictriz[c][2] ,'','']

        elif dictriz[c][3] == dictriz[c][5]:
         noemoamelat = ['', dictriz[c][2] ,'']

        dictriz[c] = dictriz[c] + noemoamelat + [int(date)] 
        print(dictriz[c]) 
        timemoamele = dictriz[c][1] 

    excel3 = Workbook()
    excel3.active.append(['id','trade time',"volume",'last','sell', 'buy','sell_side','buy_side','unknown','day'])

   

    # for d in sorted(dictriz , reverse = False):
    #      excel3.active.append(dictriz[d])
         
    # excel3.save(f'C:/Users/MSI/Desktop/uni/1-payan/codes/vv.xlsx')
    
    return [excel3.active.append(dictriz[d]) for d in sorted(dictriz, reverse=False)]



for date in list_of_dates:
    adress = f'https://cdn.tsetmc.com/api/Trade/GetTradeHistory/{inscode}/{date}/false'
    header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36'}
    if len(requests.get(url=adress , headers=header , timeout=15).text.split('},{'))>1 :
            get_data(65576885779918210,date)
            excel3.save(f'C:/Users/MSI/Desktop/uni/1-payan/codes/vv.xlsx')
    else:
        print(f"Skipping invalid data in date {date}")


