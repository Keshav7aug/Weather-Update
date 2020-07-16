from openpyxl import load_workbook
import requests
import time
import os
Flg=True
try:
    import tkinter as tk
except:
    Flg=False
def get_temp(city,unit='m'):
    url = "http://api.weatherstack.com/current"
    prm={'access_key':'5430d6683e036b75deb21cb722950136','query':city,'units':unit}
    ans = requests.get(url,prm).json()
    try:
        temp,humidity = ans['current']['temperature'],ans['current']['humidity']
    except:
        return False,False
    return temp,humidity
wb = load_workbook('Record.xlsx')
sheet = wb.active
if Flg:
    window = tk.Tk()
    window.title("Weather")
    window.configure(bg='black')
    flg=True
else:
    flg=False

labels = []
z=0
i=2
j=0
colrs=['black','grey']
if Flg:
    for _,i in enumerate(["City Name", "Temp", "Humidity"]):
        tk.Label(window,text = i ,bg=colrs[_%2],fg='white',font=("gothic", 44)).grid(column=_,row=0)
while True:
    j+=1
    i=2
    while True:
        citylbl = []
        unit='m'
        cv='A{}'.format(i)
        city=sheet[cv].value
        if city==None:
            break
        tmpcell = 'B{}'.format(i)
        humdcell = 'C{}'.format(i)
        unitcell = 'D{}'.format(i)
        upcell = 'E{}'.format(i)
        if sheet[unitcell].value == 'f':
            unit='f'
            dUnit = 'f'
        else:
            dUnit = 'c'
        if str(sheet[upcell].value) == '1':
            newtemp,newHumidity = get_temp(city,unit)
            
            if newtemp!=False:
                sheet[tmpcell].value = newtemp
                sheet[humdcell].value = newHumidity
            else:
                newtemp,newHumidity = "NA","NA"
            print(".",end='')
        else:
            newtemp = sheet[tmpcell].value
            newHumidity = sheet[humdcell].value
        try:
            x=window.state()
        except:
            flg=False
        if flg:
            Temp = "{}{}".format(newtemp,dUnit.upper())
            data = [city,Temp,newHumidity]
            for _,d in enumerate(data):
                citylbl.append(tk.Label(window,text = d,bg=colrs[_%2],fg='white'))
            for _,c in enumerate(citylbl):
                c.config(font=("gothic", 44))
                c.grid(column=_,row=i+1)
        i+=1
        wb.save("Record.xlsx")
        if flg:
            window.update_idletasks()
            window.update()
    time.sleep(1) ## update interval 1 seconds
window.mainloop()