# -*- coding: utf-8 -*-
import xlwt
from tempfile import TemporaryFile
book = xlwt.Workbook()
sheet1 = book.add_sheet('sheet1')
fin=open('teraterm.log',errors='ignore')
d=[]
l=[]
lis=[]
#Enter PRU ID
#info charge(Usually use) :FD:89:8A:DE:04:23
#powerwow (Usually use(Android port))   :  C2:76:C6:56:69:6C
pru_id="FD:89:8A:DE:04:23"
#Count 
j=0
#handle log file 
for line in fin:
    li=line.split("\t")
    for i in range(0,len(li)):
        li[i]=li[i].strip(" \n")
#        print(li)
        if pru_id in li[i]:
            #d.append(int(li[6]))
            #l.append(int(li[3]))
            #j+=1
            print(li)
            #print("Loading... " +str(j), "record")
        if 'OVP' in li[i]:
            lis.append(li)
            print(lis)
#store into excel
"""
supersecretdata1=d
supersecretdata2=l
for i,e in enumerate(supersecretdata1):
    sheet1.write(i,1,e)
for i,e in enumerate(supersecretdata2):
    sheet1.write(i,0,e)
name = "random.xls"
book.save(name)
book.save(TemporaryFile())
avr=sum(d)/len(d)
avr1=sum(l)/len(l)
print("-------------------------")
print('Total Data: ' +str(len(l)))
print('Average Tmp : '  +str(avr))
print('Average Iout : ' +str(avr1))
fin.close()
#print(l)
#print(d)

"""

