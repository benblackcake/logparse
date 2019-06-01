# -*- coding: utf-8 -*-
import xlwt
from tempfile import TemporaryFile
book = xlwt.Workbook()
sheet1 = book.add_sheet('sheet1')
fin=open('console_20170727-110655(40min).log',errors='ignore')
d=[]
l=[]
k=[]
lis=[]
#Enter PRU ID
#info charge(Usually use) :FD:89:8A:DE:04:23
#powerwow (Usually use(Android port))   :  C2:76:C6:56:69:6C
pru_id="C2:76:C6:56:69:6C"
#Count 
j=0
#handle log file 
for line in fin:
    li=line.split("\t")
    for i in range(0,len(li)):
        li[i]=li[i].strip(" \n")
        #print(li)
        try:
            if 'Ptrans' in li[i]:
                d.append(int(li[6]))#Vin 
                l.append(int(li[7]))#Iin
                k.append(int(li[11]))#Icoil
                j+=1
                print("Loading... " +str(j), "record")
        except IndexError:
            pass
            

#store into excel
supersecretdata1=d
supersecretdata2=l
supersecretdata3=k
for i,e in enumerate(supersecretdata1):
    sheet1.write(i,0,e)
for i,e in enumerate(supersecretdata2):
    sheet1.write(i,1,e)
for i,e in enumerate(supersecretdata3):
    sheet1.write(i,2,e)
name = "Vin Iin Icoil information.xls"
book.save(name)
book.save(TemporaryFile())
avr=sum(d)/len(d)
avr1=sum(l)/len(l)
avr2=sum(k)/len(k)
print("-------------------------")
print('Total Data: ' +str(len(l)))
print('Average Vin : '  +str(avr))
print('Average Iin : ' +str(avr1))
print('Average Icoil : ' +str(avr2))
fin.close()
#print(l)
#print(d)


