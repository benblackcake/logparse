# -*- coding: utf-8 -*-
import xlwt
from tempfile import TemporaryFile
book = xlwt.Workbook()
sheet1 = book.add_sheet('sheet1')
fin=open('console_20170728-171614.log',errors='ignore')
Vin=[]
Iin=[]
Icoil=[]
lis=[]
#Enter PRU ID
#info charge(Usually use) :FD:89:8A:DE:04:23
#powerwow (Usually use(Android port))   :  C2:76:C6:56:69:6C

#Count 

#handle log file 
def logpoen():
    j=0
    for line in fin:
        li=line.split("\t")
        for i in range(0,len(li)):
            li[i]=li[i].strip(" \n")
            #print(li)
            try:
                
                if 'Ptrans' in li[i]:
                    Vin.append(int(li[6]))#Vin 
                    Iin.append(int(li[7]))#Iin
                    Icoil.append(int(li[11]))#Icoil
                    j+=1
                    print("Loading... " +str(j), "record")
            except IndexError:
                pass
    return Vin,Iin,Icoil
#store into excel
def store_excel(Vin,Iin,Icoil):
    for i,e in enumerate(Vin):
        sheet1.write(i,0,e)
    for i,e in enumerate(Iin):
        sheet1.write(i,1,e)
    for i,e in enumerate(Icoil):
        sheet1.write(i,2,e)
    name = "Vin Iin Icoil information.xls"
    book.save(name)
    book.save(TemporaryFile())
    print("-------------------------")
    print('Your Excel File : ',name,'Has Been Saved')
def decide_pass(Vin,Iin,Icoil):
    try:
        
        avr=sum(Vin)/len(Vin)
        avr1=sum(Iin)/len(Iin)
        avr2=sum(Icoil)/len(Icoil)
        print("-------------------------")
        if avr < 11500:
            print("Vin Error")
        elif avr1 < 800:
            print("Iin Error")
        elif avr2 < 400:
            print("Icoil Error")
        else:
            print("Pass")
        print("-------------------------")
        print('Total Data: ' +str(len(Vin)))
        print('Average Vin : '  +str(avr))
        print('Average Iin : ' +str(avr1))
        print('Average Icoil : ' +str(avr2))
    except ZeroDivisionError:
        avr=0
        avr1=0
        avr2=0
        print("-------------------------")
        print('Total Data: ' +str(len(Vin)))
        print('Average Vin : '  +str(avr))
        print('Average Iin : ' +str(avr1))
        print('Average Icoil : ' +str(avr2))

logpoen()
store_excel(Vin,Iin,Icoil)
decide_pass(Vin,Iin,Icoil)
fin.close()
#print(l)
#print(d)


