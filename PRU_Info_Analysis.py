# -*- coding: utf-8 -*-
import xlwt
from tempfile import TemporaryFile
book = xlwt.Workbook()
sheet1 = book.add_sheet('sheet1')
fin=open("test use.log",errors='ignore')
Tmp=[]
Iout=[]
Vout=[]

count_times=0#Count times
count_indexerror=0
count_valueerror=0
count_ovp=0
count_otp=0
#Enter PRU ID
#F4:19:0D:DF:F9:6A
#DC:A3:1F:2B:59:C9
pru_id="C2:76:C6:56:69:6C"
#info charge(Usually use) :FD:89:8DC:A3:1F:2B:59:C9A:DE:04:23
#powerwow (Usually use(Android port))   :  C2:76:C6:56:69:6C

#handle log file z
def logopen(pru_id):
    global count_times
    global count_ovp
    global count_otp
    global count_indexerror
    global count_valueerror
    for line in fin:
        li=line.split("\t")
        for i in range(0,len(li)):
            li[i]=li[i].strip(" \n")
            try:
                if pru_id in li[i]:
                    Tmp.append(int(li[6]))
                    Iout.append(int(li[3]))
                    Vout.append(int(li[2]))
                    count_times+=1
                    print("Loading... " +str(count_times), "record")
                elif 'OVP' in li[i]:
                    count_ovp+=1
                elif 'OTP' in li[i]:
                    count_otp+=1
            except IndexError:
                count_indexerror+=1
                pass    
            except ValueError:
                count_valueerror+=1
                pass
    return Tmp,Iout,Vout,count_ovp

#store into excel
def store_excel():

    for i,e in enumerate(Tmp):
        sheet1.write(i,2,e)
    for i,e in enumerate(Iout):
        sheet1.write(i,1,e)
    for i,e in enumerate(Vout):
        sheet1.write(i,0,e)
    name = "Vout Iout Tmp Infomation.xls"
    book.save(name)
    book.save(TemporaryFile())
    print("-------------------------")
    print('Your Excel File : ',name,'Has Been Saved')
    
def decide_pass():
    try:
        avr=sum(Tmp)/len(Tmp)
        avr1=sum(Iout)/len(Iout)
        avr2=sum(Vout)/len(Vout)
        print("-------------------------")        
        if avr > 85:
            print('Tmp over Height')
        elif avr1 <600:
            print('Iout Too Low')
        elif avr2 < 4500:
            print('Vout Too Low')
        else:
            print(pru_id ,'Pass')
        print("-------------------------")
        print('Total Data: ' +str(len(Iout)))
        print('Average Tmp : '  +str(avr))
        print('Average Iout : ' +str(avr1))
        print('Average Vout :' +str(avr2))
        print('OVP Times: ',count_ovp)
        print('OTP Times: ',count_otp)
        print('-------------------------')
        if count_indexerror != 0 or count_valueerror !=0:    
            print("Index Error : " +str(count_indexerror))
            print("Vallue Error :" +str(count_valueerror))
    except ZeroDivisionError:
        print("Your PRU ID is not correct ! ! ! ")
        avr=0
        avr1=0
        avr2=0
        print("-------------------------")
        print('Total Data: ' +str(len(Iout)))
        print('Average Tmp : '  +str(avr))
        print('Average Iout : ' +str(avr1))
        print('Average Vout :' +str(avr2))
    #return avr,avr1,avr2
logopen(pru_id)
store_excel()
decide_pass()
fin.close()

                        


