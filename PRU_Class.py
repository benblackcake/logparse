# -*- coding: utf-8 -*-
import xlwt
from tempfile import TemporaryFile
book = xlwt.Workbook()
sheet_Pru = book.add_sheet('PRU')
sheet_Ptu = book.add_sheet('PTU')
class Pru:
    def __init__(self,pru_id):
        self.pru_id=pru_id
        self.Tmp=[]
        self.Iout=[]
        self.Vout=[]
        self.Vin=[]
        self.Iin=[]
        self.Icoil=[]
        #Counter
        self.count_times=0
        self.count_indexerror=0
        self.count_valueerror=0
        self.count_ovp=0
        self.count_otp=0
        #set up Pass condition
        self.Tmp_Set=85
        self.Iout_Set=600
        self.Vout_Set=4500
        self.Vin_set=11500
        self.Iin_Set=700
        self.Icoil_Set=400   
        
    def logopen(self,fin,pru_id):
        for line in fin:
            li=line.split("\t")
            for i in range(0,len(li)):
                li[i]=li[i].strip(" \n")
                try:
                    if pru_id in li[i]:
                        self.Tmp.append(int(li[6]))
                        self.Iout.append(int(li[3]))
                        self.Vout.append(int(li[2]))
                        self.count_times+=1
                        print("Loading... " +str(self.count_times), "record")
                    elif 'OVP' in li[i]:
                        self.count_ovp+=1
                    elif 'OTP' in li[i]:
                        self.count_otp+=1
                    if 'Ptrans' in li[i]:
                        self.Vin.append(int(li[6]))#Vin 
                        self.Iin.append(int(li[7]))#Iin
                        self.Icoil.append(int(li[11]))#Icoil
                except IndexError:
                    self.count_indexerror+=1
                    pass    
                except ValueError:
                    self.count_valueerror+=1
                    pass
                
    def decide_pass_Pru(self,pru_id):
        try:
            
            avr_Tmp=sum(self.Tmp)/len(self.Tmp)
            avr_Iout=sum(self.Iout)/len(self.Iout)
            avr_Vout=sum(self.Vout)/len(self.Vout)
            
            print("-------------------------")        
            if avr_Tmp > self.Tmp_Set:
                print('Tmp over Height')
            elif avr_Iout < self.Iout_Set:
                print('Iout Too Low')
            elif avr_Vout < self.Vout_Set:
                print('Vout Too Low')
            else:
                print(pru_id ,'PRU Pass')
            print("-------------------------")
            print('Total Data: ' +str(len(self.Iout)))
            print('Average Tmp : '  +str(avr_Tmp))
            print('Average Iout : ' +str(avr_Iout))
            print('Average Vout :' +str(avr_Vout))
            print('OVP Times: ',self.count_ovp)
            print('OTP Times: ',self.count_otp)
            print('-------------------------')
            if self.count_indexerror != 0 or self.count_valueerror !=0:    
                print("Index Error : " +str(self.count_indexerror))
                print("Vallue Error :" +str(self.count_valueerror))
        except ZeroDivisionError:
            print("Your PRU ID is not correct ! ! ! ")
            avr_Tmp=0
            avr_Iout=0
            avr_Vout=0
            print("-------------------------")
            print('Total Data: ' +str(len(self.Iout)))
            print('Average Tmp : '  +str(avr_Tmp))
            print('Average Iout : ' +str(avr_Iout))
            print('Average Vout :' +str(avr_Vout))
     
    def decide_pass_Ptu(self,pru_id):
        try:

            avr_Vin=sum(self.Vin)/len(self.Vin)
            avr_Iin=sum(self.Iin)/len(self.Iin)
            avr_Icoil=sum(self.Icoil)/len(self.Icoil)
            print("-------------------------")
            if avr_Vin < self.Vin_set:
                print("Vin To Low")
            elif avr_Iin < self.Iin_Set:
                print("Iin To Low")
            elif avr_Icoil < self.Icoil_Set:
                print("Icoil To Low")
            else:
                print("PTU Pass")
            print("-------------------------")
            print('Total Data: ' +str(len(self.Vin)))
            print('Average Vin : '  +str(avr_Vin))
            print('Average Iin : ' +str(avr_Iin))
            print('Average Icoil : ' +str(avr_Icoil))
        except ZeroDivisionError:
            avr_Vin=0
            avr_Iin=0
            avr_Icoil=0
            print("-------------------------")
            print('Total Data: ' +str(len(self.Vin)))
            print('Average Vin : '  +str(avr_Vin))
            print('Average Iin : ' +str(avr_Iin))
            print('Average Icoil : ' +str(avr_Icoil))
            
    def store_excel(self):
        #PRU
        for i,e in enumerate(self.Tmp):
            sheet_Pru.write(i,2,e)
        for i,e in enumerate(self.Iout):
            sheet_Pru.write(i,1,e)
        for i,e in enumerate(self.Vout):
            sheet_Pru.write(i,0,e)
        #PTU
        for i,e in enumerate(self.Vin):  
            sheet_Ptu.write(i,0,e)
        for i,e in enumerate(self.Iin):
            sheet_Ptu.write(i,1,e)
        for i,e in enumerate(self.Icoil):
            sheet_Ptu.write(i,2,e)
        name_PRU = "PRU and PTU Infomation_NE.xls"
        book.save(name_PRU)
        book.save(TemporaryFile())
        print("-------------------------")
        print('Your Excel File : ',name_PRU,'Has Been Saved') 

 