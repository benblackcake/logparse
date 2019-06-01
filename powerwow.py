# -*- coding: utf-8 -*-
from PRU_Class import Pru
if __name__ == '__main__':
    print("---------------------------")
    print("***********Notice**********")
    print("Add *.log After your filename ")
    logfile=input("Please Input Your Logfile : ")
    pru_number=input("Please Input Your PRU ID : ")
    fin=open(logfile,errors='ignore')
    act=Pru(pru_number)
    act.logopen(fin,pru_number)
    act.decide_pass_Pru(pru_number)
    act.decide_pass_Ptu(pru_number)
    act.store_excel()
    fin.close()
