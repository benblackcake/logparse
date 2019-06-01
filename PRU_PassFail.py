# -*- coding: utf-8 -*-
from PRU_Class import Pru
fin=open("teraterm.log",errors='ignore')
pru_number="F9:8B:64:8D:12:FF"
act=Pru(pru_number)
act.logopen(fin,pru_number)
act.decide_pass_Pru(pru_number)
act.decide_pass_Ptu(pru_number)
act.store_excel()
fin.close()
