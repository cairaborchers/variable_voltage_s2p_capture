import pyvisa
import xlwings as xw
import pandas as pd
import numpy as np
from get_s7530_VNA_data_VISA import get_s7530_VNA_data_VISA 
from WF_SDK import device, static, supplies, error     # import instruments
from time import sleep 

#test commit2
vna = get_s7530_VNA_data_VISA()


#vna.setup_Visa()

try:
    # connect to the device
    device_data = device.open()

# start the positive supply
    supplies_data = supplies.data()
    supplies_data.master_state = True
    supplies_data.state = True
    supplies_data.positive_state = True
    
    i = .5
    while i <= 5:
        #Step from 0v on Vcc to 5v on Vcc
        supplies_data.positive_voltage = i
        supplies.switch(device_data, supplies_data)
        sleep(1)  # delay
        
        vna.single_Sweep(str(round(i, 1)))
        

        i = i + .1
 
    """-----------------------------------"""

    # close the connection
    device.close(device_data)

except error as e:
    print(e)
    # close the connection
    device.close(device.data)


vba_book = xw.Book("test.xlsb")
s2pImportMacro = vba_book.macro("main.s2p_import")
s2pImportMacro()


