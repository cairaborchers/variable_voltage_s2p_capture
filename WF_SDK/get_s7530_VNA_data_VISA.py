import pyvisa
from time import sleep 

class get_s7530_VNA_data_VISA:

    s7530 = 0

    def setup_Visa():
    #use ni visa 64 bit as backend
        rm = pyvisa.ResourceManager('C:/WINDOWS/system32/visa64.dll')

        #connect to the socket server
        try:
            s7530 = rm.open_resource('TCPIP0::127.0.0.1::5025::SOCKET')
        except:
            print("Failure to connect to VNA!")
            print("Check network settings")

        #the VNA ends each line with \n
        s7530.read_termination='\n'

        s7530.write("*ESR?")
        esr_response = s7530.read()

        #Get cylce time
        s7530.write("SYSTem:CYCLe:TIME:MEASurement?")
        cycle_time = s7530.read()

        #change to computer triggering
        s7530.write("TRIG:SOUR BUS")

        #end with an *OPC? to make sure the setups are complete (blocks)
        s7530.write("*OPC?")
        opc_response = s7530.read()



    def single_Sweep():
        
        #trigger a single sweep
        s7530.write("TRIG:SING")

        while esr_response == 0:
            
            sleep(10)
            s7530.write("*ESR?")
            esr_response = s7530.read()

        s7530.write('MMEM:STOR:SNP:DATA C:\\s2p_save_remote\\v' + str(round(i, 1)) + '.s2p')
        
