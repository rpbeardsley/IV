# Ryan Beardsley 17/10/2016
# Control a Keysight N6705B DC power analyser via USB


from tkinter import *
import visa
import numpy as np
import time



######################
#enter some measurement parameters
FlPth = r"FILE_PATH_HERE" #the full file path to save the data
MaxV = 1 #V, maximum voltage to try
MinV = 0 #V, minimum volatge to try
ResV = 0.1 #V, voltage step size
CurrLim = 0.3 #A, current limit
IntTim = 0.1 #seconds, integration time
######################

rm = visa.ResourceManager() # Open an instance of the resource manager and assign an object handle rm
#print(rm.list_resources()) #print a list of available resources
KeysightDCPA =  rm.open_resource('USB0::0x0957::0x0F07::MY53003369::INSTR') #open an instance of the USB resource class and assign a task handle

#get the instrument ID to check we are talking and set measurement protocols
SetIntTim = 'SENS:ELOG:PER ' + str(IntTim) + ', (@1)'
SetCurrLim = 'CURR:LIM ' + str(CurrLim) + ', (@1)'
 
InstrumentID = KeysightDCPA.query('*IDN?')
KeysightDCPA.timeout = None #set no timeout
KeysightDCPA.write('FORM:DATA ASCII') #set the data format
KeysightDCPA.write('CURR:LIM:COUP ON, (@1)') #set current limit tracking to on (- = + current limit)
KeysightDCPA.write(SetCurrLim) #set the current limit
KeysightDCPA.write(SetIntTim) #set the integration time

#crude way of measuring IV (this can probably be done by the instrument with the right commands)
Npts = int((MaxV-MinV)/ResV)
V = np.linspace(MinV, MaxV, Npts)
I = np.zeros(Npts)
a = 0 # counter for button pushes

KeysightDCPA.write('SOUR:VOLT:LEV:IMM:AMPL 0, (@1)') #set the voltage to 0 V to start
KeysightDCPA.write('OUTP ON, (@1)') #turn on the output


root = Tk()

def Measure(event):
    global a
    #apply the next voltage and record the current
    SetV = 'SOUR:VOLT:LEV:IMM:AMPL ' + str(V[a]) + ', (@1)'
    time.sleep((3*IntTim))
    KeysightDCPA.write(SetV) #set the voltage
    time.sleep(0.2) #pause for thought - N6705B can't think fast enough
    I[a] = KeysightDCPA.query('MEAS:CURR? (@1)') #measure the current
    #y[a] = KeysightDCPA.query('SENS:CURR:DC?, (@1)') #measure on sense
    time.sleep(0.2) #pause for thought - N6705B can't think fast enough
    print (str(I[a]) + ', ' + str(V[a]))
    a+=1 #add 1 to the counter a 
    
    
    if a >= len(V): #save the data and exit the function
        
        #write the data to a text file
        f = open(FlPth, "w")
        for b in range(0,Npts): 
           f.write(str(I[b]) + "   " + str(V[b]))       
           f.write("\n")
        f.close()
        

        KeysightDCPA.write('SOUR:VOLT:LEV:IMM:AMPL 0, (@1)') #set the voltage yo 0 V to finish
        KeysightDCPA.write('OUTP OFF, (@1)') #turn off the output
        # clear and close the instance of the USB instrument class, and close the resource manager class
        KeysightDCPA.clear()
        KeysightDCPA.close()
        rm.close()
        
        a = 0

button_1 = Button(root, text="Step the IV")
button_1.bind("<Button-1>", Measure)
button_1.pack()

root.mainloop()













































