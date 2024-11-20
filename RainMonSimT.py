#!c:\ProgramData\miniconda3\envs\py27\python.exe
# -*- coding: utf-8 -*-
"""
Created on Mon Apr 29 23:15:40 2019

@author: Kym
"""

import os
#sys.path.append(os.path.dirname(os.path.realpath(__file__)))
#sys.path.append(os.environ["USERPROFILE"]+'\Desktop\BT Telescope\Python\AutoGuider')
#sys.path.append('C:\Users\hill\Desktop\BT Telescope\Python\cfw-10')
#sys.path.append('C:\Documents and Settings\Kym\Desktop\BT Telescope\Python\cfw10')

from tkinter import *
import serial
import tkinter.messagebox as tkMessageBox
import time
import random
from math import *

#import io
C_GREEN="#00B000"
C_RED="#B00000"
C_WHITE="#FFFFFF"
C_BLACK="#000000"
"""
' Ascii input is of the following 5 character format
' *RNX carriage Return
' * - lead in character
' R - rain sensor
' N = Unit number - typically they may be up to 4 units
'                   Unit number is set by links on main PCB
' x =  S  - Status request - returns P, M, I, D, W, w, E.
                            '="P"  Parked - upside down
                            '="M"  Moving
                            '="I" Initialising only at prog start
                            '="D" Sensor Dry
                            '="W" Sensor completely Wet
                            '="w" Sensor 1/2 Wet
                            '="E" Error all stopped needs manual help
'      P  - Parks the sensor (rotate 180 degrees)
'      I  - Restarts Program - Re Initialise if you like


The status request returned string is of the following format:
hserout ["*BISDEE RAIN SENSOR ", ThisUnit," STATUS = ", Current, $0D]
ThisUnit is the sensor number 0,1,2,3
Current is the status single letter as above. 
"""
validCommands = ['R','P','I']
validStatus = ["P","M","I","D","W","w","E"]
validSensor = {"0":0,"1":1,"2":2,"3":3}
currentStates = ["P","P","P","P"]
currentTimes = [-10.0,-10.0,-10.0,-10.0]
norTOffset = ["08","08","08","08"]
wetTOffset = ["0F","0F","0F","0F"]
parkTime = 8.0
initTime = 20.0
moveTime = 8.0
Beta=4100.0
time.clock = time.perf_counter
t=time.clock()

#Resistance of thermistor
def R(T):
    return 10000.0/exp( Beta*(T+273.15-298.15)/(298.15*(T+273.15)) )   
#Temperature of rain sensor
def Temperature(t):
     return (30.-15.)*(1.0-exp(-t/40))+15.0       

def adcT(rt):
    return int(1024*rt/(rt+10000.0))
def myreadline(port):
    c=""
    buffer=""    
    while c!=chr(13):
        c=port.read().decode('utf-8')
#        print "got:"+c
        if c!=chr(13):
            buffer+=c
    print("Read:>"+buffer+"<")
    return buffer

portname = 'COM30'
abort=0
#print os.getcwd()
#print os.path.dirname(os.path.realpath(__file__))

#change working directory to where the main program lives
print(os.path.dirname(os.path.realpath(__file__)))
os.chdir(os.path.dirname(os.path.realpath(__file__)))

#open the serial port 
try:
    port = serial.Serial(portname, 9600)
except: 
   top= Tk()
   top.withdraw()
   tkMessageBox.showinfo("Petal Starup Error", "Couldn't open COM port '{}'.\n Check petals.ini".format(portname))
   abort=1  #Quit after the error message has been acknowledged
   top.destroy()
status="C"
movingTo="C"
mtime=0.0
cmd=""
#port.write("Petal Sim running\r\n")
port.flushInput()
port.flushOutput()

if not abort:   
    print("running")
    while cmd!="Q":
        x=myreadline(port)
        print("                Got command:"+x)
        if len(x) == 6: #x->*R1Dxy
            commandOK = (x[0]=="*") and (x[1]=="R") and (x[3] in "ND") 
            print("Len 6", commandOK)
        elif len(x)==4: # x->*R1S
            commandOK = (x[0]=="*") and (x[1]=="R") and (x[2] in validSensor.keys()) and (x[3] in "SAI")
        elif len(x) == 3:  #x->*I3 EG reset detector 3
            commandOK = (x[0]=="*") and (x[1]=="I") and (x[2] in validSensor.keys())
        elif len(x)==2:        
            commandOK = (x=="*P")

        if commandOK:
           if x[1]=="P":   #Park
              for i in validSensor.values():
                 if currentStates[i]=="P": #No delay if already parked
                    currentTimes[i]=time.clock()
                 else:
                    currentTimes[i]=time.clock()+parkTime+2.0*random.random()-1.0
                    currentStates[i]='MP'
           elif x[1]=="I" :  #Re-Init
              #for i in validSensor.values():
              i = validSensor[x[2]]
              if currentStates[i][0] not in "MI": #Can only re-init if not alread doing so
                 if currentStates[i][0] == 'P':
                     currentTimes[i]=time.clock()+initTime+8.0*random.random()-4.0
                     currentStates[i]='IM'
                 else:
                     currentTimes[i]=time.clock()+moveTime+2.0*random.random()-1.0
                     currentStates[i]='MIM'
           elif x[1]=="R":   #Status
              if x[3]=="S":
                  ss=validSensor[x[2]]  #convert string to index using dict
                  if time.clock() < currentTimes[ss]: #Something Scheduled
                     state=currentStates[ss][0]
                  else:
                     #print ss,currentStates
                     #print ss,currentTimes
                     if len(currentStates[ss])>1:  #Schedule complete
                         currentStates[ss] = currentStates[ss][1:]
                         if currentStates[ss][0] == 'I':
                             currentTimes[ss] = time.clock()+initTime+8.0*random.random()-4.0
                         elif currentStates[ss][0] == 'M':
                             currentTimes[ss]=time.clock()+moveTime+2.0*random.random()-1.0
                     else: #reached final state
                         if currentStates[ss]=="M":
                             currentStates[ss]="D"
                         elif currentStates[ss]=="P":
                             currentStates[ss]="P"
                     state=currentStates[ss][0]
                  
                  #Random state to DWw
                  if currentStates[ss][0]=="D": # 2% chance of w
                     if random.randint(1,101)>=98:
                         currentStates[ss]="w"
                  if currentStates[ss][0]=="w": #if W then 20% chance of W
                     if random.randint(1,11)>8:
                         currentStates[ss]="W"
                  if currentStates[ss][0] in "wW": #if W or w 10% chance of D
                     if random.randint(1,21)>18:
                         currentStates[ss]="D"
                  Temp = "{:04X}".format(adcT(R(Temperature(time.clock()))))
                  port.write("*BISDEE RAIN SENSOR MK3  {:1s} STATUS = {:1s} {:4s}\r".format(x[2],currentStates[ss][0],Temp).encode('utf-8'))
                  print("*BISDEE RAIN SENSOR MK3  {:1s} STATUS = {:1s} {:4s}\r".format(x[2],currentStates[ss][0],Temp))
                  #port.write("*BRS{:1s}{:1s}\r".format(x[2],currentStates[ss][0]))                  
                  #print "*BRS{:1s}{:1s}\r".format(x[2],currentStates[ss][0])
              elif x[3]=="I": #Re-Init
                  #for i in validSensor.values():
                  i = validSensor[x[2]]
                  if currentStates[i][0] not in "MI": #Can only re-init if not alread doing so
                     if currentStates[i][0] == 'P':
                         currentTimes[i]=time.clock()+initTime+8.0*random.random()-4.0
                         currentStates[i]='IM'
                     else:
                         currentTimes[i]=time.clock()+moveTime+2.0*random.random()-1.0
                         currentStates[i]='MIM'
              elif x[3]=="A":
                  i = validSensor[x[2]]
                  port.write("*BISDEE RAIN SENSOR MK3  {:1s} Tamb=0256 Tnormal={:2s} Tdrying={:2s}\r".format(x[2],norTOffset[i],wetTOffset[i]).encode('utf-8'))
              elif x[3]=="D": #New drying offset
                  print("X[3]==D")
                  i = validSensor[x[2]]
                  wetTOffset[i] = x[4:6]
                  print(wetTOffset)
                  port.write("DONE\r".encode('utf-8'))
              elif x[3]=="N": #New normal offset
                  i = validSensor[x[2]]
                  norTOffset[i] = x[4:6]
                  print(norTOffset)
                  port.write("DONE\r".encode('utf-8'))
                  
    #Set Window geometry and Title
port.close()
print("Finished.")
