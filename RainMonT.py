#!c:\ProgramData\miniconda3\envs\py27\python.exe
# -*- coding: utf-8 -*-
"""
Created on Fri Mar 24 22:04:34 2017

@author: Kym
Notes for MK3 of hardware:
Previously to get the status (lets assume this is detector ) you sent *R3S<return>
and you rece3ived:
*BISDEE RAIN SENSOR  3 STATUS = D 012F
The last 4 digits being the detector temperature measured by a thermistor (A/D output)

For MK3 detectors, you send the same *R3S return
and you get back:
*BISDEE RAIN SENSOR MK3  3 STATUS = D 012F

I have added a command to get all the temperatures *R3A return
and you get back:
*BISDEE RAIN SENSOR MK3  3 Tamb=0126 Tnormal=08 Tdrying=0F (ascii hex)

You can also change the two offsets which I have called Tnormal &  Tdrying (defaults 08  &  15 respectively)
Send *R3NXY<return> to change the normal offset
     *R3DXY<return> to change the drying offset
     -  X - Hi Nibble of ascii hex data  &  Y - Lo nibble of ascii hex data
So if you send OF you set the offset to 15 degrees which will be written to the Eeprom.

If you send *R3I return you will restart ( Initialise) the program this could be useful if you get stuck with an error but dangeous.

Hope all this makes sense - any thing can be changed by request at this point.

@editor: Bryn

Have added TCP functionality so this can talk to whatever is needed.
Currently talks to the 50cm dome controller to automatically close the dome in
case of rain detection on 2 or more sensors (to prevent false positive)

Rainmon should stop sending a close signal to the dome shutter once
the dome has been closed automatically, until a response is received
that the user has re-opened the dome

"""


import sys
sys.path.insert(0,'C:/Python27/Lib')
sys.path.insert(0,'C:/Python27/Lib/site-packages/win32') 
import os
#sys.path.append(os.path.dirname(os.path.realpath(__file__)))
#sys.path.append(os.environ["USERPROFILE"]+'\Desktop\BT Telescope\Python\AutoGuider')
#sys.path.append('C:\Users\hill\Desktop\BT Telescope\Python\cfw-10')
#sys.path.append('C:\Documents and Settings\Kym\Desktop\BT Telescope\Python\cfw10')

#from Tkinter import *
import tkinter as Tk
from tkinter import ttk
import serial
import socket
import select
import logging
import logging.handlers
#import io
import time
import win32com.client
from pydub import AudioSegment
from pydub.playback import play
from queue import Queue
import threading
from math import log, exp
#import pyttsx3
#import pyttsx
#from pyttsx3.drivers import sapi5



C_GREEN="#00B000"
C_RED="#B00000"
C_WHITE="#FFFFFF"
C_BLACK="#000000"
C_GRAY="#CCCCCC"
Beta=4100.0
#Given Thermistor ADC value return temperature in deg C
def T_adc(adc):
    return -273.15+Beta*298.15/(Beta - 298.15*log((1024.0 - adc)/adc))

#Given t in Deg C return the ADC value    
def adc(t):
    k = (Beta - (Beta*298.15/(t+273.15)))/298.15
    return 1024.0/(1+exp(k))

#Extract the 3 temperatures from a MK3 detector and return them as a Tupple
def decodeMK3(s):
    i = s.find("Tamb=") + 5
    try:
        amb = (500*float(int(s[i:i+4],16))/1024.0) - 273.15
    except:
        amb = -999.
    i = s.find("Tnormal=") + 8
    try:
        norm = int(s[i:i+2],16)
    except:
        norm = -999
    i = s.find("Tdrying=") + 8
    try:
        dry =  int(s[i:i+2],16)
    except:
        dry = -999
    return (amb,norm,dry)
    
class RainWatch(Tk.Frame):
    STATUS_PREFIX ="*BISDEE RAIN SENSOR "
    
    def schedTlog(self):
        self.logTNow = True
        
    def schedWetAlert(self):
        self.wetAlert = True
        #print "sched Alert"
    def cancelWetAlert(self):
        schedJobs = self.tk.call('after', 'info') #Get all scheduled tasks
        if self.schedWetAlertID in schedJobs :   #Cancel wet sensor alert schedule if it exists
           self.after_cancel(self.schedWetAlertID) #cancel any scheduled wetAlerts
           self.schedWetAlertID = "x"
        
    def repeater(self):                          # on every N millisecs
        self.statusUpdate()
        self.wetSensorCount = self.currentStatus.count("w") + 2*self.currentStatus.count("W")
        
        if self.currentStatus != self.oldDetectorState:
           if self.writeLog:
               logging.info("".join(self.currentStatus)+" "+",".join(self.currentTemp))
           self.update_idletasks()
           if self.wetSensorCount>0:
               self.wetAlert = True
           else:
               self.wetAlert = False
               #and cancel scheduled Wet Alerts
               self.cancelWetAlert()
               try:
                   speaker.speak_async("Rain sensor Dry. ")
                   #engine.say("Rain sensor Dry. ")
                   #engine.runAndWait()
               except:
                   print("Audio failed - for Dry Sensor: ",time.asctime())
           self.oldDetectorState = self.currentStatus[:]

        # if wetAlert then give an audio alert with sensor status
        if self.wetAlert:
           # cancel any scheduled wet alerts
           self.cancelWetAlert()
           if self.wetSensorCount>0: #only give alert if we have wet sensors
              try:
                  speaker.speak_async("Rain Detected. "+"{:1d} of {:1d} sensors wet".
                                    format(self.wetSensorCount,self.activeSensorCount))
                  #engine.say("Rain Detected. "+"{:1d} of {:1d} sensors wet".
                  #                  format(self.wetSensorCount,self.activeSensorCount))
                  #engine.runAndWait()
              except:
                   print("Audio failed - for WET Sensor: ",time.asctime())
              #re-sched alert in 30s, while sensors are wet   
              self.schedWetAlertID = self.after(30000,self.schedWetAlert)  #re-sched alert in 30s
           self.wetAlert = False
           self.update_idletasks()
        
        if self.wetSensorCount > 0:  # allow for 1 sensor wet (might be false +)
            if not self.close_issued:
                self.checkWetAndClose()
            
        
        #Log the temperature periodically if required     
        if self.tLogging:
            if self.logTNow :
               logging.info("TempLog: "+",".join(self.currentTemp))   
               self.logTNow = False
               self.schedTempID = self.after(self.tLogCount*1000, self.schedTlog)

        self.repeaterSchedID=self.after(self.msecs, self.repeater)    # reschedule handler
    
    def checkWetAndClose(self):
        # logic for what to do under different wet conditions
        # if more than one sensor is wet or has recently dried, then close the dome immediately
        if self.wetSensorCount > 1 or self.wetSensorCount == 0:
            # send a message with the word 'close' which will close the dome
            if not self.close_issued and self.TCP_connected:
                self.TCP_send(f"Rain detected. {self.wetSensorCount:1d} sensors wet. Close command issued")
            # Dome button goes blue when closed due to rain
                if self.TCP_connected:
                    self.status_led.config(bg='Blue')
                self.timeoutDome(300000)  # times out any commands for 5 minutes
        if self.wetSensorCount == 1:
            if not self.close_issued:
                # only check if close hasn't already been issued
                self.after(10000, self.checkWetAndClose)
    
    def timeoutDome(self, duration):
        # sets close_issued to True so no more commands sent to dome controller
        self.close_issued = True
        self.after(duration, self.cancelDomeTimeout)        
        
    def cancelDomeTimeout(self):
        # sets close_issued to False so commands can be sent to dome again
        self.close_issued = False
        # Set dome button back to green 
        if self.TCP_connected:
            self.status_led.config(bg='Green')
        
    def __init__(self,detectors,writeLog,tLogCount,msecs=1000):              # default = 1 second
        Tk.Frame.__init__(self)
        self.writeLog = writeLog
        self.msecs = msecs
        self.close_issued = False
        if tLogCount>0:
            self.tLogging = True
            self.logTNow = True  #Force a Temperature logging to start in repeatrer
        else:
            self.tLogging = False
            self.logTNow = False 
        self.tLogCount = tLogCount
        self.schedTempID = ""     #holds 'after' id of the temperature logging scheduler
        self.schedWetAlertID = "x" #holds 'afer' id of wet alert scheduler
        self.activeSensorMap = ""
        self.initialSensorState = ""
        self.wetAlert = False
        self.TCP_connected = False
        
        self.STATUSVALUES = ["P","M","I","D","W","w","E","e"]
        self.statusColours = {"P":["white","black"],
                           "M":["black","gray"],
                           "I":["black","cyan"],
                           "D":["black","limegreen"],
                           "W":["white","blue"],
                           "w":["black","royalblue"],
                           "E":["white","red"],
                           "e":["white","red"]}
        
        self.pack(fill="both", expand=True)
        
        # Main container for layout
        main_frame = Tk.Frame(self)
        main_frame.pack(fill="both", expand=True)
        
        # Canvas with map image on the left
        self.canvas = Tk.Canvas(main_frame, width=268, height=432)
        self.canvas.grid(row=0, column=0, rowspan=3, sticky="nw")
        
        # Load map image
        try:
            self.map_photo = Tk.PhotoImage(file="./BT_SiteVectorMap.png")
            self.canvas.create_image(0, 0, anchor=Tk.NW, image=self.map_photo)
        except Exception as e:
            print(f"Error loading image: {e}")
        
        # Detector positions and GUI elements
        self.detector_positions = [(98, 293), (205, 258), (123, 95), (42, 398)]
        self.detectorsGUI = []
        self.currentStatus=["-","-","-","-"]
        self.currentTemp=["x","x","x","x"]
        
        # Initialize detector labels on canvas
        for i, pos in enumerate(self.detector_positions):
            detector_label = Tk.Label(self.canvas, width=1, padx=5,
                                      justify=Tk.CENTER,
                                      text="-",
                                      font=("CourierBold", 16),
                                      bg="white",
                                      relief=Tk.SUNKEN)
            self.canvas.create_window(pos[0], pos[1], window=detector_label)
            self.detectorsGUI.append(detector_label)
        
        # Right side frames
        right_frame = Tk.Frame(main_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=10)
        
        # Configure the main frame grid to ensure the right side stretches
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)
        
        # Frame for status indicators (one per detector) and IDs
        frameStatus = Tk.Frame(right_frame)
        frameStatus.grid(row=0, column=0, sticky="nsew")
        
        # Configure grid columns in frameStatus for equal width distribution
        for i in range(3):
            frameStatus.grid_columnconfigure(i, weight=1)
        
        # Header labels for ID, Status, and Temp
        header_id = Tk.Label(frameStatus, text="ID", font=("CourierBold", 11))
        header_id.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        
        header_status = Tk.Label(frameStatus, text="Status", font=("CourierBold", 11))
        header_status.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        
        header_temp = Tk.Label(frameStatus, text="Temp", font=("CourierBold", 11))
        header_temp.grid(row=0, column=2, padx=5, pady=5, sticky="nsew")
        
        self.detector_names = ['H127', 'H50', 'ACC', 'NYI']
        self.leds = []
        self.temps = []
        self.ID_labels = []
        
        # Row loop for ID, Status (LED), and Temp fields
        for i in range(4):
            # Add an ID label to the left of each row
            id_label = Tk.Label(frameStatus, text=f"{self.detector_names[i]}", font=("CourierBold", 9))
            id_label.grid(row=i + 1, column=0, padx=5, pady=5, sticky="nsew")
            self.ID_labels.append(id_label)
        
            # LED indicator - square appearance with consistent padding
            led = Tk.Label(frameStatus, width=2, text=" ", font=("CourierBold", 9), bg="red", relief=Tk.RAISED)
            led.grid(row=i + 1, column=1, padx=5, pady=5)
            self.leds.append(led)
        
            # Temperature indicator
            temp = Tk.Label(frameStatus, text="x", font=("CourierBold", 9), bg="white", relief=Tk.SUNKEN)
            temp.grid(row=i + 1, column=2, padx=5, pady=5, sticky="nsew")
            self.temps.append(temp)
        
        # Add the RESET button directly below the headers
        frame_reset_button = Tk.Frame(frameStatus)
        frame_reset_button.grid(row=6, column=0, columnspan=3, pady=(5, 5), sticky="nsew")
        
        # RESET button inside frameStatus
        self.reInitButton = Tk.Button(frame_reset_button, text="RESET", command=self.reInit)
        self.reInitButton.pack(side=Tk.TOP, pady=5)
        
        # Frame for Dome button and status at the bottom
        frameButtons = ttk.LabelFrame(right_frame, labelanchor="n", text="Dome Auto-Close", style="TLabelframe")
        frameButtons.grid(row=1, column=0, pady=(0, 0), sticky="new")
        
        # Configure grid rows in the frameButtons to organize the widgets
        frameButtons.grid_rowconfigure(0, weight=1)  # Connect/Disconnect buttons
        frameButtons.grid_rowconfigure(1, weight=1)  # Status label and LED indicator
        frameButtons.grid_rowconfigure(2, weight=1)  # Error label and Show Error button
        
        # Top row: Connect and Disconnect buttons
        self.connect_button = Tk.Button(frameButtons, text="CONNECT", command=self.dome_connect)
        self.connect_button.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")
        
        self.disconnect_button = Tk.Button(frameButtons, text="DISCONNECT", command=self.dome_disconnect)
        self.disconnect_button.grid(row=0, column=1, padx=5, pady=(0, 5), sticky="ew")
        
        # Second row: Status label and LED indicator (representing status)
        self.status_label = Tk.Label(frameButtons, text="Status", font=("CourierBold", 9))
        self.status_label.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        
        # LED indicator (green = connected, red = disconnected, blue = close command from wet issued)
        self.status_led = Tk.Label(frameButtons, text=" ", bg="#555555", width=2, height=1, relief=Tk.RAISED)
        self.status_led.grid(row=1, column=1, padx=5, pady=5)
        
        # Third row: Error label and Show Error button
        self.error_label = Tk.Label(frameButtons, text="Err?", font=("CourierBold", 9))
        self.error_label.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")
        
        # Show Error button (only enabled when there is an error)
        self.error_button = Tk.Button(frameButtons, text="View", command=self.show_error, state="normal")
        #self.error_button.grid(row=2, column=1, padx=5, pady=5)
        
        # Configure grid rows in the right frame to fill space properly
        right_frame.grid_rowconfigure(0, weight=3)  # Status indicators take up remaining space
        right_frame.grid_rowconfigure(1, weight=1)  # Place Dome control below the status


        self.update_idletasks()
        self.probeDetectors(detectors)
        self.oldDetectorState = ""
        self.activeSensorCount = self.activeDetectors.count(True)*2  #Two sensors per detector
        self.schedRepeaterID = ""
        # close TCP connections on window close
        self.master.protocol("WM_DELETE_WINDOW", self.close_connection)
        self.repeater() #Start monitoring

    def statusUpdate(self):
        for i in range(4):
            if self.activeDetectors[i]:
                self.leds[i].configure(bg="blue")
                self.update_idletasks()
                self.after(30)
                try:
                   port.write("*R{:1d}S\r".format(i).encode('utf-8'))
                except:
                   self.currentStatus[i] = "e"
                else:
                   status = readPort()
                   if "Timeout" in status:
                      self.currentStatus[i] = "e"
                   else:
                      #verify returned status *BRSxs
                      #mk3 = False
                      if "MK3" in status:  #Remove "MK3" from status
                          #print "status->:{}->".format(status),
                          if status[20:25] == "MK3  ":
                            status = status[:20] + status[25:]
                            #mk3 = True
                            #Request and log current temperature setpoints here (MK3 only)
                            try:
                              
                              port.write("*R{:1d}A\r".format(i).encode('utf-8'))
                            except:
                               #self.activeSensors.append(False)  #Disable sensors that dont respond
                               #self.currentStatus[i] = "E"
                               mk3status = "MK3 sensor {:1d} *RxA status port write error".format(i)
                            else:
                                mk3status = readPort()
                                amb,norm,dry = decodeMK3(mk3status)
                                mk3status = mk3status.split("=")
                                fmtstatus ="MK3 sensor {:1d}:{:s}:{:s}".format(i,status[22:-5],mk3status[0]) + \
                                           "={:4s}({:5.2f}C){:s}".format(mk3status[1][:4],amb,mk3status[1][4:]) + \
                                           "={:2s}({:2d}C){:s}".format(mk3status[2][:2],norm,mk3status[2][2:]) + \
                                           "={:2s}({:2d}C){:s}".format(mk3status[3][:2],dry,mk3status[3][2:])
                                if debug : print(fmtstatus)
                                #print "MK3 Detector {:1d}) status: {:s} -> (amb={:5.2f}C, norm={:2d}C, dry={:2d}C)".format(i,status,amb,norm,dry)

                      #print status
                      if len(status)==37:
                          adc=status[-4:]
                          status = status[:32]
                          #print adc,
                          try:
                              t=T_adc(int(adc,16))
                              tc="{:3.1f}".format(t)
                              #print "{:1d},{:3.1f}C".format(i,t)
                          except:
                              adc=""
                      else:
                          adc=""
                      if (len(status)==32) and (status[0:20]==RainWatch.STATUS_PREFIX): #Status looks good
                         if (status[20:21]=="{:1d}".format(i)) and (status[-1] in self.STATUSVALUES):
                            self.currentStatus[i] = status[-1]
                            if adc !="":
                                self.currentTemp[i]=tc
                         else:  #Got bad reply ?
                            self.currentStatus[i] = "e"
                      else:
                            self.currentStatus[i] = "e"
                cols=self.statusColours[self.currentStatus[i]]
                if self.currentStatus[i].upper()=="E":
                    self.leds[i].configure(bg="red")
                else:
                    self.leds[i].configure(bg="green")
                self.detectorsGUI[i].configure(fg=cols[0],bg=cols[1],text=self.currentStatus[i])
                self.temps[i].configure(text=self.currentTemp[i])

    def probeDetectors(self,detectors):
        #request a status from detectors 0-3,  EG[True,False,True,True] for detectors 0,2,3 active
        self.activeDetectors = []  #We build this list in probeDetectors, values are True or False
        for i in range(4):
            #Check detectors that are enabled, detectors=[True,True,False,True] for detectors 0,1,3
            #Only if a detector responds will it be marked active.
            if detectors[i]:     
                self.leds[i].configure(bg="blue") #UI is blue when detector being probed
                try:
                  port.write("*R{:1d}S\r".format(i).encode('utf-8'))
                except:
                   self.activeDetectors.append(False)  #Disable detectors that get an error
                   self.currentStatus[i] = "e"
                else:
                   status = readPort()
                   if "Timeout" in status:  #If detector doesnt respond then disable it
                      self.activeDetectors.append(False)
                      self.currentStatus[i] = "e"
                   else:
                      #For MK2 and MK3 get the temperature of the detector
                      #mk3 = False
                      if "MK3" in status:  #Remove "MK3" from status and set mk3 true
                          if status[20:25] == "MK3  ":
                            status = status[:20] + status[25:]
                            #mk3 = True
                            #Request and log current temperature setpoints here (MK3 only)
                            try:
                              port.write("*R{:1d}A\r".format(i).encode('utf-8'))
                            except:
                               #self.activeDetectors.append(False)  #Disable detectors that dont respond
                               #self.currentStatus[i] = "E"
                               mk3status = "MK3 detector {:1d} *RxA status port write error".format(i)
                               logging.info("{:s}".format(mk3status))
                            else:
                                mk3status = readPort()
                                if debug : print("MK3 detector {:1d} status: {:s}".format(i,mk3status))
                                if self.writeLog:
                                    amb,norm,dry = decodeMK3(mk3status)
                                    mk3status = mk3status.split("=")
                                    fmtstatus ="MK3 detector {:1d} status: {:s}".format(i,mk3status[0]) + \
                                               "={:4s}({:5.2f}C){:s}".format(mk3status[1][:4],amb,mk3status[1][4:]) + \
                                               "={:2s}({:2d}C){:s}".format(mk3status[2][:2],norm,mk3status[2][2:]) + \
                                               "={:2s}({:2d}C){:s}".format(mk3status[3][:2],dry,mk3status[3][2:])
                                    logging.info("{:s}".format(fmtstatus))
                      if len(status)==37:  #Strip Temperature from MK2/MK3 status
                          adc=status[-4:]
                          status = status[:32]  #Remove temperature from status
                          #print adc,
                          try:
                              t=T_adc(int(adc,16))
                              tc="{:3.1f}".format(t)
                              #print "{:1d},{:3.1f}C".format(i,t)
                          except:
                              adc=""
                      else:  #No ADC Temperatur MK1 Harware
                          adc=""
                      
                      #Now decode the status and verify it                     
                      if (len(status)==32) and (status[0:20]==RainWatch.STATUS_PREFIX): #Status looks good
                         if (status[20:21]=="{:1d}".format(i)) and (status[-1] in self.STATUSVALUES):
                            self.activeDetectors.append(True)
                            self.currentStatus[i] = status[-1]
                            if adc != "":
                                self.currentTemp[i] = tc
                         else:  #Got bad reply ? Mark detector active with an error
                            self.activeDetectors.append(True)
                            self.currentStatus[i] = "e"
                      else:
                            self.activeDetectors.append(False)
                            self.currentStatus[i] = "e"
            else: 
                self.activeDetectors.append(False) #For detectors not enables IE i not in detectors
            
            #update UI to reflect status
            if self.activeDetectors[i]:
                cols=self.statusColours[self.currentStatus[i]]
                self.leds[i].configure(bg="green")
                self.detectorsGUI[i].configure(fg=cols[0],bg=cols[1],text=self.currentStatus[i])
            else:
                cols=["black","red"]
                self.leds[i].configure(bg="red")
                self.detectorsGUI[i].configure(fg=cols[0],bg=cols[1])
                
        if self.writeLog:
            self.activeSensorMap = f"Active Sensor map: {str(self.activeDetectors)}"
            self.initialSensorState = (
                f"Initial sensor state: {''.join(self.currentStatus)} {','.join(self.currentTemp)}"
            )
            logging.info(self.activeSensorMap)
            logging.info(self.initialSensorState)
        return
        
    def park(self):
        port.write("*P\r".encode('utf-8'))

    def reInit(self):
        for i in range(4):
            port.write("*R{:1d}I\r".format(i).encode('utf-8'))
    
    def connect_TCP(self, controller_ip, controller_port):
        '''
        Adding in the ability to connect to the 50cm dome software and
        close the dome automatically when rain detected
        '''
        # creates TCP socket
        self.c_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # connects to the server
        try:
            self.c_socket.connect((controller_ip, controller_port))
            self.c_socket.setblocking(False)
            # make the dome button green if successful
            self.status_led.config(bg='Green')
            self.error_message = None
            self.error_button.grid_forget()
            self.TCP_connected = True
        except Exception as e:
            print(f"Connection Error: {e}")
            # make dome button red
            self.status_led.config(bg='Red')
            self.error_message = f"{e}"
            self.error_button.grid(row=2, column=1, padx=5, pady=5)
            self.TCP_connected = False
        
        # send some data to the server, checking if there is a response
        if self.TCP_connected:
            data = self.TCP_send('Connection from Rainmon.')
            if data:
                print(f"Server: {data}")
            # check if the TCP server is active occasionally
            self.master.after(1000, self.TCP_check)
    
    

    def TCP_send(self, message):
        # send a packet via tcp to connected device. Does nothing if no tcp active
        if self.TCP_connected:
            try:
                self.c_socket.sendall(message.encode('utf-8'))
                ready_to_read, ready_to_write, _ = select.select([self.c_socket],
                                                             [], [], 1)
                if ready_to_read:
                    response = self.c_socket.recv(1024)
                    return response.decode()
            except BlockingIOError:
                pass
            except Exception as e:
                print(f'Error sending TCP message: {e}')
            
    def TCP_check(self):
        # periodically check if TCP connection is still alive
        if self.TCP_connected:
            try:
                # Check if socket is writable (for send) and readable (for recv)
                readable, writable, _ = select.select([self.c_socket],
                                                      [self.c_socket],
                                                      [],
                                                      1)
                
                # Send a heartbeat if the socket is writable
                if writable:
                    self.c_socket.send(b'beep')
                
                # Check for a response if the socket is readable
                if readable:
                    response = self.c_socket.recv(1024).decode()
                    if response:
                        print(f'Host: {response}')
                        if "open" in response.lower():
                            self.cancelDomeTimeout()
                    else:
                        # No data means the connection is closed
                        raise ConnectionResetError
    
                # Schedule the next check
                self.error_message = None
                self.error_button.grid_forget()
                self.master.after(10000, self.TCP_check)
            
            except Exception as e:
                # Connection lost or another error occurred
                print(f"TCP connection lost: {e}")
                self.c_socket.close()
                self.c_socket = None
                self.TCP_connected = False
                self.status_led.config(bg='Red')
                self.error_message = f"{e}"
                self.error_button.grid(row=2, column=1, padx=5, pady=5)
    
    def show_error(self):
        Tk.messagebox.showerror("Error", self.error_message)

    def close_connection(self):
        # Close the socket when the window is closed
        try:
            if self.c_socket:
                self.c_socket.close()
                print("Socket closed.")
        except Exception as e:
            print(f"Error closing socket: {e}")
        finally:
            self.master.destroy()
            
    def dome_connect(self):
        '''
        Connects to the dome controller, if it is running on PlaneWave
        '''
        self.connect_TCP("127.0.0.1",1338)
        
    def dome_disconnect(self):
        '''
        Connects to the dome controller, if it is running on PlaneWave
        '''
        try:
            if self.c_socket:
                self.c_socket.close()
                print("Socket closed.")
                self.status_led.config(bg='#555555')
                self.c_socket = None
                self.TCP_connected = False
        except Exception as e:
            print(f"Error closing socket: {e}")

###############################################       
def readPort():
    global port
    c=""
    buffer=""    
    while c!=chr(13):
        c=port.read().decode('utf-8')
#        print "got:"+c
        if c=="": 
            return "Timeout"
        if c==chr(13):
            port.flushInput()
            #print "Returning>"+buffer+"<"
            return buffer
        buffer +=c
        
#if __name__ == '__main__': Alarm(msecs=3000).mainloop()

# TODO: Fix rolling log
class CustomTimedRotatingFileHandler(logging.handlers.TimedRotatingFileHandler):
    # Extend the TimedRotatingFileHandler classs in logging.handlers to put
    # necessary header info into the log.
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.active_sensor_map = ""
        self.initial_sensor_state = ""
    
    def setState(self, active_sensor_map, initial_sensor_state):
        self.active_sensor_map = active_sensor_map
        self.initial_sensor_state = ""
    
    def doRollover(self):
        # Perform the original rollover process
        if self.stream:
            self.stream.close()
            self.stream = None
        
        super().doRollover()
        
        # Write header information to the new log file
        if self.stream:  # Ensure the new log file is open
            self.stream.write("----- Logging Started -----\n")
            if self.active_sensor_map:
                self.stream.write(self.active_sensor_map + "\n")
            if self.initial_sensor_state:
                self.stream.write(self.initial_sensor_state + "\n")
            self.stream.flush()

#############  Main #################################


#print os.getcwd()
#print os.path.dirname(os.path.realpath(__file__))

#change working directory to where the main program lives
#print os.path.dirname(os.path.realpath(__file__))
os.chdir(os.path.dirname(os.path.realpath(__file__)))

#process settings from the RainMon.ini file, if it exists
portname = 'COM31'
writeLog="0"
abort=0
debug = False
detectors="0123"  #Active detectors, all by default, override with active= in rainmon.ini
tLogCount = 0 #Temperature logging off
activeDetectors=[]
for i in "0123":
    if i in detectors:
        activeDetectors.append(True)
    else:
        activeDetectors.append(False)

#engine = pyttsx3.init()
#engine.setProperty('rate',140)
#engine = pyttsx.init()
#engine.setProperty('rate',140)

ttk.Style().theme_use('winnative')
ttk.Style().configure('TLabelframe')
ttk.Style().configure('TLabelframe.Label',font=("TkDefaultFont", 9, 'bold'))
ttk.Style().configure('TLabel')

#TODO: Fix speaking
class Speaker:
    def __init__(self):
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.speaker.Rate = -3
        self.queue = Queue()
        self.thread = threading.Thread(target=self._process_queue, daemon=True)
        self.thread.start()

    def _process_queue(self):
        while True:
            item = self.queue.get()
            if item is None:  # Stop signal
                break
            try:
                if isinstance(item, str):  # Text input
                    self.speaker.Speak(item)
                elif isinstance(item, dict) and item.get("type") == "audio":  # Audio input
                    file_path = item.get("file_path")
                    self._play_audio(file_path)
                else:
                    print(f"Invalid input to Speaker: {item}")
            except Exception as e:
                print(f"Error processing queue: {e}")

    def _play_audio(self, file_path):
        try:
            audio = AudioSegment.from_file(file_path)
            play(audio)
        except Exception as e:
            print(f"Error playing audio file {file_path}: {e}")

    def speak_async(self, text):
        #queue a text message for text-to-speech.
        self.queue.put(text)

    def play_audio_async(self, file_path):
        # queue an audio file to play.
        self.queue.put({"type": "audio", "file_path": file_path})

    def shutdown(self):
        #Shutdown the speaker system
        self.queue.put(None)
        self.thread.join()

try:             
    f=open("RainMon.ini","r")
    ini=f.read()
    f.close()
    for i in ini.split("\n"):
        if ("=" in i) and (i[0]!="#"): #Ignore comments in the file
            i=i.split("=")
            if i[0].upper()=="COM":
                portname=i[1]
            elif i[0].upper()=="ACTIVE":
                detectors=i[1]
            elif  i[0].upper()=="LOG":
                writeLog = i[1]
            elif  i[0].upper()=="TLOG":
                tLogCount = int(i[1])
            elif i[0].upper() =="DEBUG":
                if i[1].upper() == "TRUE":
                    debug = True 
                else:
                    debug = False
except:
    print("RainMon.ini error")
    #print sys.exc_info()
if (writeLog == "1") or (writeLog.upper() == "TRUE"):
    writeLog = True
    logName = "RainMonT-{}.log".format(time.strftime("%Y%m%d_%H%M"))
    logName="RainMonT.log"  #The current log has no date or extension.
else:
    writeLog=False

if writeLog:
#    logging.basicConfig(filename=logName,
#                        format='%(asctime)s %(message)s',
#                        datefmt='%d/%m/%Y %H:%M:%S',
#                        level=logging.DEBUG)
    loggerInstance = logging.getLogger()
    loggerInstance.setLevel(logging.DEBUG)
    handler = CustomTimedRotatingFileHandler(
                            logName,
                            when="W6",
                            interval=1,
                            backupCount=0,
                            delay=True
                            )
    handler.prefix = '%Y%m%d_%H%M'
    #Specify the required format                                               
    formatter = logging.Formatter('%(asctime)s %(message)s',datefmt='%d/%m/%Y %H:%M:%S')
    #Add formatter to handler
    handler.setFormatter(formatter)
    #remove any existing handlers including stream handler -> console output
    for h in loggerInstance.handlers[:]:
        h.close()
        loggerInstance.removeHandler(h)
    #Initialize logger instance with handler
    loggerInstance.addHandler(handler)
   
#open the serial port to the xy stage
try:
    port = serial.Serial(portname, 9600,timeout=0.9,rtscts=False,dsrdtr=False)
except: 
   top= Tk.Tk()
   top.withdraw()
   Tk.messagebox.showinfo("RainMon Starup Error", "Couldn't open COM port '{}'.\n Check RainMon.ini".format(portname))
   abort=1  #Quit after the error message has been acknowledged
   top.destroy()
   
if not abort:  
    #Set Window geometry and Title
    port.flushInput()
    port.writeTimeout=0.4
    speaker = Speaker()
    myapp=RainWatch(activeDetectors,writeLog,tLogCount,msecs=1000)
    handler.setState(myapp.activeDetectors, myapp.currentStatus)
    logging.info("----- Logging Started -----")
    logging.info("Active Sensor map: " + str(myapp.activeDetectors))
    logging.info("Initial sensor state: " + "".join(myapp.currentStatus) + " " + ",".join(myapp.currentTemp))
    myapp.master.title("RainWatch v0.3")
    myapp.configure(background=C_GRAY)
    #myapp.master.maxsize(200, 500)
    #myapp.master.minsize(200,100)
    #myapp.master.geometry('160x60')
    # start the program
    myapp.focus_set()
    myapp.mainloop()
port.close()
print("Finished.")
