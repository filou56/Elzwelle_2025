import configparser
import os
import platform
import time
import traceback
import uuid
import paho.mqtt.client as paho
import tkinter
import functools

from   tkinter import ttk
from   tkinter import messagebox
from   paho    import mqtt
from   tksheet import Sheet
from OpenGL.GL.APPLE import row_bytes

#---------------------- Fix local.DE -------------------
class locale:
    
    @staticmethod
    def atof(s):
        return float(s.strip().replace(',','.'))

    @staticmethod
    def format_string(fmt, *args):
        return  (fmt % args).replace('.',',')
    
#-------------------------------------------------------------------
# Define the GUI
#-------------------------------------------------------------------
class sheetapp_tk(tkinter.Tk):
    
    def __init__(self,parent):
        tkinter.Tk.__init__(self,parent)
        self.parent = parent
        self.initialize()
        self.run  =  0
        self.xRow = -1
        self.xCol = -1
        self.xVal = ''

    def showError(self, *args):
        err = traceback.format_exception(*args)
        messagebox.showerror('Exception',err)
        
        # but this works too
        tkinter.Tk.report_callback_exception = self.showError

    def initialize(self):
        self.hfs = config.getint('view','header_font_size')
        self.cfs = config.getint('view','cell_font_size')
        self.cr  = config.getint('view','competition_rows')
        self.rr  = config.getint('view','result_rows')
        
        noteStyle = ttk.Style()
        noteStyle.theme_use('default')
        noteStyle.configure("TNotebook", background='lightgray')
        noteStyle.configure("TNotebook.Tab", background='#eeeeee')
        noteStyle.map("TNotebook.Tab", background=[("selected", '#005fd7')],foreground=[("selected", 'white')])
        
        self.geometry("1200x800")
        
        self.resultHeader = tkinter.Label(self,text="Wertung",
                                        font=("Arial", self.hfs),
                                        bg='#D3E3FD')
        self.resultHeader.pack(fill ="x") 
         
        #----- Start Page -------
                 
        self.resultSheet = Sheet(self,
                               name = 'resultSheet',
                               #data = [['00:00:00','0,00','',''] for r in range(2)],
                               header = ['Rang','Startnummer','Name','ZS Start','ZS Ziel','Fahrzeit','Strafzeit','Gesamtzeit'],
                               header_bg = "azure",
                               header_fg = "black",
                               header_font = ("Calibri", self.cfs, "normal"),
                               index_bg  = "azure",
                               index_fg  = "gray",
                               font = ("Calibri", self.cfs, "bold"),
                               data = [['','','','0,00','0,00','0,00','0','9999,00'] for i in range(self.rr)],
                               auto_resize_columns = 200,
                               auto_resize_rows = 30,
                            )
        self.resultSheet.enable_bindings()
        # self.resultSheet.grid(column = 0, row = 0)
        # self.resultSheet.grid(row = 0, column = 0, sticky = "nswe")
        self.resultSheet.span('A:').align('right')
        self.resultSheet.span('A').readonly()
        self.resultSheet.span('B').readonly()
        
        self.resultSheet.disable_bindings("All")
        # self.resultSheet.enable_bindings("edit_cell","single_select","drag_select","row_select","copy")
        # self.resultSheet.extra_bindings("end_edit_cell", func=self.startEndEditCell)

        self.resultSheet.pack(fill = "x") 
        
        self.competitionHeader = tkinter.Label(self,text="Wettbewerb",
                                        font=("Arial", self.hfs),
                                        bg='#D3E3FD')
        self.competitionHeader.pack(fill ="x") 
        
        self.competitionSheet = Sheet(self,
                               name = 'competitionSheet',
                               #data = [['00:00:00','0,00','',''] for r in range(2)],
                               header = ['','Startnummer','Name','ZS Start','ZS Ziel','Fahrzeit','Strafzeit','Gesamtzeit'],
                               header_bg = "azure",
                               header_fg = "black",
                               header_font = ("Calibri", self.cfs, "normal"),
                               index_bg  = "azure",
                               index_fg  = "gray",
                               font = ("Calibri", self.cfs, "bold"),
                               data = [['','','','0,00','0,00','0,00','0','0,00'] for i in range(self.cr)],
                               auto_resize_columns = 200,
                               auto_resize_rows = 30,
                            )
        self.competitionSheet.enable_bindings()
        # self.resultSheet.grid(column = 0, row = 0)
        # self.resultSheet.grid(row = 0, column = 0, sticky = "nswe")
        self.competitionSheet.span('A:').align('right')
        self.competitionSheet.span('A').readonly()
        self.competitionSheet.span('B').readonly()
        
        self.competitionSheet.disable_bindings("All")
        # self.resultSheet.enable_bindings("edit_cell","single_select","drag_select","row_select","copy")
        # self.resultSheet.extra_bindings("end_edit_cell", func=self.startEndEditCell)
        
        # self.competitionSheet.set_column_widths([100,200,200,300,300,300,300])
    
        self.competitionSheet.pack(expand = True, fill = "both")
    
    def getCompetitionRows(self):
        return self.cr
    
    def getResultRows(self):
        return self.rr
#-------------------------------------------------------------------

# setting callbacks for different events to see if it works, print the message etc.
def on_connect(client, userdata, flags, rc, properties=None):
    """
        Prints the result of the connection with a reasoncode to stdout ( used as callback for connect )

        :param client: the client itself
        :param userdata: userdata is set when initiating the client, here it is userdata=None
        :param flags: these are response flags sent by the broker
        :param rc: stands for reasonCode, which is a code for the connection result
        :param properties: can be used in MQTTv5, but is optional
    """
    print("CONNACK received with code %s." % rc)
    
    # subscribe to all topics of encyclopedia by using the wildcard "#"
    client.subscribe("elzwelle/stopwatch/#", qos=1)
    
    # a single publish, this can also be done in loops, etc.
    client.publish("elzwelle/monitor", payload="running", qos=1)
    

FIRST_RECONNECT_DELAY   = 1
RECONNECT_RATE          = 2
MAX_RECONNECT_COUNT     = 12
MAX_RECONNECT_DELAY     = 60

def on_disconnect(client, userdata, rc):
    print("Disconnected with result code: %s", rc)
    reconnect_count, reconnect_delay = 0, FIRST_RECONNECT_DELAY
    while reconnect_count < MAX_RECONNECT_COUNT:
        print("Reconnecting in %d seconds...", reconnect_delay)
        time.sleep(reconnect_delay)

        try:
            client.reconnect()
            print("Reconnected successfully!")
            return
        except Exception as err:
            print("%s. Reconnect failed. Retrying...", err)

        reconnect_delay *= RECONNECT_RATE
        reconnect_delay = min(reconnect_delay, MAX_RECONNECT_DELAY)
        reconnect_count += 1
    print("Reconnect failed after %s attempts. Exiting...", reconnect_count)

# with this callback you can see if your publish was successful
def on_publish(client, userdata, mid, properties=None):
    """
        Prints mid to stdout to reassure a successful publish ( used as callback for publish )

        :param client: the client itself
        :param userdata: userdata is set when initiating the client, here it is userdata=None
        :param mid: variable returned from the corresponding publish() call, to allow outgoing messages to be tracked
        :param properties: can be used in MQTTv5, but is optional
    """
    print("Publish mid: " + str(mid))


# print which topic was subscribed to
def on_subscribe(client, userdata, mid, granted_qos, properties=None):
    """
        Prints a reassurance for successfully subscribing

        :param client: the client itself
        :param userdata: userdata is set when initiating the client, here it is userdata=None
        :param mid: variable returned from the corresponding publish() call, to allow outgoing messages to be tracked
        :param granted_qos: this is the qos that you declare when subscribing, use the same one for publishing
        :param properties: can be used in MQTTv5, but is optional
    """
    print("Subscribed: " + str(mid) + " " + str(granted_qos))

# print message, useful for checking if it was successful
def on_message(client, userdata, msg):
    """
        Prints a mqtt message to stdout ( used as callback for subscribe )

        :param client: the client itself
        :param userdata: userdata is set when initiating the client, here it is userdata=None
        :param msg: the message with topic and payload
    """
    print(msg.topic + " " + str(msg.qos) + " " + str(msg.payload))
    
    payload = msg.payload.decode('ISO8859-1')        # ('utf-8')
                
    if msg.topic == 'elzwelle/stopwatch/start/number':
        try:
            data    = payload.split(' ')
            time    = data[0].strip()
            stamp   = data[1].strip()
            num     = int(data[2].strip())
            nums    = app.competitionSheet.span("B").data
    
            row     = findFreeRow(nums)
            if row >= 0:
                app.competitionSheet.insert_row(['',num,'',stamp,
                                                 ],row)   
                app.competitionSheet[row].highlight(bg='khaki')
                app.competitionSheet.del_row(app.getCompetitionRows()) 
            else:
                app.competitionSheet.del_row(0) 
                app.competitionSheet.insert_row(['',num,'',stamp,
                                                 ]) 
                app.competitionSheet[app.cr-1].highlight(bg='khaki')
            
        except Exception as e:
            print("MQTT Decode exception: ",e,msg.payload)
            
    if msg.topic == 'elzwelle/stopwatch/finish/number':
        try:
            data = payload.split(' ')
            time    = data[0].strip()
            stamp   = data[1].strip()
            num     = int(data[2].strip())
            nums    = app.competitionSheet.span("B").data
            
            row = findRow(nums,num)
            if row >= 0:
                app.competitionSheet.set_cell_data(row, 4, stamp)
            
        except Exception as e:
            print("MQTT Decode exception: ",e,msg.payload)
    
    if msg.topic == 'elzwelle/stopwatch/penalty/update':
        try:
            data    = payload.split(' ')
            num     = int(data[0].strip())
            penalty = int(data[1].strip())
            nums    = app.competitionSheet.span("B").data
            
            row = findRow(nums,num)
            if row >= 0:
                app.competitionSheet.set_cell_data(row, 6, penalty)
                
        except Exception as e:
            print("MQTT Decode exception: ",e,msg.payload)
        
    if msg.topic == 'elzwelle/stopwatch/result/update':
        try:
            data    = payload.split(' ')
            num     = int(data[0].strip())
            trip    = data[1].strip()
            penalty = data[2].strip()
            final   = data[3].strip()
            nums    = app.competitionSheet.span("B").data
            
            row = findRow(nums,num)
            if row >= 0:
                app.competitionSheet.set_cell_data(row, 5, trip)
                app.competitionSheet.set_cell_data(row, 6, penalty)
                app.competitionSheet.set_cell_data(row, 7, final)
                app.competitionSheet[row].highlight(bg='#D3E3FD')
                updateResultSheet(app.competitionSheet[row].data)
                
        except Exception as e:
            print("MQTT Decode exception: ",e,msg.payload)
        
def findFreeRow(numlist):
    for i in range(len(numlist)):
        if numlist[i] == '':
            return i
    return -1

def findRow(numlist,num):
    for i in range(len(numlist)):
        if numlist[i] != '':
            if int(numlist[i]) == num:
                return i
    return -1

def compResults(r1,r2):
    #print("Comp: ",r1,r2)
    f1 = locale.atof(r1[7])
    f2 = locale.atof(r2[7])
    
    if f1 == f2: return  0
    if f1 <  f2: return -1
    if f1 >  f2: return  1
    return 0

def updateResultSheet(row):
    print("Update result sheet")
    
    results = app.resultSheet.data + [row]
#   sortedResults = sorted(results, key=lambda item: item[7])
    sortedResults = sorted(results, key=functools.cmp_to_key(compResults))
    app.resultSheet.data = sortedResults[:len(sortedResults)-1]
    for row in range(app.rr):
        if app.resultSheet[row,7].data != '9999,00':
            app.resultSheet.set_cell_data(row,0,row+1)
            app.resultSheet[row].highlight('aquamarine')
    
    
#-------------------------------------------------------------------
# Main program
#-------------------------------------------------------------------

if __name__ == '__main__':    
   
    myPlatform = platform.system()
    print("OS in my system : ", myPlatform)
    myArch = platform.machine()
    print("ARCH in my system : ", myArch)

    config = configparser.ConfigParser()
   
    config['mqtt']   = { 
        'url':'144db7091e4a45cbb0e14506aeed779a.s2.eu.hivemq.cloud',
        'port':8883,
        'tls_enabled':'yes',
        'auth_enabled':'no',
        'user':'welle',
        'password':'elzwelle', 
    }
      
    config['mqtt']   = { 
         'header_font_size':22,
         'cell_font_size':16,
    }
    # Platform specific
    if myPlatform == 'Windows':
        # Platform defaults
        config.read('windows.ini') 
    if myPlatform == 'Linux':
        config.read('linux.ini')

    #--------------------------------- MQTT --------------------------

    # using MQTT version 5 here, for 3.1.1: MQTTv311, 3.1: MQTTv31
    # userdata is user defined data of any type, updated by user_data_set()
    # client_id is the given name of the client
    try:
        mqtt_client = paho.Client(client_id="elzwelle_"+str(uuid.uuid4()), userdata=None, protocol=paho.MQTTv311)
    
        # enable TLS for secure connection
        if config.getboolean('mqtt','tls_enabled'):
            mqtt_client.tls_set(tls_version=mqtt.client.ssl.PROTOCOL_TLS)
        # set username and password
        if config.getboolean('mqtt','auth_enabled'):
            mqtt_client.username_pw_set(config.get('mqtt','user'),
                                    config.get('mqtt','password'))
        # connect to HiveMQ Cloud on port 8883 (default for MQTT)
        mqtt_client.connect(config.get('mqtt','url'), config.getint('mqtt','port'))
    
        # setting callbacks, use separate functions like above for better visibility
        mqtt_client.on_connect      = on_connect
        mqtt_client.on_subscribe    = on_subscribe
        mqtt_client.on_message      = on_message
        mqtt_client.on_publish      = on_publish
        
        mqtt_client.loop_start()
    except Exception as e:
        messagebox.showerror(title="Fehler", message="Keine Verbindung zum MQTT Server!")
        print("Error: ",e)
        exit(1)   
        
    # ---------- setup and start GUI --------------
    app = sheetapp_tk(None)
        
    app.title("MQTT Display Tabelle Elz-Zeit")
    
    # run
    app.mainloop()
    print(time.asctime(), "GUI done")
          
    # Stop all dangling threads
    os.abort()