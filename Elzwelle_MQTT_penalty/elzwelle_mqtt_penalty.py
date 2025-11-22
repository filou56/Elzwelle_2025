import configparser
import os
import platform
import time
import io
import csv
import traceback
import uuid
import paho.mqtt.client as paho
import tkinter

from   easygui import multenterbox
from   tkinter import ttk
from   tkinter import messagebox
from   tkinter import filedialog
from   os.path import normpath
from   paho    import mqtt
from   tksheet import Sheet

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
        self.editing = False

    def showError(self, *args):
        err = traceback.format_exception(*args)
        messagebox.showerror('Exception',err)
        
        # but this works too
        tkinter.Tk.report_callback_exception = self.showError

    def initialize(self):
        
        self.uuid = str(uuid.uuid4()).replace("-","")
        
        self.geometry("300x500")
        
        self.menuBar = tkinter.Menu(self)
        self.config(menu = self.menuBar)
        
        self.menuFile = tkinter.Menu(self.menuBar, tearoff=False)
        self.menuFile.add_command(command = self.saveSheet, label="Blatt speichern")
        self.menuFile.add_command(command = self.loadSheet, label="Blatt laden")
        self.menuFile.add_command(command = self.clearSheet, label="Blatt löschen")
        self.menuFile.add_command(command = editConfig, label="Konfiguration")
        
        self.menuBar.add_cascade(label="Datei",menu=self.menuFile)
        
        self.pageHeader = tkinter.Label(self,text="Strafsekunden Eingabe",
                                        font=("Arial", 18),
                                        bg='#D3E3FD')
        self.pageHeader.pack(expand = 0, fill ="x") 
        
        self.strNrFrame = tkinter.Frame(self)
        
        self.stNrLabel = tkinter.Label(self.strNrFrame,text="Startnummer",
                                       font = ("Calibri", 14, "bold"))
        self.stNrLabel.pack(side='left',expand = 1,padx = 10) 
        
        self.stNrEdit = tkinter.Entry(self.strNrFrame,font = ("Calibri", 14, "bold"),
                                      validatecommand=self.entryValidate,
                                      validate="focusout")
        
        self.stNrEdit.bind("<Return>", lambda event: self.penaltySheet.focus_set() )
        
        self.stNrEdit.pack(side='left',expand = 1,padx = 10)
                               
        self.strNrFrame.pack(side='top',fill = "x",pady = 10)              
                         
        self.firstGate = config.getint('view','first_gate')                
        self.lastGate  = config.getint('view','last_gate')
                         
        self.penaltySheet = Sheet(self,
                               name = 'penaltySheet',
                               data = [[f'{i+self.firstGate}','0']for i in range(self.lastGate-self.firstGate+1)],
                               header = ['Tor','Zeit'],
                               header_bg = "azure",
                               header_fg = "black",
                               index_bg  = "azure",
                               index_fg  = "gray",
                               font = ("Calibri", 12, "bold")
                            )
        self.penaltySheet.enable_bindings()
        self.penaltySheet.span('A:').align('right')
        self.penaltySheet.span('A').readonly()
        
        self.penaltySheet.disable_bindings("All")
        self.penaltySheet.enable_bindings("edit_cell","single_select","right_click_popup_menu",
                                          "drag_select","row_select","copy","up","down")
        self.penaltySheet.extra_bindings("end_edit_cell", func=self.endEditCell)
        self.penaltySheet.extra_bindings("begin_edit_cell", func=self.beginEditCell)
        
        self.penaltySheet.edit_validation(self.validateEdits)
        
        self.penaltySheet.pack(side='top',fill = "x")
        
        self.sendButton = tkinter.Button(self,text="Senden",
                                         font = ("Calibri", 14, "bold"),
                                         bg   = "steelblue3",
                                         fg   = "White",
                                         command = self.processPenaltyList
                                         )
        self.sendButton.pack(side='top',fill = "x",pady = 10,padx = 10)
        self.sendButton.bind("<Return>", self.buttonSendCommand)
        
    def buttonSendCommand(self,event):   
        print("Button Return") 
        self.processPenaltyList()
    
    def beginEditCell(self, event):
        print(event.key)
        print("Begin EditCell: ")
        if event.key == "Return":
            return 0
        else:
            self.editing = True
            return event.key
        
    def endEditCell(self, event):
        print("End EditCell: ")
        self.editing = False
                     
    def processPenaltyList(self):
        if self.editing:
            print("Alwas editing")
            messagebox.showerror("MQTT", "Eingabe mit Return beenden !")
        else:
            stNr = self.stNrEdit.get()
            if stNr != "":
                if messagebox.askyesno("MQTT", "Strafsekunden senden ?"):
                    for row in self.penaltySheet.data:
                        if int(row[1]) != 0:
                            print("Row: ","{:},{:},{:}".format(0,row[0],row[1]))
                            self.penaltySheet.after_idle(self.sendPenaltyMsg,"{:},{:},{:},,{:}".format(stNr,row[0],row[1],str(self.uuid)))   
            else:                
                messagebox.showerror("MQTT", "Startnummer fehlt !")
                
    def sendPenaltyMsg(self,*args):
        if len(args) == 1:
            #print("Send: ",args[0])
            mqtt_client.publish("elzwelle/stopwatch/course/data", payload=args[0], qos=1)
    
    def validateEdits(self, event):
        print("Validate: ")
        for cell, value in event.cells.table.items():
            try:
                num = int(value.replace(',','.'))
                print(cell[0],self.lastGate-self.firstGate)
                if int(cell[0]) == self.lastGate-self.firstGate:
                    self.penaltySheet.after(100,lambda: self.penaltySheet.deselect(self.lastGate-self.firstGate,1))
                    #self.penaltySheet.after(200,lambda: self.sendButton.config(state="normal"))
                    self.penaltySheet.after(300,lambda: self.sendButton.focus_set())
                return "{:d}".format(num)
            except Exception as error:
                print(error)
                messagebox.showerror(title="Fehler", message="Keine gültige Zahl !")
        return
    
    def saveSheet(self):
        saveSheet = self.getSelectedSheet()
        print("Save: "+saveSheet.name)
        # create a span which encompasses the table, header and index
        # all data values, no displayed values
        sheet_span = saveSheet.span(
            header=True,
            index=True,
            hdisp=False,
            idisp=False,
        )
        
        filepath = filedialog.asksaveasfilename(
            parent=self,
            title="Save sheet as",
            filetypes=[("CSV File", ".csv"), ("TSV File", ".tsv")],
            defaultextension=".csv",
            confirmoverwrite=True,
        )
        if not filepath or not filepath.lower().endswith((".csv", ".tsv")):
            return
        try:
            with open(normpath(filepath), "w", newline="", encoding="utf-8") as fh:
                writer = csv.writer(
                    fh,
                    dialect=csv.excel if filepath.lower().endswith(".csv") else csv.excel_tab,
                    lineterminator="\n",
                )
                writer.writerows(sheet_span.data)
        except Exception as error:
            print(error)
            return

    def loadSheet(self):
        loadSheet = self.getSelectedSheet()
        print("Load: "+loadSheet.name)
        
        sheet_span = loadSheet.span(
            header=True,
            index=True,
            hdisp=False,
            idisp=False,
        )
        
        filepath = filedialog.askopenfilename(parent=self, title="Select a csv file")
        if not filepath or not filepath.lower().endswith((".csv", ".tsv")):
            return
        try:
            with open(normpath(filepath), "r") as filehandle:
                filedata = filehandle.read()
            loadSheet.reset()
            sheet_span.data = [
                r
                for r in csv.reader(
                    io.StringIO(filedata),
                    dialect=csv.Sniffer().sniff(filedata),
                    skipinitialspace=False,
                )
            ]
        except Exception as error:
            print(error)
            return
        
    def clearSheet(self):
        if messagebox.askyesno("Start/Ziel", "Alle Daten löschen ?"):
            self.setRange()
            self.penaltySheet.deselect()
            self.penaltySheet.data = []
            self.penaltySheet.data = [[f'{i+self.firstGate}','0']for i in range(self.lastGate-self.firstGate+1)]
           
    def entryValidate(self):
        print("StNr Input validate")
        for row in range(len(self.penaltySheet.get_column_data(0))):
            self.penaltySheet[row].highlight("white")
            self.penaltySheet.set_cell_data(row,1,0)
        self.penaltySheet.select_cell(0,1)
        #self.sendButton.config(state="disabled")
        return True  
    
    def setRange(self):
        self.firstGate = config.getint('view','first_gate')                
        self.lastGate  = config.getint('view','last_gate')
        
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
    client.publish("elzwelle/penalty", payload="running", qos=1)
    

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
    
    if msg.topic == 'elzwelle/stopwatch/course/data/akn':
        try:
            data = payload.split(',')            
            if len(data) == 5:
                if data[4] == app.uuid:
                    app.penaltySheet[int(data[1].strip())-1].highlight(bg = "aquamarine")
                    print(data)
                    
        except Exception as e:
            print("MQTT Decode exception: ",e,msg.payload)
   
def editConfig():
    global config
      
    msg    = "Felder ausfüllen / ändern"
    title  = "INI File Editor"
    
    keys   = []
    values = []
    sections = []
    
    for each_section in config.sections():
        for (each_key, each_val) in config.items(each_section):
            #print("[{:}]\t{:} = {:}".format(each_section,each_key,each_val))
            sections.append(each_section)
            keys.append(each_key)
            values.append(each_val)
    
    newValues = multenterbox(msg, title, keys, values)
    
    for i in range(len(keys)):
        #print("[{:}]\t{:} = {:}".format(sections[i],keys[i],newValues[i]))
        config.set(sections[i],keys[i],newValues[i])
    
    if messagebox.askyesno("Konfiguration", "Anwendung neu starten falls Änderungen "+
                                            "im Abschnitt MQTT vorgenommen wurden. "+
                                            "Für Werte die den Bereich der Tore betreffen, "+
                                            "genügt es das Blatt zu löschen.\n\n"+
                                            "Änderungen abspeichern ?"):
        with open('linux.ini', 'w') as configfile:    # save
            config.write(configfile)

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
      
    # Platform specific
    if myPlatform == 'Windows':
        iniFile = 'windows.ini' 
    if myPlatform == 'Linux':
        iniFile = 'linux.ini'
    
    config.read(iniFile)
    
    #--------------------------------- MQTT --------------------------

    # using MQTT version 5 here, for 3.1.1: MQTTv311, 3.1: MQTTv31
    # userdata is user defined data of any type, updated by user_data_set()
    # client_id is the given name of the client
    try:
        mqtt_client = paho.Client(client_id="elzwelle_"+str(uuid.uuid4()), userdata=None, protocol=paho.MQTTv311)
    
        # enable TLS for secure connection
        if config.getboolean('mqtt','tls_enabled'):
            #mqtt_client.tls_set(certfile="/etc/ssl/certs/isrgrootx1.pem",tls_version=mqtt.client.ssl.PROTOCOL_TLS)
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
    
    if not config.getboolean('view','start_enabled'):
        app.tabControl.tab(0, state="hidden")
        
    if not config.getboolean('view','finish_enabled'):
        app.tabControl.tab(1, state="hidden") 
    
    app.title("MQTT Strafsekunden")
    
    app.penaltySheet.popup_menu_add_command(
        "Clear sheet data",
        app.clearSheet,
    )
    
    # run
    app.mainloop()
    print(time.asctime(), "GUI done")
          
    # Stop all dangling threads
    os.abort()