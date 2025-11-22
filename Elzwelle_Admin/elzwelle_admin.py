import gspread
import configparser
import gc
import os
import platform
import time
import tkinter

from   tkinter import messagebox
from   pathlib import Path

# Google Spreadsheet ID for publishing times
# Elzwelle        SPREADSHEET_ID = '1obtfHymwPSGoGoROUialryeGiMJ1vkEUWL_Gze_hyfk'
# FilouWelle      SPREADSHEET_ID = '1M05W0igR6stS4UBPfbe7-MFx0qoe5w6ktWAcLVCDZTE'
# SPREADSHEET_ID = '1M05W0igR6stS4UBPfbe7-MFx0qoe5w6ktWAcLVCDZTE'

#-------------------------------------------------------------------
# Define the GUI
#-------------------------------------------------------------------
class simpleapp_tk(tkinter.Tk):
    global wks_start
    global wks_finish
    
    def __init__(self,parent):
        tkinter.Tk.__init__(self,parent)
        self.parent = parent
        self.initialize()

    def initialize(self):
        self.grid()

        #Add a label with the text leftbound black font(fg) on white background(bg) at (0,0) over 2 columns,
        #sticking to the left and to the right of the cell
        self.labelVariable = tkinter.StringVar()
        label = tkinter.Label(self,text="Administration",anchor="c",fg="black",bg="white")
        label.grid(row=0,column=0,columnspan=3,sticky="EW")

        label1 = tkinter.Label(self,text="Start",relief=tkinter.SUNKEN,bg="white")
        label1.grid(row=1,column=0,sticky="EW")

        label2 = tkinter.Label(self,text="Ziel",relief=tkinter.SUNKEN,bg="white")
        label2.grid(row=1,column=1,sticky="EW")
     
        label3 = tkinter.Label(self,text="Strecke",relief=tkinter.SUNKEN,bg="white")
        label3.grid(row=1,column=2,sticky="EW")
        
        button1 = tkinter.Button(self,text="Format",command=self.PrepareStartButtonClicked)
        button1.grid(row=2,column=0,sticky="EW")

        button2 = tkinter.Button(self,text="Format",command=self.PrepareFinishButtonClicked)
        button2.grid(row=2,column=1,sticky="EW")
        
        button3 = tkinter.Button(self,text="Format",command=self.PrepareCourseButtonClicked)
        button3.grid(row=2,column=2,sticky="EW")
        
        button4 = tkinter.Button(self,text="Loeschen",command=self.ClearStartButtonClicked)
        button4.grid(row=3,column=0,sticky="EW")

        button5 = tkinter.Button(self,text="Loeschen",command=self.ClearFinishButtonClicked)
        button5.grid(row=3,column=1,sticky="EW")
        
        button6 = tkinter.Button(self,text="Loeschen",command=self.ClearCourseButtonClicked)
        button6.grid(row=3,column=2,sticky="EW")
              
        #Make the first column (0) resize when window is resized horizontally
        self.grid_columnconfigure(0,weight=1)
        self.grid_columnconfigure(1,weight=1)
        self.grid_columnconfigure(2,weight=1)

        self.geometry("500x500")
        #Make the user only being able to resize the window horrizontally
        self.resizable(True,True)
     
    def PrepareStartButtonClicked(self):
        print("Prepare Start Spreadsheet")
        wks_start.update([["Uhrzeit","Zeitstempel","Startnummer","Kommentar"]],"A1")
        wks_start.update([["00:00:00","0,00","0"," "]],"A2")
        
        wks_start.format("D2:D",  { 
                "horizontalAlignment": "Left",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        )
        wks_start.format("C2:C",  { 
                "numberFormat": { "type": "NUMBER","pattern": '#0' },
                "horizontalAlignment": "RIGHT",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        )
        wks_start.format("B2:B",  { 
                "numberFormat": { "type": "NUMBER","pattern": "#.00" },
                "horizontalAlignment": "RIGHT",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        )
        wks_start.format("A2:A",  { 
                "numberFormat": { "type": "TIME" },
                "horizontalAlignment": "RIGHT",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        ) 
        wks_start.format("A1:D1",  { 
                "horizontalAlignment": "CENTER",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        )

    def PrepareFinishButtonClicked(self):
        print("Prepare Finish Spreadsheet")
        wks_finish.update([["Uhrzeit","Zeitstempel","Startnummer","Kommentar"]],"A1")
        wks_finish.update([["00:00:00","0,00","0"," "]],"A2")
        
        wks_finish.format("D2:D",  { 
                "horizontalAlignment": "Left",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        )
        wks_finish.format("C2:C",  { 
                "numberFormat": { "type": "NUMBER","pattern": '#0' },
                "horizontalAlignment": "RIGHT",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        )
        wks_finish.format("B2:B",  { 
                "numberFormat": { "type": "NUMBER","pattern": "#.00" },
                "horizontalAlignment": "RIGHT",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        )        
        wks_finish.format("A2:A",  { 
                "numberFormat": { "type": "TIME" },
                "horizontalAlignment": "RIGHT",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        ) 
        wks_finish.format("A1:D1",  { 
                "horizontalAlignment": "CENTER",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        )
        
    def PrepareCourseButtonClicked(self):
        print("Prepare Course Spreadsheet")
        wks_course.update([["Startnummer","Tornummer","Strafzeit","Kommentar"]],"A1")
        wks_course.update([["0","0","0"," "]],"A2")
        
        wks_course.format("D2:D",  { 
                "horizontalAlignment": "Left",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        )
        wks_course.format("C2:C",  { 
                "numberFormat": { "type": "NUMBER","pattern": '#0' },
                "horizontalAlignment": "RIGHT",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        )
        wks_course.format("B2:B",  { 
                "numberFormat": { "type": "NUMBER","pattern": "#0" },
                "horizontalAlignment": "RIGHT",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        )        
        wks_course.format("A2:A",  { 
                "numberFormat": { "type": "NUMBER","pattern": "#0" },
                "horizontalAlignment": "RIGHT",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        ) 
        wks_course.format("A1:D1",  { 
                "horizontalAlignment": "CENTER",
                "textFormat": {
                    "bold": True,
                    "fontSize": 12,
                },
            }
        )
               
    def ClearStartButtonClicked(self):
        print("Clear Finish Spreadsheet")
        wks_start.batch_clear(['A3:A','B3:B','C2:C','D2:D'])
        
    def ClearFinishButtonClicked(self):
        print("Clear Finish Spreadsheet")
        wks_finish.batch_clear(['A3:A','B3:B','C2:C','D2:D'])
        
    def ClearCourseButtonClicked(self):
        print("Clear Course Spreadsheet")
        wks_course.batch_clear(['A3:A','B3:B','C2:C','D2:D'])    
        
#-------------------------------------------------------------------
# Main program
#-------------------------------------------------------------------
if __name__ == '__main__':    
    GPIO = None
   
    myPlatform = platform.system()
    print("OS in my system : ", myPlatform)
    myArch = platform.machine()
    print("ARCH in my system : ", myArch)

    config = configparser.ConfigParser()
   
    #config['google'] = {'spreadsheet_id':SPREADSHEET_ID}
    
    # Platform specific
    if myPlatform == 'Windows':
        # Platform defaults
        config.read('windows.ini') 
    if myPlatform == 'Linux':
        # Platform defaults
        config.read('linux.ini')

    gc = gspread.service_account(filename='../../.elzwelle/client_secret.json')
    
    client_secret_file = config.get('google','client_secret_json')
    try:
        if client_secret_file.startswith(".elzwelle"):
            home_dir = Path.home()
            print( f'Home path: { home_dir }' )
            client_secret_file = os.path.join(home_dir,client_secret_file)
        print( f'Client secret path: { client_secret_file }' )
        google_client = gspread.service_account(filename=client_secret_file)
        
        # Open a sheet from a spreadsheet in one go
        wks_start = gc.open(config.get('google','spreadsheet_name')).get_worksheet(0)
        #print("Ranges: ",gc.open("timestamp").list_protected_ranges(0))
        # Open a sheet from a spreadsheet in one go
        wks_finish = gc.open(config.get('google','spreadsheet_name')).get_worksheet(1)
        
        # Open a sheet from a spreadsheet in one go
        wks_course = gc.open(config.get('google','spreadsheet_name')).get_worksheet(2)
    
        # setup and start GUI
        app = simpleapp_tk(None)
        app.title("Elzwelle Zeitmessung")
    #    app.refresh()
        app.mainloop()
    except Exception as e:
        messagebox.showerror(title="Fehler", message="Keine Verbindung zu Google Sheets!")
        print("Error: ",e)
        
    print(time.asctime(), "GUI done")
          
    # Stop all dangling threads
    os.abort()