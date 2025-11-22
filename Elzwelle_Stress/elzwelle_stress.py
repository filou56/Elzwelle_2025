import configparser
# import gc
# import os
import platform
import time
#import io
import csv
#import traceback
import uuid
import paho.mqtt.client as paho
import tksheet

#from   os.path import normpath
from   paho    import mqtt

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
    # global time_stamps_start_dirty 
    # global time_stamps_finish_dirty
    """
        Prints a mqtt message to stdout ( used as callback for subscribe )

        :param client: the client itself
        :param userdata: userdata is set when initiating the client, here it is userdata=None
        :param msg: the message with topic and payload
    """
    print(msg.topic + " " + str(msg.qos) + " " + str(msg.payload))
    
    payload = msg.payload.decode('ISO8859-1')        # ('utf-8')
    print(payload)
    
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
        'user':'welle',
        'password':'elzwelle', 
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
        print("Error: ",e)
        exit(1)   
     
    race = 1
    gStartNr            = tksheet.alpha2num("D")-1
    if race == 1:
        print("COM: Lauf 1")
        gTzStartIdx     = tksheet.alpha2num("M")-1
        gTzFinishIdx    = tksheet.alpha2num("N")-1
        gTor_Idx        = tksheet.alpha2num("O")-1
        gTor_Ziel       = tksheet.alpha2num("AN")-1
    if race == 2:
        print("COM: Lauf 2")
        gTzStartIdx     = tksheet.alpha2num("AR")-1
        gTzFinishIdx    = tksheet.alpha2num("AS")-1
        gTor_Idx        = tksheet.alpha2num("AT")-1
        gTor_Ziel       = tksheet.alpha2num("BS")-1
 
    
    r = 0
    with open('eingaben.csv', newline='') as csvfile:
        inputReader = csv.reader(csvfile, delimiter=';', quotechar='"');
        for row in inputReader:
            if row[gStartNr].endswith("ff"):
                strNr = row[gStartNr][0:3]
            else:
                strNr = row[gStartNr]
            print(strNr,row[gTzStartIdx],row[gTzFinishIdx],row[gTor_Idx],row[gTor_Ziel])
            r = r + 1
            #if r > 100:
            # elzwelle/stopwatch/start 18:12:12 154,06 0
            # elzwelle/stopwatch/start/number 18:12:12 154,06 1 
            # elzwelle/stopwatch/start/number/akn 18:12:12 154,06 1
            mqtt_client.publish("elzwelle/stopwatch/start",
                                    payload='{:} {:} {:}'.format("10:00:00",row[gTzStartIdx],0), 
                                    qos=1)
            time.sleep(2)
            mqtt_client.publish("elzwelle/stopwatch/start/number",
                                    payload='{:} {:} {:}'.format("10:00:00",row[gTzStartIdx],strNr), 
                                    qos=1)
            time.sleep(1)
            
            # elzwelle/stopwatch/course/data 1,1,50,Tor verfehlt 
            # elzwelle/stopwatch/course/data/akn 1,1,50,Tor verfehlt 
            for i in range(26):
                if row[gTor_Idx+i] != "":
                    if int(row[gTor_Idx+i]) != 0:
                        mqtt_client.publish("elzwelle/stopwatch/course/data",
                                            payload='{:},{:},{:},Bemerkung'.format(strNr,i+1,row[gTor_Idx+i]), 
                                            qos=1)
                        time.sleep(1)
            
            mqtt_client.publish("elzwelle/stopwatch/finish",
                                    payload='{:} {:} {:}'.format("10:00:00",row[gTzFinishIdx],0), 
                                    qos=1)
            time.sleep(2)
            mqtt_client.publish("elzwelle/stopwatch/finish/number",
                                    payload='{:} {:} {:}'.format("10:00:00",row[gTzFinishIdx],strNr), 
                                    qos=1)
            time.sleep(1)
            
    