#pip install python-dateutil
#pip install python-registery
#pip install python-evtx
#%APPDATA%/Local/Programs/Python/Python35/Lib/site-packages/Evtx/BinaryParser.py
#change Line 110 from except ValueError to just except:



import Evtx.Evtx as evtx
import Evtx.Views as e_views
import evtx2xml
import xml.etree.ElementTree as ET

import itertools
from dateutil import tz
from datetime import datetime


import os
from tkinter import *
from tkinter import ttk
from tkinter.ttk import Combobox
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askopenfilenames
from tkinter.filedialog import askdirectory
from tkinter.filedialog import asksaveasfilename
import tkinter.scrolledtext as scrolledtext
from promptuser import pufiles, pufile, pudir, pusavefile

import time
import sys

from Registry import Registry



global sysdidrun
global powerschemes
sysdidrun = False

#python evtx2xml.py System.evtx
# -> system.xml

def cuttimestr(timestring, b):
    #timestr = str(datetime)
    #b = "D" or "M" or "S"
    d={
        "D":10,
        "M":16,
        "S":19,
    }
    border = d[b]
    return timestring[:border]


def utc2local(timestring):
    '''
    Converts UTC Time to Local Time (Zurich)
    '''
    #timestring = string and returvalue:
    #datetime.datetime object
    ltz = tz.gettz("Europe/Zurich")
    return datetime.strptime(timestring, "%Y-%m-%d %H:%M:%S.%f").replace(tzinfo=tz.tzutc()).astimezone(ltz)

#Standyby Reasons (ID 42)
reasons = {
        4:"Application API",
        7:"System idle",
        6:"Hibernate from Sleep - Fixed Timeout",
        0:"Button or Lid",
    }

#EventIDs of Interest
evids = {
        #-1: "Unerwartet heruntergefahren",
        -2: "Letzter Log Eintrag",
        None: "None",
        1:"time set",
        12:"Start",
        12.1:"Changed Power Settings",
        13:"Heruntergefahren",
        42:"In Standbymodus versetzt",
        107:"Aus Standbymodus aufgeweckt",
        6008:"Unerwartet heruntergefahren"
    }

#Defines Class Entry - every Event will be considered an Entry with attributes self.eventid ...
class Entry():
    def __init__(self):
        self.eventid = None
        self.utct = None
        self.data = {}
        self.localtime = None


    def __repr__(self):
        return '\nEventID:' + str(self.eventid) + '\tEventDescr:' + str(evids.get(self.eventid, "None")) + '\nUTC Time:' + str(self.utct) + '\nData:' + str(self.data)


def analyse_xml(o, evids):
    #o=path to xml
    '''
    Does all the work of analysing the xml log
    '''

    ns="{http://schemas.microsoft.com/win/2004/08/events/event}"

    unexpected_shutdown_id = -1

    global PowerGuid
    PowerGuid = {
        "{a1841308-3541-4fab-bc81-f71556f20b4a}":"Power Saver",
        "{381b4222-f694-41f0-9685-ff5bb260df2e}":"Balanced",
        "{8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c}":"High Performance",
    }

    
    dataids = {
        1: ["NewTime", "OldTime"],
        12.1: ["OldSchemeGuid", "NewSchemeGuid", "ProcessPath"],
        42: ["Reason"],
        6008: []
        }
    
    global reasons
    

    dataids = {str(each):v for each,v in dataids.items()}
    reasons = {str(each):v for each,v in reasons.items()}
    evids = {str(each):v for each,v in evids.items()}

    #
    #
    #

    tree = ET.parse(o)
    xmlroot = tree.getroot()

    loe = []

    # Loop over all Events, add Events of interest to 
    for event in xmlroot:
        e = Entry()
        for system in event.iter(ns+"System"):
            for eid in system.iter(ns+"EventID"):
                #EVENT ID = e.eventid
                #DO stuff with it
                e.eventid=eid.text
                if eid.text=="12":
                    #IF EVENT ID == 12:
                    for provider in system.iter(ns+"Provider"):
                        if provider.get("Name")=="Microsoft-Windows-Kernel-General":
                            break
                            #COMPUTER TURN ON
                        if provider.get("Name")=="Microsoft-Windows-UserModePowerService":
                            #POWER SETTING CHANGED
                            e.eventid="12.1"
 

            for timecreated in system.iter(ns+"TimeCreated"):
                e.utct = timecreated.get("SystemTime")

                #Get additional Data from <EventData>
                #Special Treat for 6008 because there is no Name for the fields
        if e.eventid != None:
            if e.eventid in dataids:
                if e.eventid=="6008":
                    for ed in event.iter(ns+"EventData"):
                        for dataf in ed:
                            ds = bytes(dataf.text, "utf-8")
                            ds = ds.replace(b'\xe2\x80\x8e', b"")
                            break
                        print("___", ds)
                        rt = str(ds[34:44])[2:-1]+" "+str(ds[8:16])[2:-1]
                        ltz = tz.gettz("Europe/Zurich")
                        e.data["RealTime"]=datetime.strptime(rt, "%d.%m.%Y %H:%M:%S").replace(tzinfo=ltz).astimezone(ltz)
                        
    
                else:
                    for ed in event.iter(ns+"EventData"):
                        
                        for dataf in ed:                    
                            if dataf.get("Name") in dataids[e.eventid]:
                                e.data[dataf.get("Name")]=dataf.text
                                #sys.exit()

            #print(e)
            loe.append(e)
        


    #Later Information Processing
    #Fix Start Time
    weirdid = {
        12:"Start",
        13:"Heruntergefahren",
        42:"In Standbymodus versetzt",
        107:"Aus Standbymodus aufgeweckt",
        6008:"Unerwartet heruntergefahren"}
    #event IDs that need a fix in time:
    timefixstart = [107,]
    timefixstart = [str(each) for each in timefixstart]
    loeit = iter(loe)
    c=0
    for e in loeit:
        print(e)
        if e.eventid in timefixstart:
            print("1")
            weird = False
            for ofs in range(20):
                n = next(loeit)
                c+=1
                if n.eventid in weirdid:
                    weird = True
                if n.eventid == "1":
                    print("2")
                    if "NewTime" in n.data and "OldTime" in n.data:
                        print("3")
                        if not weird or cuttimestr(n.data["OldTime"], "M") == cuttimestr(e.utct, "M"):
                            print("4")
                            e.utct=n.data["NewTime"]
                            loe[c-ofs-1]=e
                            break
                        else:
                            e.data["TFE"]="Zeit vermutlich nicht korrekt"
                            loe[c-ofs-1]=e
                            break
            #check up to x next events for a time change
                        
        c+=1

    #unexpected shutdown add event
    def unxse(timestamp):
        e = Entry()
        e.eventid = str(unexpected_shutdown_id)
        e.utct = timestamp
        return e
    #ADD UNEXPECTED SHUTDOWN AS LAST LOG ENTRY BEFORE START EVENT -> REPLACED WITH E ID 6008
    '''i=0
    while i < len(loe)-1:
        if loe[i+1].eventid == "12":
            if loe[i].eventid not in ["13", "-1"]:
                e = unxse(loe[i].utct)
                loe.insert(i+1, e)
        i=i+1'''




    #Last Log Entry
    lle = loe.pop()
    lle.eventid="-2"
    loe.append(lle)





    #only keep event ids of interes
    loe = [each for each in loe if each.eventid in evids]



    #Change all to real Timezone
    for e in loe:
        e.localtime = utc2local(e.utct)
 


    #Fix time of unexpected shutdown
    for e in loe:
        if e.eventid == "6008":
            e.localtime = cuttimestr(str(e.data["RealTime"]), "S")
            e.localtime = e.data["RealTime"]
    

    # Sort by Timestamp: (should already be, except for unexpected shutdown 6008)
    loe.sort(key=lambda x: x.localtime, reverse=False)



    #if sysdidrun:
    #enrich energy setings standby time to log info
    global powerschemes
    global sysdidrun
    '''if sysdidrun:
        global powerschemes
        for e in loe:
            if e.eventid=="12.1":
                e.data["NewScheme"]=powerschemes[e.data["NewSchemeGuid"][1:-1]]'''
    #only in the newest ones where it is new
    
    

    if sysdidrun:
        us = {x:True for x in powerschemes}
        for e in reversed(loe):
            if e.eventid=="12.1":
                if e.data["NewSchemeGuid"][1:-1] in us:
                    us.pop(e.data["NewSchemeGuid"][1:-1])
                    e.data["NewScheme"]=powerschemes[e.data["NewSchemeGuid"][1:-1]]


    

    #print results
    # 1 is timeset
    for e in (x for x in loe if x.eventid!="1"):
        if e.eventid=="12.1":
            print(cuttimestr(str(e.localtime), "S"), evids[e.eventid], PowerGuid.get(e.data["OldSchemeGuid"],"Unknown"), "->", PowerGuid.get(e.data["NewSchemeGuid"],"Unknown"), e.data.get("NewScheme", ""))
        elif e.eventid=="42":
            print(cuttimestr(str(e.localtime), "S"), evids[e.eventid], "Reason:", reasons.get(e.data["Reason"], e.data["Reason"]))
        elif e.eventid=="107":
            print(cuttimestr(str(e.localtime), "S"), evids[e.eventid], e.data.get("TFE", ""))
        else:
            print(cuttimestr(str(e.localtime), "S"), evids[e.eventid])

    
    return [x for x in loe if x.eventid!="1"]




'''
    #HKEY CURRENT USER:
    "Control Panel\\PowerCfg": "Power Configuration Scheme",#in NTUSER.DAT
    "Control Panel\\PowerCfg\\PowerPolicies\\0":"pp0",
    "Control Panel\\PowerCfg\\PowerPolicies\\1":"pp1",
    "Control Panel\\PowerCfg\\PowerPolicies\\2":"pp2",
    "Control Panel\\PowerCfg\\PowerPolicies\\3":"pp3",
    "Control Panel\\PowerCfg\\PowerPolicies\\4":"pp4",
    "Control Panel\\PowerCfg\\PowerPolicies\\5":"pp5",
'''

def analyse2file(loe, o):
    #To CSV File ; for excel
    o=o[:-3]+"csv"
    global PowerGuid
    global reasons
    global evids
    reasons = {str(each):v for each,v in reasons.items()}
    evids = {str(each):v for each,v in evids.items()}
    pps = {} #previous power setting with times for idle standby
    cps = ""
    with open(o, "w") as f:
        f.write("Zeit; Event; Ursache; Info\n")
        for e in loe:
            if e.eventid=="12.1":
                #f.write(cuttimestr(str(e.localtime), "M") +";"+ evids[e.eventid] +";"+""+";"+"\n")#
                pps[PowerGuid.get(e.data["NewSchemeGuid"], "Unknown")] = e.data.get("NewScheme", "")
                cps = PowerGuid.get(e.data["NewSchemeGuid"], "Unknown")
            elif e.eventid=="42":#standby
                fd = pps.get(cps, "")
                fds = ""
                for k in fd:
                    if type(fd[k])==dict:
                        fds = fds + " " + k + ": " + "".join([str(key)[:2]+": "+str(int(v)//60)+" Minuten" for key,v in fd[k].items()])
                ############################### CONVERT DICTIONARY TIME DATA TO STRING FORM
                if reasons.get(e.data["Reason"], e.data["Reason"])!="System idle":
                    f.write(cuttimestr(str(e.localtime), "S") +";"+ evids[e.eventid] +";"+ reasons.get(e.data["Reason"], e.data["Reason"]) +";" + "" + "\n")
                else:
                    f.write(cuttimestr(str(e.localtime), "S") +";"+ evids[e.eventid] +";"+ reasons.get(e.data["Reason"], e.data["Reason"]) +";"+ fds +"\n")
            elif e.eventid=="107":#wake up
                f.write(cuttimestr(str(e.localtime), "S") +";"+ evids[e.eventid] +";"+""+";"+ e.data.get("TFE", "") +"\n")
            else:
                f.write(cuttimestr(str(e.localtime), "S") +";"+ evids[e.eventid] +";"+""+";"+"\n")
            
    
    #TO Raw Text GUI
    txt=""
    
    for e in loe:
        if e.eventid=="12.1":
            #f.write(cuttimestr(str(e.localtime), "M") +";"+ evids[e.eventid] +";"+""+";"+"\n")#
            pps[PowerGuid.get(e.data["NewSchemeGuid"], "Unknown")] = e.data.get("NewScheme", "")
            cps = PowerGuid.get(e.data["NewSchemeGuid"], "Unknown")
        elif e.eventid=="42":#standby
            fd = pps.get(cps, "")
            fds = ""
            for k in fd:
                if type(fd[k])==dict:
                    fds = fds + " " + k + ": " + "".join([str(key)[:2]+": "+str(int(v)//60)+" Minuten" for key,v in fd[k].items()])
            ############################### CONVERT DICTIONARY TIME DATA TO STRING FORM
            if reasons.get(e.data["Reason"], e.data["Reason"])!="System idle":
                txt = txt + (cuttimestr(str(e.localtime), "S") +"\t"+ evids[e.eventid] +"\t("+ reasons.get(e.data["Reason"], e.data["Reason"]) +")\t" + "" + "\n")
            else:
                if fds == "":
                    t=""
                else:
                    t="\t"
                txt = txt + (cuttimestr(str(e.localtime), "S") +"\t"+ evids[e.eventid] +"\t("+ reasons.get(e.data["Reason"], e.data["Reason"]) + t + fds +")\n")
        elif e.eventid=="107":#wake up
            txt = txt + (cuttimestr(str(e.localtime), "S") +"\t"+ evids[e.eventid] +"\t"+e.data.get("TFE", "")+"\t"+"\n")
        else:
            txt = txt + (cuttimestr(str(e.localtime), "S") +"\t"+ evids[e.eventid] +"\t"+""+"\t"+"\n")
    helplabel.config(state="normal", height=150)
    helplabel.delete(1.0, END)
    helplabel.insert(END, txt)
    helplabel.config(state="disabled")


#ACTUAL SCREEN TIMEOUT TIME:
#Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\381b4222-f694-41f0-9685-ff5bb260df2e\7516b95f-f776-4464-8c53-06167f40cc99\3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e


ROIV = {
    "CurrentPowerPolicy":{
        "0":"Home/Office Desk",
        "1":"Portable/Laptop",
        "2":"Presentation",
        "3":"Always On",
        "4":"High Performance",
        "5":"Max Battery",
    },
    2:2,

}


def recursive_registry(key, dl, depth=0):
    sk = str(bytes(key.path(), "utf-8"))
    if sk.find("")>=0:
        print(depth*"\t" + sk)
    
    if depth <= dl:
        for subkey in key.subkeys():
            recursive_registry(subkey, depth+1)

def analyse_registry_sys(f):
    reg = Registry.Registry(f)

    #recursive_registry(reg.root(), 1)
    for value in reg.open("Select").values():
        print(value.name(), value.value())
        if value.name()=="Current":
            ccs = value.value()
            #ccs = current control set number

    basepath = "ControlSet"+"00"+str(ccs)+"\\Control\\Power\\User\\PowerSchemes"


    for value in reg.open(basepath).values():
        if value.name()=="ActivePowerScheme":
            aps = value.value()
    
    print("___")
    print(ccs, aps)


    # Current setting = aps - match with log analysis?

    # Get all possible settings:
    # powerschemes = ["381b4222-f694-41f0-9685-ff5bb260df2e", "3af9B8d9-7c97-431d-ad78-34a8bfea439f", "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"] 
    print("possible schemes:")
    p = basepath
    global powerschemes
    powerschemes = {}
    #powerschemes = {scheme: {name:Performance}, scheme: {name:Balanced, Standby:{AC: 1, DC: 2} } }
    for scheme in reg.open(p).subkeys():
        short = scheme.path()[scheme.path().rfind("\\")+1:]
        powerschemes[short]={}
        for value in scheme.values():
            if value.name()=="FriendlyName":
                powerschemes[short]["name"]=value.value()[value.value().rfind(",")+1:]
        # Get Times for each setting:
        ''' # maybe these things dont exist -> never
        PowerSchemes
            Scheme1
                useless
                    Seconds until Hibernate
                useless
                    Seconds until Screen
            Scheme2
                useless
                    Seconds until Hibernate
                useless
                    Seconds until Screen
        '''
        '''
        9d7815a6-7ee4-497e-8888-515a05f02364 = Ruhezustand nach Sekunden
        3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e = Bildschirm ausschalten nach Sekunden
        29f6c1db-86da-48c5-9fdb-f2b67b1f44da = Energiesparmodus
        '''
        wantedinfo = {
            "9d7815a6-7ee4-497e-8888-515a05f02364":"Ruhezustand nach",
            "3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e":"Bildschirm aus nach",
            "29f6c1db-86da-48c5-9fdb-f2b67b1f44da":"Energiesparmodus nach",
        }
        #loop over sub ksub keys
        for useless in scheme.subkeys():
            for fact in useless.subkeys():
                print("fact:", fact.path())
                shortfact = fact.path()[fact.path().rfind("\\")+1:]
                for wanted in wantedinfo:
                    if shortfact == wanted:
                        powerschemes[short][wantedinfo.get(shortfact, shortfact)]={}
                        for value in fact.values():
                            if value.value()!=0:
                                powerschemes[short][wantedinfo.get(shortfact, shortfact)][value.name()]=value.value()
        
    powerschemes = {scheme:{name:v2 for name,v2 in v.items() if v2!={}} for scheme,v in powerschemes.items()}
    print(powerschemes)


    KOI = {
        "ControlSet"+"00"+str(ccs)+"\\Control\\Power\\User\\PowerSchemes\\381b4222-f694-41f0-9685-ff5bb260df2e\\7516b95f-f776-4464-8c53-06167f40cc99\\3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e":("ACSettingIndex", "Sec until Monitor Off"),
        
    }
    

    for k in KOI:
        pass


def analyse_registery_nt(f):
    global ROIV
    reg = Registry.Registry(f)

    #recursive_registry(reg.root(), 1)
    

     
    ROI = {
    "Control Panel\\PowerCfg": "Power Configuration Scheme",#in NTUSER.DAT
    "Control Panel\\PowerCfg\\PowerPolicies\\0":"pp0",
    "Control Panel\\PowerCfg\\PowerPolicies\\1":"pp1",
    "Control Panel\\PowerCfg\\PowerPolicies\\2":"pp2",
    "Control Panel\\PowerCfg\\PowerPolicies\\3":"pp3",
    "Control Panel\\PowerCfg\\PowerPolicies\\4":"pp4",
    "Control Panel\\PowerCfg\\PowerPolicies\\5":"pp5",
}

    regresults = {}

    for k in ROI:
        print("___________________________________")
        print(k)
        print(ROI[k])
        try:
            key = reg.open(k)
            print("Found------------------------------------")
        except Registry.RegistryKeyNotFoundException:
            print("Reg Key not found")
            #sys.exit()
            continue
        for value in [v for v in key.values() \
                   if v.value_type() == Registry.RegSZ or \
                      v.value_type() == Registry.RegBin or \
                      v.value_type() == Registry.RegExpandSZ]:
            print(value.value())
            #print(ROIV)
            print(value.name())
            print(value.name(), ":", ROIV.get(value.name(), {}).get(value.value(), value.value()))

            if value.name() == "CurrentPowerPolicy":
                regresults[value.name()]=value.value()
            if value.name() == "Policies" and ROI[k][:2]=="pp":
                regresults[ROI[k]]=int.from_bytes(value.value()[56:58], byteorder="little")
    print("________________________")
    for r, v in regresults.items():
        print(r, v)




#cmd.exe powercfg /q
# HKEy_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes ActivePowerScheme 8c5e7fda-....


#PARSE REGISTERY FOR ENERGY SAVING INFO MINUTES TILL SLEEP ETC
# min until standby
#Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\PowerSettings\238C9FA8-0AAD-41ED-83F4-97BE242C8F20\29f6c1db-86da-48c5-9fdb-f2b67b1f44da\DefaultPowerSchemeValues
# min until hibernation
#Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\PowerSettings\238C9FA8-0AAD-41ED-83F4-97BE242C8F20\9d7815a6-7ee4-497e-8888-515a05f02364\DefaultPowerSchemeValues

''' Registry Data does not reflect actual settings
'''
def buttonreg_nt():
    f = pufile()
    analyse_registery_nt(f)


def buttonreg_sys():
    f = pufile()
    analyse_registry_sys(f)
    global sysdidrun 
    sysdidrun = True

def buttonevtx():
    global evids
    global statustext
    evtxpath = pufile()
    f=evtxpath
    o="System_"+datetime.now().strftime("%Y-%m-%d_%H-%M-%S")+".xml"
    statustext.set("Status: evtx to xml")
    root.update_idletasks()
    root.update()   
    evtx2xml.evtx2xml(f, o)
    statustext.set("Status: analyse xml")
    root.update_idletasks()
    root.update()   
    loe = analyse_xml(o, evids)
    statustext.set("Status: Write to File")
    root.update_idletasks()
    root.update()
    os.remove(o)
    analyse2file(loe, o)
    statustext.set("Status: idle")

#################
###### GUI ######
#################

root=Tk()

class Window(Frame):
    '''
    THE GUI
    '''
    def __init__(self, master=None):
        
        Frame.__init__(self, master)
        self.master = master
        self.pack(fill=BOTH, expand=1)



# Windows Registery Row
winregrow=Frame(root)
winregrow.pack(side=TOP, fill=X, padx=5, pady=5)
# Registery File
'''
regbutton = Button(winregrow, text="Select NTUSER.DAT", command=buttonreg_nt)
regbutton.pack(side=LEFT, padx=10, pady=10)
'''
# Registery File
regbuttonsys = Button(winregrow, text="Select SYSTEM", command=buttonreg_sys)
regbuttonsys.pack(side=LEFT, padx=10, pady=10)

# Windows Event Log Row
winlogrow=Frame(root)
winlogrow.pack(side=TOP, fill=X, padx=5, pady=5)
# Evtx 2 XML
evtxbutton = Button(winlogrow, text="Select System.evtx", command=buttonevtx)
evtxbutton.pack(side=LEFT, padx=10, pady=10)

#Status Text Label
statusrow = Frame(root)
statusrow.pack(side=TOP, fill=X, padx=5, pady=5)
statustext = StringVar()
statustext.set("Status: idle")
statuslabel = Label(statusrow, textvariable=statustext, font=("Helvetica", 12), anchor=W, justify=LEFT, width=30)
statuslabel.pack(side=LEFT, padx=10, pady=10)

#Help Text Label
helprow = Frame(root)
helprow.pack(side=TOP, fill=X, padx=5, pady=5)
helptext = StringVar()
helptext.set("File Locations:\n\nC:\\Windows\\System32\\winevt\\Logs\\System.evtx\nC:\\Windows\\System32\\config\\System")
#helplabel = Label(helprow, textvariable=helptext, font=("Helvetica", 12), anchor=W, justify=LEFT)
helplabel = scrolledtext.ScrolledText(helprow, undo=True)
helplabel.insert(END, helptext.get())
helplabel.config(state="disabled")
helplabel.pack(side=LEFT, padx=10, pady=10, expand=1, fill=BOTH)

#GUI Window Title and Size:
root.wm_title("Windows Event Log Analyser")
root.geometry(str(int(500))+"x"+str(int(500)))

#Make the left-Labels the same size:
root.update()


root.mainloop()
