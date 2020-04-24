#python evtx2xml.py System.evtx
# -> system.xml
import Evtx.Evtx as evtx
import Evtx.Views as e_views
#pip install python-evtx
#%APPDATA%/Local/Programs/Python/Python35/Lib/site-packages/Evtx/BinaryParser.py
#change Line 110 from except ValueError to just except:

def evtx2xml(f, o):
    with evtx.Evtx(f) as log:
        with open(o, "w") as of:
            of.write(e_views.XML_HEADER)
            of.write("<Events>\n")
            for record in log.records():
                of.write(record.xml()+"\n")
            of.write("</Events>")
