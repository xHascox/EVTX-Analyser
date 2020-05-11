# EVTX-Analyser

You have the Windows Eventlogs (System.evtx) of a machine and you need to know when it was turned on? Then this program will make your life easy.

### How To:

Run the [Windows Executable](https://github.com/xHascox/EVTX-Analyser/blob/master/EXE/EVTX_Analyser.exe) or the [Script](https://github.com/xHascox/EVTX-Analyser/tree/master/Source).

Select the SYSTEM Registry File (C:\Windows\System32\config\System) of the target Machine if you want to enrich the data with Power Settings (System Idle time until Standby) - this will rarely be working since Power Settings may be changed often and only the most recent one is stored in the Registry.

Select the System.evtx Windows Eventlog file (C:\Windows\System32\winevt\Logs\System.evtx).

Wait some time (10MB ~ 5min)

The Resulting list will be displayed in the GUI and stored to a Excel compatible CSV file where the Script/Executable was run.

### This Program uses:

* [python-evtx](https://github.com/williballenthin/python-evtx) Apache License 2.0
* [python-registry](https://github.com/williballenthin/python-registry) Apache License 2.0

### Tested Python Version:

* 3.5
