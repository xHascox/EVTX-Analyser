# EVTX-Analyser

### This Program uses:

* [python-evtx](https://github.com/williballenthin/python-evtx) Apache License 2.0
* [python-registry](https://github.com/williballenthin/python-registry) Apache License 2.0


Dieses Programm liest die Laufzeiten eines Windows-Computers aus den Windows eventlog Dateien (C:\Windows\System32\winevt\Logs\System.evtx) aus. Dieses Daten kann man mit den Energiespar-Einstellungen des Computers anreichern, wenn man auch die Registry-Datei (C:\Windows\System32\config\System) angibt. Dies muss man machen bevor man "Select System.evtx" drückt. Das Anreichern nützt aber oft nichts, da die Einstellungen zum Teil oft geändert werden und nur die aktuellsten Einstellungne ausgelesen werden können.

Das Programm gibt eine Liste der Laufzeiten des Computers direkt im Fenster (GUI) aus, wie auch als Excel Datei im Ordner im dem das Programm ausgeführt wurde.

Zu Debug Zwecken kann man das Programm auch per Konsole mit "EVTX_Analyser.exe > debug.txt" ausführen und erhält den Konsolen-Output in der debug.txt Datei.
