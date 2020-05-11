Dieses Programm liest Windows Event Log IDs aus der Windows Event Log System.evtx Datei aus und stellt diese in einer Liste (nach Datum und Uhrzeit sortiert) dar.

Standardmässig zeigt es alle Events an, die mit den Laufzeiten des PCs zu tun haben, also Start, Stop, Standby und Aufwachen. Man kann aber auch eigene Event IDs angeben. Dazu einfach im Textfeld oben links die Event ID (Zahl), ein Gleichheitszeichen (=), und ein Kurzbeschrieb was der Event bedeutet, eigeben. (keine Leerzeichen vor und nach dem Gleich)

Beispiel:
6008=Unerwartet Heruntergefahren

Die standard Daten kann man mit den Energiespar-Einstellungen des Computers anreichern, wenn man auch die Registry-Datei (C:\Windows\System32\config\System) angibt. Dies muss man machen bevor man "Select System.evtx" drückt. Das Anreichern nützt aber oft nichts, da die Einstellungen zum Teil oft geändert werden und nur die aktuellsten Einstellungne ausgelesen werden können.

Das Programm gibt eine Liste der gewünschten Events direkt im Fenster (GUI) aus, wie auch als Excel Datei (CSV) im Ordner im dem das Programm ausgeführt wurde.

Zu Debug Zwecken kann man das Programm auch per Konsole mit "EVTX_Analyser.exe > debug.txt" ausführen und erhält den Konsolen-Output in der debug.txt Datei.