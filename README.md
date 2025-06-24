# ebay_scraper
Ebay Scraper für Sammelkarten

# Verison: 1.0 
# Stannd: 22.06.2025


# Beschreibung: 
Durchsucht eBay nach den letzten Verkäufen und speichert die Ausgaben in einzelnen Excel-Dateien im (ggf.)
neu zu erstellenden Ordner "Preisverläufe". Die Ergebnisse werden überdies im Terminal ausgegeben. Es wird
berücksichtigt, dass die Daten nicht doppelt erhoben werden.

Das Datum ist im Schema JJJJ.MM.TT aufgebaut. Die Tabellen werden mit jedem Durchlauf nach Datum aufsteigend sortiert
und der Durchschnittspreis sowie der Höchstverkaufspreis wird gespeichert.
Die URLs können entweder in Zeile 249 manuell kopiert werden, oder es wird nur der Name angepasst. Die Sucheinstellungen
betreffen Verkäufe in der Europäischen Union. Mehrere URLs müssen in der Liste durch ein Komma getrennt werden.
Die Excel-Listen können abschließend um Fehlerfassungen bereinigt werden. Dazu kann die URL direkt aus der Excel-Datei
geöffnet werden.

Ebay speichert die Verkaufsdaten der letzten 3 Monate. Das Script sollte daher bestenfalls alle 3 Monate ausgeführt
werden, damit keine Fehlerfassung nach der manuellen Löschung erneut erfasst werden.
