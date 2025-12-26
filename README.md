# ServiceNow INC SLA Warning Bot
Ein automatisiertes Überwachungs-Tool ("Watchdog"), das ServiceNow-Incidents abruft, SLA-Fristen basierend auf Prioritäten in Echtzeit berechnet und proaktive Warn-E-Mails via Outlook versendet.

# Hauptfunktionen
- API-Integration: Verbindet sich mit der ServiceNow Table REST API, um alle offenen Incidents abzurufen
- Dynamische SLA-Berechnung: Mappt Prioritäten auf spezifische Stundenlimits (definiert in der externen Datei SLAmapping.py)
- Zeitzonen-Korrektur: Rechnet die ServiceNow-Systemzeit in die lokale Zeit um
- Excel-Reporting: Generiert einen detaillierten Incident Report (.xlsx)
- Outlook-Integration: Nutzt Pywin32, um den lokalen Outlook-Client (win32) für den Mailversand zu steuern


# Tech Stack
- Python 3.13
- Requests (REST API Handling)
- Pandas (Datenbereinigung, Berechnung & Excel-Export)
- PyWin32 (Lokale Outlook-Automatisierung)
- Python-Dotenv (Sicherheit & Credential-Management)
- nutzt Windows Task Scheduler zur geplanten Ausführung des Python Skripts


