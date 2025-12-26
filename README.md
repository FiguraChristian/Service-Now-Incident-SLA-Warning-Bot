# ServiceNow Incident SLA Warning Bot

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
- Windows Task Scheduler zur stündlichen Ausführung des Python Skripts


# Warning per Mail 

<img width="1642" height="417" alt="image" src="https://github.com/user-attachments/assets/52945f4b-56b2-4131-98a6-20a6d8f9af44" />

<img width="1754" height="669" alt="image" src="https://github.com/user-attachments/assets/0c5a820c-929a-4481-befd-be52534d8aa7" />



# Screenshots Incidents Service Now

<img width="1919" height="1077" alt="image" src="https://github.com/user-attachments/assets/c4b76a8d-d079-43e3-8ed9-c10818c7c9ad" />
<img width="1920" height="1002" alt="image" src="https://github.com/user-attachments/assets/32d71c50-fa9b-4f88-a30b-514f5d140266" />
<img width="1725" height="443" alt="image" src="https://github.com/user-attachments/assets/ce0ac11c-53af-4437-8019-597f13e2c737" />


# Excel Report

<img width="1913" height="1032" alt="image" src="https://github.com/user-attachments/assets/89670fe2-2049-4563-abb7-d0735ececc1d" />


# Windows Task Scheduler 

<img width="942" height="714" alt="image" src="https://github.com/user-attachments/assets/c89e8dac-a989-41bd-82b7-6aae6721f012" />
<img width="679" height="746" alt="image" src="https://github.com/user-attachments/assets/927153d7-86f5-40d5-a84f-af620f76bfb3" />

