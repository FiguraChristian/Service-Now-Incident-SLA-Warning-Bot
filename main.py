import os
import requests
from dotenv import load_dotenv
import pandas as pd
from datetime import datetime, timedelta
from SLAmapping import sla_table # own SLA mapping table
import win32com.client as win32

# looking for local .env file
load_dotenv()

# get information from .env
url = os.getenv("SN_URL")
username = os.getenv("SN_UN")
password = os.getenv("SN_PW")
mail_adress = os.getenv("Mail")
reporting_path=os.getenv("REPORT_PATH")

# URL INC Table 
incident_url = "/api/now/table/incident"

# create target path
target_path = f"{url}{incident_url}"


# params for clean names instead of numbers
params = {
    "sysparm_display_value": "true",
}

# send and request JSON
headers = {
    "Content-Type":"application/json",
    "Accept":"application/json"
}

print("Loading data from ServiceNow...")


# function mail send - python controlling outlook
def send_warning_mail(ticket_nr, hours_left):
    try:
        # using outlook to create a new mail
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        
        # Set Recipient
        mail.To = mail_adress
        
        # Set subject
        mail.Subject = f"⚠️ SLA WARNING: Ticket {ticket_nr}"
        mail.Body = f"Warning, the ticket {ticket_nr} will breach its SLA in {hours_left} hours !"
        
        mail.Send()
        print(f"-> Mail for {ticket_nr} sent.")
    except Exception as e:
        print(f"Cannot send Mail: {e}")




# Send Requests all INC
inc_response = requests.get(target_path, auth=(username,password), headers=headers,params=params)


# create dictionary for data handling
inc_data = inc_response.json()


# extract list of data from variable data
all_incidents = inc_data["result"]

# data container 
inc_my_data = []


# filter for incident
for element in all_incidents:

    sla_deadline = "Error/No Date"
    hours_left = 0
    allowed_hours = 48 # Standard-Fallback

    # get prio and open_time and convert
    prio_raw = element.get("priority")
    opentime_raw = element.get("opened_at")

    # calculate deadline
    if opentime_raw:

        if prio_raw:
            prio_clear = prio_raw[0]
            # get hours from table
            allowed_hours = sla_table.get(prio_clear, 48)

        try:
            # convert text
            start_date = datetime.strptime(opentime_raw, "%Y-%m-%d %H:%M:%S")

            #correct time zone 
            start_date_german = start_date + timedelta(hours=9)

            # calc deadline
            deadline_obj = start_date_german + timedelta(hours=allowed_hours)  

            # calc left time
            time_left = deadline_obj - datetime.now()
            
            # formatting data
            sla_deadline = deadline_obj.strftime("%Y-%m-%d %H:%M")
            hours_left = round(time_left.total_seconds() / 3600, 1)

            if 0 < hours_left < 4:
                print(f"ALARM: {element.get('number')} is critical ({hours_left}h)!")
                send_warning_mail(element.get('number'), hours_left)

        except ValueError:
            print(f"Datumsfehler bei Ticket {element.get('number')}")
    


    # create dictionary for excel output
    inc_excel_dict = {
            "INC-Number": element.get("number"),
            "Description": element.get("short_description"),
            "Open Time": element.get("opened_at"),
            "Caller / Opended By": element.get("caller_id", {}).get("display_value", "Unbekannt"),
            "Last Update: " : element.get("sys_updated_on"),
            "Priority": element.get("priority"),
            "SLA Deadline": sla_deadline, 
            "Hours Left": hours_left,
            "SLA Hours Total": allowed_hours 
            }
    

    inc_my_data.append(inc_excel_dict)

# Create DataFrame with Pandas and copy to Excel
df = pd.DataFrame(inc_my_data)
df.to_excel(reporting_path,index=False)