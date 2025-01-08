#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import zipfile
import os
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Basic info
your_email=os.getenv("YOUR_EMAIL")

# Extracted from Outlook path
base_directory = f"./data/Outlook for Mac Archive/Accounts/{your_email}"

# Extract .olm file if it exists
olm_files = [f for f in os.listdir("data") if f.endswith(".olm")]
print(f"Extracting .olm file... {olm_files}")
if olm_files:
    
    olm_path = os.path.join("data", olm_files[0])
    extract_path = os.path.join("data", "Outlook for Mac Archive")
    
    # Remove existing extracted folder if it exists
    if os.path.exists(extract_path):
        import shutil
        shutil.rmtree(extract_path)
        
    # Extract the .olm file
    with zipfile.ZipFile(olm_path, 'r') as zip_ref:
        zip_ref.extractall("data/Outlook for Mac Archive")

# Ask user for date start, set end range to one week later.
start_date = input("Enter start date (YYYY-MM-DD): ")

# Convert string to datetime before adding timedelta
start_date = datetime.strptime(start_date, '%Y-%m-%d')  # Adjust format string to match your date format
end_date = start_date + timedelta(days=7)

print(f"Extracting calendar data from {start_date} to {end_date}...")

# Format file paths for data directory if not specified. 
contacts_file = os.getenv("CONTACTS_FILE")
if "/" not in contacts_file:
    contacts_file = f"./data/{contacts_file}"

save_file = os.getenv("SAVE_FILE")
if "/" not in save_file:
    save_file = f"./data/{save_file}"

# Read contacts file and create email mapping
print(f"Reading contacts file... {contacts_file}")
contacts_df = pd.read_excel(contacts_file, names=['Name', 'Affiliation', 'Type', 'Role', 'Email'])
email_mapping = {
    row['Email'].lower(): f"{row['Name']}, {row['Role']}, {row['Affiliation']}"
    for _, row in contacts_df.iterrows()
    if pd.notna(row['Email'])
}

def format_participants(participants_str):
    if not participants_str:
        return None
    
    participants = participants_str.split(', ')
    formatted_participants = []
    
    for participant in participants:
        participant_lower = participant.lower()
        # Skip CNAS email addresses
        if participant_lower.endswith(os.getenv("EMAIL_DOMAIN").lower()):
            continue
        if participant_lower in email_mapping:
            formatted_participants.append(email_mapping[participant_lower])
        else:
            formatted_participants.append(participant)
    
    return ', '.join(formatted_participants) if formatted_participants else None

# Extract appointments from the XML file
def extract_appointments(file_path, member_name):
    print(f"Extracting appointments from {file_path}...")
    tree = ET.parse(file_path)
    root = tree.getroot()
    
    appointments = []
    
    for appointment in root.iter('appointment'):
        title = appointment.find('OPFCalendarEventCopySummary')
        start_time = appointment.find('OPFCalendarEventCopyStartTime')
        mod_date = appointment.find('OPFCalendarEventCopyModDate')
        
        if title is None or start_time is None or mod_date is None:
            continue
            
        event_date = datetime.strptime(start_time.text, '%Y-%m-%dT%H:%M:%S')
        event_title = title.text
            
        # Skip excluded titles
        if event_title.lower() in ["private event", "appointment", "new event"]:
            continue
            
        excluded_keywords = os.getenv("IGNORE_PHRASES").split(", ")
        
        # Fix keyword filtering by removing extra spaces and using direct string containment
        event_title_lower = event_title.lower()
        if any(keyword in event_title_lower for keyword in excluded_keywords):
            continue
        
        to_append = {
            'Date': event_date.strftime('%Y-%m-%d'),
            'Title': event_title,
            'Member': member_name if member_name != "Calendar" else your_email,
            'ModificationDate': datetime.strptime(mod_date.text, '%Y-%m-%dT%H:%M:%S'),
            'EventDate': event_date
        }
        
        location = appointment.find('OPFCalendarEventCopyLocation')
        to_append['Location'] = location.text if location is not None else None
        
        participants = []
        attendee_list = appointment.find('OPFCalendarEventCopyAttendeeList')
        if attendee_list is not None:
            for attendee in attendee_list.findall('appointmentAttendee'):
                address = attendee.attrib.get('OPFCalendarAttendeeAddress', None)
                if address:
                    participants.append(address)
        
        participants_str = ', '.join(sorted(set(participants))) if participants else None
        to_append['Participants'] = format_participants(participants_str)
        to_append['Topic'] = event_title
        
        details = appointment.find('OPFCalendarEventCopyDescription')
        if details is not None:
            soup = BeautifulSoup(details.text, 'html.parser')
            to_append['Details'] = soup.get_text()
        else:
            to_append['Details'] = None
        
        appointments.append(to_append)
    
    return appointments

# Main processing
all_appointments = []

for folder_name in os.listdir(base_directory):
    folder_path = os.path.join(base_directory, folder_name)
    if os.path.isdir(folder_path):
        xml_file_path = os.path.join(folder_path, 'Calendar.xml')
        if os.path.exists(xml_file_path):
            member_appointments = extract_appointments(xml_file_path, folder_name)
            all_appointments.extend(member_appointments)

# Create DataFrame
df = pd.DataFrame(all_appointments)

if not df.empty:
    # Convert EventDate to datetime if it isn't already
    df['EventDate'] = pd.to_datetime(df['EventDate'])
    
    # First filter by date range
    df = df[df['EventDate'].dt.date.between(start_date.date(), end_date.date())]
    
    # Create a simplified identifier for each unique event
    df['EventKey'] = df['EventDate'].dt.strftime('%Y-%m-%d %H:%M') + '_' + df['Title'].str.lower()
    
    # Combine members for the same event while keeping other details
    df_combined = df.groupby('EventKey').agg({
        'Date': 'first',
        'Title': 'first',
        'Member': lambda x: ', '.join(sorted(set(x))),
        'Location': 'first',
        'Participants': 'first',
        'Topic': 'first',
        'Details': 'first',
        'ModificationDate': 'max'
    }).reset_index()
    
    # Create base title for handling canceled events
    df_combined['BaseTitle'] = df_combined['Title'].str.lower().str.replace('canceled:', '', regex=False).str.replace('hold:', '', regex=False).str.strip()
    df_combined['IsCanceled'] = df_combined['Title'].str.lower().str.startswith('canceled:')
    
    # For each base title, keep non-canceled version if it exists
    df_final = df_combined.sort_values('ModificationDate').groupby('BaseTitle').apply(
        lambda x: x[~x['IsCanceled']].iloc[-1] if len(x[~x['IsCanceled']]) > 0 else x.iloc[-1]
    ).reset_index(drop=True)
    
    # Drop working columns
    df_final = df_final.drop(['EventKey', 'BaseTitle', 'IsCanceled', 'ModificationDate'], axis=1)
else:
    df_final = df

# Write to Excel with the date as the sheet name
sheet_name = f"raw_{end_date.strftime('%Y%m%d')}"

if not os.path.exists(save_file):
    with pd.ExcelWriter(save_file, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name=sheet_name, index=False)
else:
    with pd.ExcelWriter(save_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_final.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Data successfully written to sheet '{sheet_name}' in '{save_file}'.")



