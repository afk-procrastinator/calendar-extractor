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
from colorama import init, Fore, Back, Style

# Initialize colorama
init(autoreset=True)

# Load environment variables
load_dotenv()

# Basic info
your_email=os.getenv("YOUR_EMAIL")
ignore_accounts = os.getenv("IGNORE_ACCOUNTS").split(", ")

# Extracted from Outlook path
base_directory = f"./data/Outlook for Mac Archive/Accounts/{your_email}"

# Extract .olm file if it exists
olm_files = [f for f in os.listdir("data") if f.endswith(".olm")]
print(f"{Fore.CYAN}Extracting .olm file...{Style.RESET_ALL} {Fore.YELLOW}{olm_files}{Style.RESET_ALL}")
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

# Replace the input prompt with a hardcoded date
start_date = input(f"{Fore.GREEN}Enter start date (YYYY-MM-DD): {Style.RESET_ALL}")
# start_date = "2025-01-01"

# Convert string to datetime before adding timedelta
start_date = datetime.strptime(start_date, '%Y-%m-%d')
end_date = start_date + timedelta(days=6)

# Pretty print start date for context
print(f"{Fore.CYAN}Extracting calendar data from {Fore.YELLOW}{start_date.strftime('%B %d')}{Fore.CYAN} to {Fore.YELLOW}{end_date.strftime('%B %d')}{Fore.CYAN}...{Style.RESET_ALL}")

# Format file paths for data directory if not specified. 
contacts_file = os.getenv("CONTACTS_FILE")
if "/" not in contacts_file:
    contacts_file = f"./data/{contacts_file}"

save_file = os.getenv("SAVE_FILE")
if "/" not in save_file:
    save_file = f"./data/{save_file}"

# Read contacts file and create email mapping
print(f"{Fore.CYAN}Reading contacts file...{Style.RESET_ALL} {Fore.YELLOW}{contacts_file}{Style.RESET_ALL}")
contacts_df = pd.read_excel(contacts_file, names=['Name', 'Affiliation', 'Type', 'Role', 'Email'])

# Get your name from .env file or use email username as fallback
your_name = os.getenv("YOUR_NAME")

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
        # Replace your email with your name
        if participant_lower == your_email.lower():
            formatted_participants.append(your_name)
        # Skip other custom email addresses
        elif participant_lower.endswith(os.getenv("EMAIL_DOMAIN").lower()):
            continue
        elif participant_lower in email_mapping:
            formatted_participants.append(email_mapping[participant_lower])
        else:
            formatted_participants.append(participant)
    
    return ', '.join(formatted_participants) if formatted_participants else None

# Extract appointments from the XML file
def extract_appointments(file_path, member_name):
    print(f"{Fore.CYAN}Extracting appointments for {Fore.YELLOW}{member_name}{Fore.CYAN}...{Style.RESET_ALL}")
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
            'Date': event_date.strftime('%m/%d/%y'),
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
            # Remove all <br> tags
            for br in soup.find_all('br'):
                br.decompose()
            to_append['Details'] = soup.get_text()
        else:
            to_append['Details'] = None
        
        appointments.append(to_append)
    
    return appointments

# Main processing
all_appointments = []

for folder_name in os.listdir(base_directory):
    # Skip folders in ignore_accounts
    if folder_name in ignore_accounts:
        continue
        
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
    
    # Replace your email with your name in the Member column
    df['Member'] = df['Member'].apply(lambda x: your_name if x == your_email else x)
    
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
    # Fixed deprecation warning by explicitly selecting columns after groupby
    df_final = (df_combined.sort_values('ModificationDate')
                .groupby('BaseTitle', group_keys=False)
                [['Date', 'Title', 'Member', 'Location', 'Participants', 'Topic', 'Details', 'IsCanceled', 'ModificationDate']]
                .apply(lambda x: x[~x['IsCanceled']].iloc[-1] if len(x[~x['IsCanceled']]) > 0 else x.iloc[-1])
                .reset_index(drop=True))
    
    # Reorder columns
    columns = ['Date', 'Title', 'Member', 'Location', 'Participants', 'Topic', 'Details']
    df_final = df_final[columns]
else:
    df_final = df

# Write to Excel with the date as the sheet name
sheet_name = f"raw_{end_date.strftime('%Y%m%d')}"

# Create Excel writer with standard date format
if not os.path.exists(save_file):
    with pd.ExcelWriter(save_file, engine='openpyxl') as writer:
        # Convert Date column to datetime with explicit format
        df_final['Date'] = pd.to_datetime(df_final['Date'], format='%m/%d/%y')
        df_final.to_excel(writer, sheet_name=sheet_name, index=False)
        # Apply Excel's built-in short date format
        worksheet = writer.sheets[sheet_name]
        for cell in worksheet['A']:
            if cell.row > 1:  # Skip header row
                cell.number_format = 'mm/dd/yy'
else:
    with pd.ExcelWriter(save_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        # Convert Date column to datetime with explicit format
        df_final['Date'] = pd.to_datetime(df_final['Date'], format='%m/%d/%y')
        df_final.to_excel(writer, sheet_name=sheet_name, index=False)
        # Apply Excel's built-in short date format
        worksheet = writer.sheets[sheet_name]
        for cell in worksheet['A']:
            if cell.row > 1:  # Skip header row
                cell.number_format = 'mm/dd/yy'

print(f"{Fore.GREEN}Data successfully written to sheet '{Fore.YELLOW}{sheet_name}{Fore.GREEN}' in '{Fore.YELLOW}{save_file}{Fore.GREEN}'.{Style.RESET_ALL}")



