# Calendar Extractor

A script to extract calendar data from an Outlook for Mac archive and format it for use in an Excel Sheet. 

## Setup

1. Install dependencies using the `init.sh` script by running `./init.sh` in the terminal. 

2. Create a .env file in the root directory with the following variables. See .env.example for an example:

- `YOUR_EMAIL`: Your email address
- `IGNORE_PHRASES`: A comma-separated list of phrases to ignore in the calendar events
- `CONTACTS_FILE`: The name of the contacts file to use. This file should be in the `data` directory.
- `SAVE_FILE`: The name of the file to save the calendar data to. This file will be saved/read from the `data` directory.

3. Add a `Contacts.xlsx` file in the root directory with the following columns:

- `Name`: The name of the contact
- `Affiliation`: The affiliation of the contact
- `Type`: The type of contact
- `Role`: The role of the contact 
- `Email`: The email address of the contact

4. Extract the calendar data from Outlook for Mac and save it in the `data/Outlook for Mac Archive` directory. It will be exported as an .olm file, extract it like a normal .zip file. 

5. Run the script by running `python CalendarExtract.py` in the terminal. 

6. The script will create an `Calendar.xlsx` file in the root directory. 