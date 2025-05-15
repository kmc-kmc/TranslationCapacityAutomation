📝 Translation Capacity using Google Calendar + Google Sheets

This project is a Google Apps Script that integrates with Google Calendar and Google Sheets to help translators automatically estimate their available translation capacity over the next two weeks, based on scheduled events.

Perfect for freelancers or remote translators who want to keep project managers informed—especially those in different time zones—this tool provides a clear visual forecast of daily availability.

What It Does
Fetches events from selected Google Calendars (e.g., work, personal, travel)
Calculates how many hours are booked each day
Deducts productivity "penalty" based on busy hours
Calculates remaining word count capacity (default: 5600 words/day)

Updates a Google Sheet with:
Date, Remaining capacity, Status (Available / Unavailable)
Shades days marked as “Unavailable” for quick visual reference

📋 Example Output
https://docs.google.com/spreadsheets/d/11787eRsj2WTeUJI3mtH8SuB3G_j3U2k2YvZ1_ZGmy1Q/edit?gid=0#gid=0

🛠 How It Works
Open your Google Sheet.
Go to Extensions → Apps Script, and paste the updateCapacityFromCalendar() function.
Modify the includedCalendars list to match your calendar names.

Set default values:
defaultCapacity = 5600 (words per day)
penaltyPerHour = 700 (words lost per calendar hour)
Run the script or set a time-driven trigger to run it daily.

🧩 Customization
Add or remove calendar names from includedCalendars
Adjust the defaultCapacity and penaltyPerHour values
Extend the forecast to more than 14 days
Send summary via email or Slack using MailApp or webhooks

✅ Requirements
Google account with Calendar and Sheets access
Calendar events must be clearly defined with start and end times
Your sheet should have an empty range starting from A2:C to avoid overwriting important data

🚀 Running the Script
Added a Trigger to run automatically every day at 1900 To run manually

🛠 Future improvement
1. To fix the issue when there are events overlapping for 2 days which might happen when taking flight and change it timezone, the calculation of the translation will be incorrect.
2. This is only made for situation which the translator only have one project manager, more information might needed (e.g. a column for project manager to put how many words alreay taken) when there are more than one project manager to avoid overbooking.
