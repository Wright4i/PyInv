"""Invoice Processor Export Module:
    Exporting calendar events from Google Calendar (GCal)
"""

import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from dateutil.parser import parse as dtparse
from export_modules.util import ExportUtil

class GCal:
    """Google Calendar export class
    """
    def __init__(self):
        self.util = ExportUtil()
        self._dates = self.util.set_dates()
        self._db_name = "gcal.db"
        if not os.path.isfile(self._db_name):
            open(self._db_name, "w", encoding="utf-8").close()

        self.util.connect_db(self._db_name)
        self._credentials = self.get_credentials("gauth-credentials.json")

    def __del__(self):
        self.util.disconnect_db()

    def export(self, _date):
        """Export GCal events

        Args:
            _date (date): run date
        """
        self._dates = self.util.set_dates(_date)
        if self._credentials is not None:
            try:
                service = build("calendar", "v3", credentials=self._credentials)
                self.save_db(service)
            except HttpError as error:
                print(f"An error occurred: {error}")
        else:
            print("No valid credentials.json file found.")

    @staticmethod
    def get_credentials(_secret):
        """Use Google Calendar API to get Credentials OAuth 2.0 client secret

        Args:
            _secret (str): path to google secret.json

        Returns:
            object Credentials: Google Credentials
        """
        scopes = ["https://www.googleapis.com/auth/calendar.readonly"]

        creds = None
        token = "token.json"

        if os.path.exists(token):
            creds = Credentials.from_authorized_user_file(token, scopes)

        if os.path.exists(_secret):
             # If there are no (valid) credentials available, log in.
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        _secret, scopes)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open(token, "w", encoding="utf-8") as token:
                    token.write(creds.to_json())

            return creds
        else:
            print("Missing gauth-credentials.json")

    def save_db(self, _service):
        """For each calendar id call Google Calendar Events API and save to sqlite db table calendar

        Args:
            _service (object build): Google Calendar build object
        """
        self.util.db_cur.execute(
            """CREATE TABLE IF NOT EXISTS calendar (
                calendar TEXT,
                title REAL,
                start TEXT,
                end TEXT,
                duration REAL
            )"""
        )

        # Create index over timesheet dates if it doesn't exist
        self.util.db_cur.execute(
            """CREATE INDEX IF NOT EXISTS calendar_date_index ON calendar (start)"""
        )

        self.util.db_cur.execute(
            "DELETE FROM calendar WHERE start BETWEEN ? and ?",
            (self._dates["fom_isoz"], self._dates["eom_isoz"])
        )

        # Get calendar ids
        calendars = []
        calendar_list = _service.calendarList().list().execute()
        for calendar_list_entry in calendar_list["items"]:
            calendars.append(calendar_list_entry["id"])

        if calendars is None:
            print("Did not find any calendars. Trying 'primary'...")
            calendars = ["primary"]

        for cal_id in calendars:

            # Call the Calendar API
            events_result = _service.events().list(
                calendarId=cal_id,
                timeMin=self._dates["fom_isoz"],
                singleEvents=True,
                orderBy="startTime"
            ).execute()
            events = events_result.get("items", [])

            if not events:
                print("No events found for calendar: " + cal_id + ".")
            else:
                # Insert all the events for last month
                for event in events:
                    start = event["start"].get("dateTime", event["start"].get("date"))
                    end = event["end"].get("dateTime", event["end"].get("date"))
                    summary = "(No title)"
                    if "summary" in event:
                        summary = event["summary"]
                    time_diff = dtparse(end) - dtparse(start)

                    # in fractional hours
                    duration = int(round(time_diff.total_seconds() / 60))

                    # Insert events that fall within date range
                    if self._dates["fom_isoz"] <= start <= self._dates["eom_isoz"]:
                        self.util.db_cur.execute(
                            "INSERT INTO calendar VALUES (?, ?, ?, ?, ?)",
                            (cal_id, summary, start, end, duration)
                        )

        self.util.commit_db()
