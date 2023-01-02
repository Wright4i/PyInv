"""Invoice Processor Export Module:
    Exporting timesheets from Microsoft Project Portfolio Management (PPM)
"""

from datetime import datetime, timedelta
from pathlib import Path
import glob
import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from export_modules.util import ExportUtil

class PPM:
    """ Microsoft Project Portfolio Management export class
    """
    def __init__(self):
        self.util = ExportUtil()
        self._path_downloads = str(Path.home() / "Downloads")
        self._file_name = "My+Timesheet*.xlsx"
        self._db_name = "ppm.db"
        if not os.path.isfile(self._db_name):
            open(self._db_name, "w", encoding="utf-8").close()

        self.util.connect_db(self._db_name)
        self._credentials = self.get_credentials()
        self._dates = self.util.set_dates()

    def __del__(self):
        self.util.disconnect_db()

    def export(self, _date):
        """Export PPM timesheets

        Args:
            _date (date): run date
        """
        self._dates = self.util.set_dates(_date)
        if self._credentials is not None:
            self.cleanup_downloads(self._path_downloads, self._file_name)
            self.selenium_run(
                self._credentials["url"],
                "tsDate"
            )
            self.save_db(self._db_name, self._path_downloads, self._file_name)
        else:
            print("No ppm.db credentials found")


    @staticmethod
    def cleanup_downloads(_path_downloads, _file_name):
        """Removes files in your downloads folder that match the PPM file name

        Args:
            _path_downloads (str): path to Downloads folder
            _file_name (str): file name to match
        """
        all_files = glob.glob(os.path.join(_path_downloads, _file_name))
        for file in all_files:
            os.remove(file)

    def get_credentials(self):
        """Get Credentials from your sqlite db credentials table

        Returns:
            dict: username and password
        """

        self.util.db_cur.execute(
            "CREATE TABLE IF NOT EXISTS credentials (username TEXT, password TEXT, url TEXT)"
        )
        self.util.db_cur.execute("SELECT username, password, url FROM credentials")
        creds = self.util.db_cur.fetchone()

        if not creds:
            creds = ["", "", ""]
            creds[0] = input("Enter PPM User: ")
            creds[1] = input("Enter PPM Password: ")
            creds[2] = input("Enter PPM URL: ")
            if input("Save password? (Y/N): ") == "Y":
                self.util.db_cur.execute(
                    "INSERT INTO credentials VALUES (?, ?, ?)",
                    (creds[0], creds[1], creds[2])
                )

        if creds:
            return {
                "user":creds[0],
                "pass":creds[1],
                "url":creds[2]
            }

    def selenium_run(self, _url, _page_parm):
        """Selenium - Run browser automation

        Args:
            _url (str): PPM url
            _page_parm (str): PPM url date parm (typically tsDate)
        """
        # Define chrome options
        options = webdriver.ChromeOptions()

        # Hack to disable logging about suspended USB devices https://stackoverflow.com/a/70476264
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        # Open chrome
        selenium_driver = webdriver.Chrome(
            options=options,
            service=Service(ChromeDriverManager().install())
        )

        # Open Timesheet
        selenium_driver.get(_url)

        # Define wait function
        selenium_wait = WebDriverWait(selenium_driver, 10)

        # Login
        self.selenium_login(selenium_driver, selenium_wait)

        # Export 6 weeks
        for week_num in range(6):
            self.selenium_export_page(
                selenium_driver,
                selenium_wait,
                self._dates["fom"] + timedelta(days=(week_num*7)),
                _url + "?" + _page_parm
            )

        # Close
        selenium_driver.close()

    def selenium_login(self, _selenium_driver, _selenium_wait):
        """Selenium - Handle logging into PPM using your credentials
        """
        # Sign in page
        if _selenium_driver.title == "Sign in to your account":
            # User name prompt
            login_user = _selenium_wait.until(lambda d: d.find_element(By.ID, "i0116"))
            login_user.send_keys(self._credentials["user"])
            login_user.send_keys(Keys.RETURN)
            time.sleep(1)

            # Password prompt
            login_pass = _selenium_wait.until(lambda d: d.find_element(By.ID, "i0118"))
            login_pass.send_keys(self._credentials["pass"])
            time.sleep(1)

            # Sign in Button
            login_button = _selenium_wait.until(
                EC.element_to_be_clickable((By.ID, "idSIButton9"))
            )
            login_button.click()
            time.sleep(1)

            # Stay logged in?
            login_button = _selenium_wait.until(
                EC.element_to_be_clickable((By.ID, "idSIButton9"))
            )
            login_button.click()
            time.sleep(1)

    def selenium_export_page(self, _selenium_driver, _selenium_wait, _week, _url):
        """Selenium - Change weeks on timesheet and export to xlsx

        Args:
            _week (datetime): a date in the week to navigate to
            _url (str): PPM url
        """
        # Change pages
        _selenium_driver.get(f"{_url}={_week.strftime('%#m/%#d/%Y')}")
        time.sleep(1)

        # Click the options ribbon
        options_ribbon = _selenium_wait.until(EC.element_to_be_clickable(
            (By.ID, "Ribbon.ContextualTabs.TiedMode.Options-title"))
        )
        options_ribbon.click()
        time.sleep(1)

        # Click the export excel button
        options_ribbon_export = _selenium_wait.until(EC.element_to_be_clickable(
            (By.ID, "Ribbon.ContextualTabs.TiedMode.Options.Share.ExportExcel-Large"))
        )
        options_ribbon_export.click()
        time.sleep(3)

    def save_db(self, _db_name, _path_downloads, _file_name):
        """Save exported xlsx files to sqlite db table timesheet

        Args:
            _db_name (str): sqlite database name
            _path_downloads (str): path to Downloads folder
            _file_name (str): file name to match
        """
        self.util.connect_db(_db_name)

        # Load files into database
        self.util.db_cur.execute(
            """CREATE TABLE IF NOT EXISTS timesheet (
                date TEXT,
                hours REAL,
                description TEXT,
                project TEXT
            )"""
        )

        # Create index over timesheet dates if it doesn't exist
        self.util.db_cur.execute(
            """CREATE INDEX IF NOT EXISTS timesheet_date_index ON timesheet (date)"""
        )

        self.util.db_cur.execute(
            "DELETE FROM timesheet WHERE date BETWEEN ? and ?",
            (self._dates["fom"], self._dates["eom"])
        )

        all_files = glob.glob(os.path.join(_path_downloads, _file_name))
        for file in all_files:
            df_ppm = pd.read_excel(file)
            df_ppm = df_ppm.drop(columns=[
                "Unnamed: 0",
                "Process Status",
                "WBS",
                "Work",
                "Remaining Work",
                "Start",
                "Finish",
                "% Work Complete",
                "Time Type"
            ])
            df_ppm = df_ppm.fillna("0")

            for day in df_ppm.columns[2:]:
                # Get date
                ppm_date = datetime.strptime(
                    day[4:] + "/" + self._dates["fom"].strftime("%Y"), "%m/%d/%Y"
                )

                for _index, row in df_ppm.iterrows():
                    # Skip rows not worked
                    if row[day] == "0":
                        continue

                    # Skip dates outside range
                    if ppm_date < self._dates["fom"] or ppm_date > self._dates["eom"]:
                        continue

                    # Duration
                    hours = float(row[day].replace("h",""))

                    # Customer
                    project = row["Project Name"]

                    # Description
                    description = row["Task Name/Description"]

                    # Format Total Row
                    if row["Project Name"] == "Total work":
                        project = "*NOTE*"
                        description = f"PPM TOTAL HOURS: {row[day].replace('h','')}"
                        hours = 0

                    # SQL
                    self.util.db_cur.execute(
                        "INSERT INTO timesheet VALUES (?, ?, ?, ?)",
                        (ppm_date, hours, description, project)
                    )

            os.remove(file)

        self.util.commit_db()
