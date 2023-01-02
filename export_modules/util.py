"""Invoice Processor Utility Module
"""
from datetime import date, datetime
import calendar
import sqlite3

class ExportUtil:
    """Export Util Class:
        Contains common methods for all export modules
        - Date calculations for first/end of month
        - Database connection handling
    """
    def __init__(self):
        self.db_conn = None
        self.db_cur = None
        self.set_dates(date.today())

    def set_dates(self, _date=date.today()):
        """Sets first/end of month using date passed in

        Args:
            _date (date): date to use as start

        Returns:
            dict: all calculated and formatted dates from get_dates()
        """
        # End of month
        date_eom = datetime.combine(
            _date.replace(day=calendar.monthrange(_date.year, _date.month)[1]),
            datetime.max.time()
        )
        date_eom_isoz = date_eom.isoformat() + "Z"

        # First of month
        date_fom = datetime.combine(
            date_eom.replace(day=1),
            datetime.min.time()
        )
        date_fom_isoz = date_fom.isoformat() + "Z"

        self._dates = {
            "eom": date_eom,
            "eom_isoz": date_eom_isoz,
            "fom": date_fom,
            "fom_isoz": date_fom_isoz
        }

        return self.get_dates()

    def get_dates(self):
        """Gets dates previously set by set_dates()

        Returns:
            dict: all calculated and formatted dates
        """
        return self._dates

    def connect_db(self, _db):
        """Connect to Database

        Args:
            _db (str): sqlite database name
        """
        if hasattr(self, "db_conn"):
            self.db_conn = sqlite3.connect(_db)

        if hasattr(self, "db_cur"):
            self.db_cur = self.db_conn.cursor()

    def commit_db(self):
        """Commit changes to Database
        """
        if hasattr(self, "db_conn"):
            if self.db_conn is not None:
                self.db_conn.commit()

    def disconnect_db(self):
        """Disconnect from Database
        """
        if hasattr(self, "db_conn"):
            self.commit_db()
            if self.db_conn is not None:
                self.db_conn.close()
            del self.db_conn

        if hasattr(self, "db_cur"):
            del self.db_cur
