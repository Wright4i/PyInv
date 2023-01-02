"""Invoice Processor Main Program
Automated time extraction from Microsoft PPM and Google Calendar.
"""

from datetime import datetime, timedelta
import csv
import os
from export_modules.gcal import GCal
from export_modules.ppm import PPM

# Get last month's date
last_month = datetime.now().replace(day=1) - timedelta(days=1)

# Prompts before running exports
run_date = input("Enter month you want to invoice (YYYY-MM): ") or last_month.strftime("%Y-%m")
export_gcal = input("Export Google Calendar? (Y/N): ") or "N"
export_ppm = input("Export PPM? (Y/N): ") or "N"

# Setup new export objects
gcal = GCal()
ppm = PPM()

# Run exports
run_date = datetime.strptime(run_date + "-01", "%Y-%m-%d")

if export_gcal == "Y":
    gcal.export(run_date)
    print("Exported Google Calendar")
else:
    gcal.util.set_dates(run_date)

if export_ppm == "Y":
    ppm.export(run_date)
    print("Exported PPM")
else:
    ppm.util.set_dates(run_date)

# Create xref and ignore tables if they don't exist
gcal.util.db_cur.execute(
    """CREATE TABLE IF NOT EXISTS project_xref (
        calendar TEXT,
        gcal_title REAL,
        inv_title REAL
    )"""
)
gcal.util.db_cur.execute(
    """CREATE TABLE IF NOT EXISTS ignore (
        calendar TEXT,
        title REAL,
        flag TEXT
    )"""
)

# Create xref and ignore tables if they don't exist
ppm.util.db_cur.execute(
    """CREATE TABLE IF NOT EXISTS project_xref (
        ppm_project TEXT,
        inv_project TEXT
    )"""
)
ppm.util.db_cur.execute(
    """CREATE TABLE IF NOT EXISTS ignore (
        project TEXT,
        description TEXT,
        flag TEXT
    )"""
)

# Get dates from GCal export module
gcal_dates = gcal.util.get_dates()

# Get dates from PPM export module
ppm_dates = ppm.util.get_dates()

# Loop through distinct titles in gcal.db where start between fom_isoz and eom_isoz
gcal.util.db_cur.execute(
    """SELECT DISTINCT calendar, title FROM calendar WHERE start BETWEEN ? and ?
        AND SUBSTR(start,11,1) = 'T'
    """,
    (gcal_dates["fom_isoz"], gcal_dates["eom_isoz"])
)
gcal_titles = gcal.util.db_cur.fetchall()

print("""
    Google Calendar Titles----------------

    When PPM has a project for meetings you should ignore the Google Calendar events that match.
    When PPM does not have a project for meetings you should enter the invoice project.

    All-day calendar events will be evenly divided between PPM projects when invoiced as *GCAL on the PPM project.
        Do NOT ignore these calendar events if you want to use this feature.
""")
for title in gcal_titles:
    print(f"Calendar:  {title[0]}")
    print(f"Title: {title[1]}")
    # Check if title is in ignore table
    gcal.util.db_cur.execute(
        "SELECT title, flag FROM ignore WHERE calendar = ? AND title = ?",
        (title[0], title[1])
    )
    ignore = gcal.util.db_cur.fetchone()
    if ignore is None:
        # Prompt user for ignore
        ignore_prompt = input(f"Ignore {title[1]}? (Y/N): ") or "N"

        # Add to ignore table
        gcal.util.db_cur.execute(
            "INSERT INTO ignore VALUES (?, ?, ?)",
            (title[0], title[1], ignore_prompt)
        )
        gcal.util.commit_db()
    else:
        print(f"Ignored (Y/N): {ignore[1]}")

    ignored = (
        (ignore is None and ignore_prompt == "Y") or
        (ignore is not None and ignore[1] == "Y")
    )

    # Check if title is in xref table
    if not ignored:
        gcal.util.db_cur.execute(
            "SELECT inv_title FROM project_xref WHERE calendar = ? AND gcal_title = ?",
            (title[0], title[1])
        )
        inv_title = gcal.util.db_cur.fetchone()
        if inv_title is None:
            # Prompt user for invoice project
            inv_title = input("Enter invoice project: ") or title[1]

            # Add to xref table
            gcal.util.db_cur.execute(
                "INSERT INTO project_xref VALUES (?, ?, ?)",
                (title[0], title[1], inv_title)
            )
            gcal.util.commit_db()
        else:
            print(f"Invoiced as: {inv_title[0]}")

# Loop through distinct projects in ppm.db where date between fom and eom
ppm.util.db_cur.execute(
    """SELECT DISTINCT
           project, description
       FROM timesheet
       WHERE date BETWEEN ? and ?
       AND INSTR(description, 'TOTAL HOURS: ') = 0
    """,
    (ppm_dates["fom"], ppm_dates["eom"])
)
ppm_projects = ppm.util.db_cur.fetchall()

print("""
    PPM Projects----------------

    PPM special feature:
    When prompted for invoice project use *GCAL to evenly divide hours
        for PPM projects using All-day Google Calendar events.
    """
)

for project in ppm_projects:
    print(f"Project:  {project[0]}")
    print(f"Description: {project[1]}")

     # Check if project is in ignore table
    ppm.util.db_cur.execute(
        "SELECT project, description, flag FROM ignore WHERE project = ? and description = ?",
        (project[0], project[1])
    )

    ignore = ppm.util.db_cur.fetchone()
    if ignore is None:
        # Prompt user for ignore
        ignore_prompt = input(f"Ignore {project[1]}? (Y/N): ") or "N"

        # Add to ignore table
        ppm.util.db_cur.execute(
            "INSERT INTO ignore VALUES (?, ?, ?)",
            (project[0], project[1], ignore_prompt)
        )
        ppm.util.commit_db()
    else:
        print(f"Ignored (Y/N): {ignore[2]}")

    ignored = (
        (ignore is None and ignore_prompt == "Y") or
        (ignore is not None and ignore[1] == "Y")
    )

    # Check if project is in xref table
    if not ignored:
        ppm.util.db_cur.execute(
            "SELECT inv_project FROM project_xref WHERE ppm_project = ?",
            (project[0],)
        )
        inv_project = ppm.util.db_cur.fetchone()
        if inv_project is None:
            # Prompt user for invoice project
            inv_project = input("Enter invoice project: ") or project[0]

            # Add to xref table
            ppm.util.db_cur.execute(
                "INSERT INTO project_xref VALUES (?, ?)",
                (project[0], inv_project)
            )
            ppm.util.commit_db()
        else:
            print(f"Invoiced as: {inv_project[0]}")

# Select all the calendar entries for run_date join to xref and ignore tables exclude any ignored
gcal.util.db_cur.execute(
    """SELECT
           TRIM(COALESCE(px.inv_title,ca.title)) as project,
           TRIM(ca.title) as notes,
           SUBSTR(ca.start,1,10) as date,
           case
                when SUBSTR(ca.start,11,1) = 'T' then
                    ROUND(duration/60,2)
                else
                    0
           end as hours,
           'gcal' as source
        FROM calendar ca
        LEFT OUTER JOIN project_xref px
        ON ca.calendar = px.calendar AND ca.title = px.gcal_title
        WHERE
        ca.start BETWEEN ? and ?
        AND (ca.calendar, ca.title) NOT IN (
            SELECT calendar, title FROM ignore WHERE flag = 'Y'
        )
        AND SUBSTR(ca.start,11,1) = 'T'
        """,
    (gcal_dates["fom_isoz"], gcal_dates["eom_isoz"])
)
gcal_calendar = gcal.util.db_cur.fetchall()

# Loop through hours and round up to the nearest 15 minutes saving to new list
gcal_calendar_rounded = []
for entry in gcal_calendar:
    if entry[3] % 0.25 != 0:
        entry = list(entry)
        print(f"Rounding up {entry[1]} on {entry[2]} from {entry[3]} to {round(entry[3] + 0.25 - (entry[3] % 0.25), 2)}")
        # Round up to nearest 15 minutes
        entry[3] = round(entry[3] + 0.25 - (entry[3] % 0.25), 2)
    # Save to new list
    gcal_calendar_rounded.append(entry)

# Set gcal_calendar to gcal_calendar_rounded
gcal_calendar = gcal_calendar_rounded

# Select all the timesheet entries for run_date join to xref and ignore tables exclude any ignored
ppm.util.db_cur.execute(
    """SELECT
            TRIM(COALESCE(px.inv_project, ts.project)) as project,
            TRIM(ts.description) as notes,
            SUBSTR(ts.date,1,10) as date,
            ts.hours as hours,
            'ppm' as source
        FROM timesheet ts
        LEFT OUTER JOIN project_xref px
        ON ts.project = px.ppm_project
        WHERE
        ts.date BETWEEN ? and ?
        AND (ts.project, ts.description) NOT IN (
            SELECT project, description FROM ignore WHERE flag = 'Y'
        )
        AND INSTR(ts.description, 'TOTAL HOURS: ') = 0
        AND COALESCE(px.inv_project, ts.project) <> '*GCAL'
    """,
    (ppm_dates["fom"], ppm_dates["eom"])
)
ppm_timesheet = ppm.util.db_cur.fetchall()

# Get worked hours
ppm.util.db_cur.execute(
    """ SELECT
            'WORKED HOURS' as project,
            '************' as notes,
            SUBSTR(ts.date,1,10) as date,
            SUBSTR(ts.description, INSTR(ts.description, 'TOTAL HOURS: ') + 13, 20) -
                COALESCE(ih.ignored_hours, 0) as worked_hours,
            '************' as source
        FROM timesheet ts
        LEFT OUTER JOIN (
            SELECT
                date,
                hours as ignored_hours
            FROM timesheet tx
            JOIN ignore ig
            ON tx.project = ig.project
            AND tx.description = ig.description
            AND ig.flag = 'Y') as ih
        ON ts.date = ih.date
        WHERE
        ts.date BETWEEN ? and ?
        AND INSTR(ts.description, 'TOTAL HOURS: ') > 0
    """,
    (ppm_dates["fom"], ppm_dates["eom"])
)
ppm_worked_hours = ppm.util.db_cur.fetchall()

# Get gcal splits
ppm.util.db_cur.execute(
    """ SELECT
            TRIM(COALESCE(px.inv_project, ts.project)) as project,
            TRIM(ts.description) as notes,
            SUBSTR(ts.date,1,10) as date,
            ts.hours as hours,
            'ppm' as source
        FROM timesheet ts
        LEFT OUTER JOIN project_xref px
        ON ts.project = px.ppm_project
        WHERE
        ts.date BETWEEN ? and ?
        AND (ts.project, ts.description) NOT IN (
            SELECT project, description FROM ignore WHERE flag = 'Y'
        )
        AND COALESCE(px.inv_project, ts.project) = '*GCAL'
    """,
    (ppm_dates["fom"], ppm_dates["eom"])
)
ppm_splits = ppm.util.db_cur.fetchall()
gcal_splits = []

# Loop through gcal_ppm_splits and find matching gcal entries
for entry in ppm_splits:
    gcal.util.db_cur.execute(
        """ SELECT
                COALESCE(px.inv_title, ca.title) as title,
                ca.title as notes
            FROM
                calendar ca
                LEFT OUTER JOIN project_xref px
                ON ca.calendar = px.calendar
                AND ca.title = px.gcal_title
            WHERE
                ? between ca.start and ca.end
                AND SUBSTR(ca.start,11,1) <> 'T'
                AND (ca.calendar, ca.title) NOT IN (
                    SELECT calendar, title FROM ignore WHERE flag = 'Y'
                )
        """,
        (entry[2],)
    )
    gcal_entries = gcal.util.db_cur.fetchall()

    split_hours = 0

    # Loop through gcal_entries and add to gcal_splits
    for gcal_entry in gcal_entries:
        # Divide hours from entry[3] by number of gcal_entries
        hours = entry[3] / len(gcal_entries)

        # Round hours to nearest 15 minutes
        if hours % 0.25 != 0:
            hours = round(hours + 0.25 - (hours % 0.25), 2)

        # Cap hours at ppm entry hours
        if split_hours + hours > entry[3]:
            hours = entry[3] - split_hours

        # Add hours to split_hours
        split_hours += hours

        gcal_splits.append([gcal_entry[0], gcal_entry[1], entry[2], hours, "*gcal"])

    if len(gcal_entries) == 0:
        print("No matching gcal entry for " + entry[1] + " on " + entry[2])

        # Add entry to gcal_splits
        gcal_splits.append([entry[0], entry[1], entry[2], entry[3], "ppm"])

# Combine the lists
detail = gcal_calendar + ppm_timesheet + ppm_worked_hours + gcal_splits

# Sort the detail list by date ascending, source ascending and project ascending
detail.sort(key=lambda x: (x[2], x[4], x[0]))

# Write detail list to text file detail.csv
with open("detail.csv", "w", encoding="UTF-8", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Project", "Notes", "Date", "Hours", "Source"])
    writer.writerows(detail)

# Create a dictionary to hold the summary
summary = {}

# Loop through the detail list and add to summary dictionary
for item in detail:
    # Skip worked hours
    if item[0] == "WORKED HOURS":
        continue
    if item[0] in summary:
        summary[item[0]] += item[3]
    else:
        summary[item[0]] = item[3]

# Convert summary dictionary to list
summary = list(summary.items())

# Sort the summary list by project ascending
summary.sort(key=lambda x: x[0])

# Get total hours
total_hours = 0
for item in summary:
    total_hours += item[1]

# Get worked hours
worked_hours = 0
for item in ppm_worked_hours:
    worked_hours += item[3]

# Write summary list to text file summary.csv
with open("summary.csv", "w", encoding="UTF-8", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Project", "Hours"])
    writer.writerows(summary)

    # If total hours is greater than worked hours add a row
    if total_hours > worked_hours:
        writer.writerow([">>>>>>>>>>>>", ""])
        writer.writerow(["Total Hours", total_hours])
        writer.writerow(["Worked Hours", worked_hours])
        writer.writerow(["<<<<<<<<<<<<", ""])
        writer.writerow(["Difference", total_hours - worked_hours])
        print("Total hours is greater than worked hours by " + str(total_hours - worked_hours) + " hours")
        print("Please check the detail.csv file for details and correct the problem")

# Close the database connections
gcal.util.disconnect_db()
ppm.util.disconnect_db()

# Open the detail.csv file
os.startfile("detail.csv")

# Open the summary.csv file
os.startfile("summary.csv")
