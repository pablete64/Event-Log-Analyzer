import datetime
import operator
import re
import json
import tkinter as tk
from tkinter import filedialog
import pandas as pd

DIS = 'dis'
RES = 'res'

VALID_STATES = [DIS, RES]

def create_report(daily_summary):
    """
    Create a report with all events.

    This function receives a list of daily summaries.

    Every summary is a dictionary with the following structure:
    * date: An string with the date
    * contains_dis: A boolean
    * ends_with_res: A boolean

    This function returns a dictionary with the following keys:
    * total_days: A integer with the count of days with events
    * events: A list of events.

    Every event is a dictionary with the following keys:
    * start: A string with the start_date
    * end: A string with the end_date
    * days: A integer

    This is an example of a report
    {
        'total_days': 6,
        'events': [
            {'start': '20240410', 'end': '20240410', 'days': 1},
            {'start': '20240411', 'end': '20240414', 'days': 5}
        ]
    }
    """
    report = {'events': [], 'total_days': 0}
    start = None
    datetime_format = '%Y-%m-%d'

    for day in daily_summary:
        if not start and day['contains_dis']:
            start = day['date']

        if start and day['ends_with_res']:
            # Get delta
            delta = datetime.datetime.strptime(day['date'], datetime_format) - \
                    datetime.datetime.strptime(start, datetime_format)
            # Create event
            event = {
                'start': start,
                'end': day['date'],
                'days': delta.days + 1
            }
            # Update global counter
            report['total_days'] += event['days']

            # Add event
            report['events'].append(event)

            # Set start to None
            start = None

    return report


def create_daily_summary(raw_report, start, end):
    """
    Create a daily summary based on all events.

    This function receives a dictionary with all events group by date, this is an example

    {
        '20240410': ['dis', 'res'],
        '20240411': ['res', 'res'],
        '20240501': ['dis', 'res']
    }

    A 'dis' event must be at the beginning in the start date, if we don't have events
    for that time, we should create one.
    A 'res'event must be the last of the start date, if we don't have events for that time, we
    should create one.

    The output of this function is a daily summary list, and it looks like this

    [
        {date: '20240410', 'contains_dis': True, 'ends_with_res': True},
        {date: '20240411', 'contains_dis': True, 'ends_with_res': False},
        {date: '20240412', 'contains_dis': True, 'ends_with_res': True},
    ]
    """
    daily_summary = []
    # Start
    try:
        raw_report[start].insert(0, DIS)
    except KeyError:
        raw_report[start] = [DIS]
    # End
    try:
        events = raw_report[end].append(RES)
    except KeyError:
        raw_report[end] = [RES]

    for date, events in sorted(raw_report.items()):
        # Discard events with a previous date
        if date < start:
            continue
        # Break the loop with dates higher than end
        if date > end:
            break

        daily_summary.append({
            'date': date,
            'contains_dis': any(DIS == event for event in events),
            'ends_with_res':  RES == events[-1]
        })

    return daily_summary


def group_events(events):
    """
    Group events by date.

    This function receives a list of lists with all events.

    A element of the list must follow this structure

    [id, date, time, event_text]

    This function return a dictionnary group by date, events are sorted by:

    date, time, id

    This is an example

    {
        'date': ['dis', 'res']
    }

    {
        'id' : {
            'date': ['dis', 'res', 'dis']
        }
    }
    """
    result = {}
    events.sort(key = operator.itemgetter(1, 2, 0))
    for event in events:
        raw_text = event[3].lower()
        text = None
        for state in VALID_STATES:
            if f'{state}:' in raw_text:
                text = state

        # Add event if we get a right event
        if text:
            date = event[1]
            # Check if there is a key with that date, just append to the event
            if date in result:
                result[date].append(text)
                # Otherwise create a new element for that date
            else:
                result[date] = [text]

    return result


def get_entry_id(string):
    """
    Search the entry id in a string.

    The substring with the id must be at the beginning

    It returns the id or None
    """
    regex = re.compile(r'Ent\((\d+)\)')

    if result := regex.match(string):
        return result.group(1)

    return None


def group_raw_info_by_entry_id(excel_rows):
    """Group raw info by entry id."""
    result = {}
    for row in excel_rows:
        raw_text = row[5]  # Assuming the event_text is in the 6th column
        if "Ent" in raw_text:
            entry_id = re.findall(r'Ent\((\d+)\)', raw_text)
            if entry_id:
                entry_id = entry_id[0]
                date = datetime.datetime.strptime(row[1], '%d-%m-%Y').strftime('%Y-%m-%d')  # Format date as requested
                time = row[2]
                line = [row[0], date, time, raw_text]
                if entry_id in result:
                    result[entry_id].append(line)
                else:
                    result[entry_id] = [line]

    return result



# Interact with the user to get start and end dates
def get_date_range_from_user():
    root = tk.Tk()
    root.withdraw()
    start_date = tk.simpledialog.askstring("Start Date", "Enter start date (YYYY-MM-DD):")
    end_date = tk.simpledialog.askstring("End Date", "Enter end date (YYYY-MM-DD):")
    return start_date, end_date


# Read Excel file and transform it into JSON
def read_excel_and_convert_to_json():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return None
    try:
        excel_data = pd.read_excel(file_path)
        excel_data = excel_data.values.tolist()[1:]  # Exclude header row
        json_data = group_raw_info_by_entry_id(excel_data)
        return json_data
    except Exception as e:
        print("Error reading Excel file:", e)
        return None


def generate_reports(json_data):
    if not json_data:
        print("No data available to generate reports.")
        return

    for entry_id, entry_data in json_data.items():
        print(f"CONTEO DE DÍAS CON DISPARO PARA LA ENTRADA: {entry_id}:")
        events_by_date = group_events(entry_data)
        start_date, end_date = min(events_by_date.keys()), max(events_by_date.keys())
        daily_summary = create_daily_summary(events_by_date, start_date, end_date)
        report = create_report(daily_summary)
        print(json.dumps(report, indent=4))
        print()


if __name__ == "__main__":
    start_date, end_date = get_date_range_from_user()
    if start_date and end_date:
        excel_data = read_excel_and_convert_to_json()
        generate_reports(excel_data)
    else:
        print("Invalid date range provided.")
