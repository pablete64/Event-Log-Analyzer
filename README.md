# Event Report Generator

## Overview
**Event Report Generator** is a Python tool that processes event data from Excel files and generates detailed event reports. The program uses a Tkinter-based graphical user interface (GUI) to collect input from the user for the date range, reads the event data from an Excel file, and generates reports summarizing the events for the specified period.

## Features
- **GUI for Date Range Input**: Allows users to select a start and end date for the analysis via a simple Tkinter input dialog.
- **Event Grouping and Reporting**: Processes events grouped by date, and generates reports with a daily summary of events that include specific start and end conditions.
- **Excel Input**: Reads event data from Excel files and organizes it by entry ID.
- **Event Filtering**: The program identifies and organizes events by their type (`dis` and `res`), filtering and grouping them by date.
- **Detailed Event Reports**: Generates reports in JSON format, detailing the number of days with events and the start and end of each event period.

## Requirements
- Python 3.8 or higher
- pandas
- tkinter
- openpyxl

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/event-report-generator.git
   cd event-report-generator
Install the required dependencies:
pip install -r requirements.txt
Usage

Run the program:
python main.py
The program will prompt you to:
Select a start and end date for the analysis (format: YYYY-MM-DD).
Select the Excel file containing the event data.
The program will process the events and generate a JSON report summarizing the events within the specified date range.
Example Report

The generated report is a JSON object like the following:

{
    "total_days": 6,
    "events": [
        {"start": "20240410", "end": "20240410", "days": 1},
        {"start": "20240411", "end": "20240414", "days": 5}
    ]
}
Functions

create_report(daily_summary)
Creates a report summarizing the events. Returns the total number of days and a list of events, including the start and end dates and the duration in days.

create_daily_summary(raw_report, start, end)
Generates a daily summary of events between the provided start and end dates. The summary includes whether each day contains a dis event and ends with a res event.

group_events(events)
Groups and sorts events by date, identifying whether the event is of type dis or res.

group_raw_info_by_entry_id(excel_rows)
Groups the raw data from the Excel file by entry ID, creating a dictionary of events associated with each entry.

get_entry_id(string)
Extracts the entry ID from a given string using a regular expression.

get_date_range_from_user()
Prompts the user to input a start and end date for the event analysis.

read_excel_and_convert_to_json()
Reads the Excel file and converts the event data into a structured JSON format.

generate_reports(json_data)
Generates event reports based on the processed data.
