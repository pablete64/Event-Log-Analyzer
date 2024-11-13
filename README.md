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

2. Install the required dependencies:
    ```bash
   pip install -r requirements.txt

## Run the program:

    ```bash
      python3 CONTADORE_DE_DIAS.py
