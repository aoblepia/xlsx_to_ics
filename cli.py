import openpyxl
from datetime import datetime
from dateutil import parser
from icalendar import Calendar, Event
import pytz

# this is the same thing but without the gui and the configurator.
# ensure that the filenames and directories are correct for input and output (Line 46-47).

def excel_to_ics(input_excel_file, output_ics_file):
    wb = openpyxl.load_workbook(input_excel_file)
    sheet = wb.active

    cal = Calendar()

    # Set default time zone to Eastern Time
    eastern = pytz.timezone('US/Eastern')

    for row in sheet.iter_rows(min_row=2, values_only=True):
        summary = row[0]
        start_date = row[1]
        start_time = row[2]
        end_date = row[3]
        end_time = row[4]

        event = Event()
        event.add('summary', summary)

        # Combine date and time strings, then parse with dateutil.parser
        start_datetime = parser.parse(f"{start_date} {start_time}")
        end_datetime = parser.parse(f"{end_date} {end_time}")

        # Set time zone
        start_datetime = eastern.localize(start_datetime)
        end_datetime = eastern.localize(end_datetime)

        event.add('dtstart', start_datetime)
        event.add('dtend', end_datetime)

        cal.add_component(event)

    with open(output_ics_file, 'wb') as f:
        f.write(cal.to_ical())

if __name__ == "__main__":
    input_excel_file = 'Book1.xlsx'
    output_ics_file = 'output_calendar.ics'

    excel_to_ics(input_excel_file, output_ics_file)
