# Excel to iCal Converter

This tool takes an .xlsx spreadsheet and converts it to an .ics file for use on Outlook or Google Calendar. You can customize the format of the converter to read in a variety of spreadsheet formats.

## Installation

1. **Clone the Repository:**

    ```bash
    git clone https://github.com/aoblepia/xlsx_to_ics.git
    cd excel-to-ical-converter
    ```

2. **Create a Virtual Environment (Optional but Recommended):**

    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows, use 'venv\Scripts\activate'
    ```

3. **Install Dependencies:**

    ```bash
    pip install -r requirements.txt
    ```

## Usage

1. **Run the Excel to iCal Converter:**

    ```bash
    python excel_to_ics_graphical.py
    ```

2. **Load Your Excel File:**

    - Click on the load button to select your Excel spreadsheet.

3. **Map Excel Columns to Criteria:**
    - The default configuration is as shown:
        | #        | Column 1     | Column 2  | Column 3  | Column 4  | Column 5 |
        | ---------|--------------|-----------|-----------|-----------|----------|
        | Content  |event_summary |start_date |start_time | end_date  | end_time |

    - To modify the default order, click the Configure button and modify as needed. Ignore columns by selecting the Ignore dropdown.

4. **Choose Output Location:**
    - Select the location where you want to save the generated iCalendar file.

5. **Convert Excel to iCal:**
    - Click on the "Convert" button to generate the iCalendar file.

6. **Conversion Complete:**
    - The tool will display a message indicating that the conversion is completed.

## Notes

- Ensure your Excel file is structured correctly and includes the necessary information.
- This tool assumes that row 1 is reserved for section headers and that data begins on row 2.
- If the parsing is returning errors, ensure that the dates and times are formatted properly in Excel.

Feel free to contribute to this project or report any issues by creating a GitHub issue.

