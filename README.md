# Google Sheets Student Status Manager

This Python script interacts with the Google Sheets API to calculate student status and the Need for Additional Work (NAF) based on attendance and grades.

## Installation

1. Clone the repository:

```bash``` 
git clone https://github.com/tithalss/gsheet-student-status-manager/

2. Install the required dependencies:

pip install -r requirements.txt

## Usage

Obtain the necessary credentials from the Google Developers Console and save them as credentials.json in the project directory.

Run the script:

python main.py

## Configuration

Before running the script, make sure to adjust the following parameters in the main.py file:

- SAMPLE_SPREADSHEET_ID: The ID of the Google Sheets spreadsheet.
- SCOPES: The necessary scopes required for accessing the Google Sheets API.
- SAMPLE_RANGE_NAME: The range of cells in the spreadsheet where the student data is located.

## How it Works

1. The script authenticates with the Google Sheets API using OAuth2 credentials.
2. It retrieves the total number of classes from the spreadsheet.
3. It retrieves the student grades and attendance data from the spreadsheet.
4. For each student, it calculates their status (Passed, Failed, Final Exam, or Failed due to absence) and the NAF if applicable.
5. It updates the spreadsheet with the calculated status and NAF for each student.

## Dependencies

- google-auth
- google-auth-httplib2
- google-auth-oauthlib
- google-api-python-client
- googleapis-common-protos
