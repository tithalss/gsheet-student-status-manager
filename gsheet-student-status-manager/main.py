import os.path
import math

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Define the Google Sheets API scopes
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Define the sample spreadsheet ID and range
SAMPLE_SPREADSHEET_ID = "1DTG2R685aQhMW4WjPV03IFvxGAp_ijKqweiAwc-Php4"
SAMPLE_RANGE_NAME = "A1:H27"


def authenticate_google_sheets():
    """Authenticate with Google Sheets API."""
    creds = None

    # Load or create credentials
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return creds


def get_total_classes(sheet):
    """Retrieve the total number of classes from the spreadsheet."""
    classes_total = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="A2")
        .execute()
    )
    total = int(classes_total["values"][0][0].split(":")[-1].strip())
    return total


def calculate_status_and_naf(absence, p1, p2, p3, total):
    """Calculate the status and NAF for a student."""
    average = math.ceil((p1 + p2 + p3) / 3)

    if absence > total * 0.25:
        status = "Failed due to absence"
        naf = 0
    else:
        if average < 50:
            status = "Failed"
            naf = 0
        elif average >= 70:
            status = "Passed"
            naf = 0
        else:
            status = "Final Exam"
            naf = math.ceil((average + max(0, 2 * (7 - average))) / 2)

    return status, naf


def main():
    """Main function to update student status and NAF in the Google Sheet."""
    # Authenticate with Google Sheets API
    creds = authenticate_google_sheets()

    # Initialize Google Sheets service
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    try:
        # Get the total number of classes
        total_classes = get_total_classes(sheet)

        # Get grades from the spreadsheet
        result = (
            sheet.values()
            .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="C4:F27")
            .execute()
        )
        grades = result["values"]

        # Process each student's data and update status and NAF
        status_naf_list = []
        for grade_row in grades:
            absence = int(grade_row[0])
            p1 = int(grade_row[1])
            p2 = int(grade_row[2])
            p3 = int(grade_row[3])
            status, naf = calculate_status_and_naf(absence, p1, p2, p3, total_classes)
            status_naf_list.append([status, naf])

        # Update status and NAF in the spreadsheet
        update_values = {"values": status_naf_list}
        result = (
            sheet.values()
            .update(
                spreadsheetId=SAMPLE_SPREADSHEET_ID,
                range="G4:H27",
                valueInputOption="USER_ENTERED",
                body=update_values,
            )
            .execute()
        )

        print("Student status and NAF updated successfully.")

    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()
