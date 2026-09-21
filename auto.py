import os
import json
import time
import logging
import smtplib

from pathlib import Path
from email.message import EmailMessage

from dotenv import load_dotenv
from groq import Groq

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build


# ============================================================
# LOAD ENVIRONMENT
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

load_dotenv(BASE_DIR / ".env")


# ============================================================
# CONFIGURATION
# ============================================================

GOOGLE_SPREADSHEET_ID = os.environ["GOOGLE_SPREADSHEET_ID"]

GOOGLE_SERVICE_ACCOUNT_FILE = (
    BASE_DIR / "service_account.json"
)

SMTP_EMAIL = os.environ["SMTP_EMAIL"]

SMTP_PASSWORD = os.environ["SMTP_PASSWORD"]

DEFAULT_OWNER_EMAIL = os.getenv(
    "DEFAULT_OWNER_EMAIL",
    SMTP_EMAIL
)

SMTP_HOST = os.getenv(
    "SMTP_HOST",
    "smtp.gmail.com"
)

SMTP_PORT = int(
    os.getenv(
        "SMTP_PORT",
        "465"
    )
)

LOOP_INTERVAL_SECONDS = 60 * 5

STATE_FILE = BASE_DIR / "agent_state.json"


# ============================================================
# GOOGLE SHEETS SCOPES
# ============================================================

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets"
]


# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

logger = logging.getLogger(
    "NeuraMind-Agent"
)


# ============================================================
# GROQ
# ============================================================

client = Groq(
    api_key=os.environ["GROQ_API_KEY"]
)


# ============================================================
# GOOGLE SHEETS CONNECTION
# ============================================================

def get_sheets_service():

    if not GOOGLE_SERVICE_ACCOUNT_FILE.exists():

        raise FileNotFoundError(
            f"\nGoogle service account file not found:\n"
            f"{GOOGLE_SERVICE_ACCOUNT_FILE}\n\n"
            f"Make sure service_account.json is in:\n"
            f"{BASE_DIR}"
        )

    credentials = (
        Credentials
        .from_service_account_file(
            str(GOOGLE_SERVICE_ACCOUNT_FILE),
            scopes=SCOPES
        )
    )

    service = build(
        "sheets",
        "v4",
        credentials=credentials
    )

    return service


# ============================================================
# GET FIRST SHEET / TAB
# ============================================================

def get_sheet_info(service):

    spreadsheet = (
        service
        .spreadsheets()
        .get(
            spreadsheetId=GOOGLE_SPREADSHEET_ID
        )
        .execute()
    )

    sheets = spreadsheet.get(
        "sheets",
        []
    )

    if not sheets:

        raise RuntimeError(
            "No sheets/tabs found in the spreadsheet."
        )

    first_sheet = sheets[0]

    properties = first_sheet[
        "properties"
    ]

    return {
        "title": properties["title"],
        "sheet_id": properties["sheetId"]
    }


# ============================================================
# READ GOOGLE SHEET
# ============================================================

def read_sheet(service, sheet_title):

    range_name = f"'{sheet_title}'!A:Z"

    response = (
        service
        .spreadsheets()
        .values()
        .get(
            spreadsheetId=GOOGLE_SPREADSHEET_ID,
            range=range_name
        )
        .execute()
    )

    values = response.get(
        "values",
        []
    )

    if not values:

        return [], []

    headers = []

    for header in values[0]:

        headers.append(
            str(header).strip()
        )

    rows = []

    for row_number, row in enumerate(
        values[1:],
        start=2
    ):

        data = {
            "_row_number": row_number
        }

        for index, header in enumerate(headers):

            if not header:

                continue

            if index < len(row):

                data[header] = str(
                    row[index]
                ).strip()

            else:

                data[header] = ""

        rows.append(data)

    return headers, rows


# ============================================================
# NORMALIZE COLUMN NAME
# ============================================================

def normalize(text):

    return (
        str(text)
        .lower()
        .strip()
        .replace("_", "")
        .replace("-", "")
        .replace(" ", "")
    )


# ============================================================
# FIND COLUMN
# ============================================================

def find_column(
    headers,
    possible_names
):

    normalized_headers = {
        normalize(header): header
        for header in headers
    }

    for name in possible_names:

        key = normalize(name)

        if key in normalized_headers:

            return normalized_headers[key]

    return None


# ============================================================
# GET VALUE FROM COMMON COLUMN NAMES
# ============================================================

def get_column_value(
    row,
    headers,
    possible_names
):

    column = find_column(
        headers,
        possible_names
    )

    if not column:

        return ""

    return str(
        row.get(
            column,
            ""
        )
    ).strip()


# ============================================================
# LOAD STATE
# ============================================================

def load_state():

    if not STATE_FILE.exists():

        return {
            "processed": []
        }

    try:

        with open(
            STATE_FILE,
            "r",
            encoding="utf-8"
        ) as file:

            state = json.load(file)

        if "processed" not in state:

            state["processed"] = []

        return state

    except Exception:

        logger.warning(
            "Could not read state file. "
            "Starting with empty state."
        )

        return {
            "processed": []
        }


# ============================================================
# SAVE STATE
# ============================================================

def save_state(state):

    temporary_file = (
        BASE_DIR /
        "agent_state.tmp"
    )

    with open(
        temporary_file,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            state,
            file,
            indent=2
        )

    temporary_file.replace(
        STATE_FILE
    )


# ============================================================
# CREATE STABLE LISTING ID
# ============================================================

def get_listing_id(
    row,
    headers
):

    # Try common ID columns first

    listing_id = get_column_value(
        row,
        headers,
        [
            "listing_id",
            "listing id",
            "id",
            "response id",
            "application id"
        ]
    )

    if listing_id:

        return listing_id

    # Google Forms normally has a timestamp.
    # Use timestamp + row number to create a stable ID.

    timestamp = get_column_value(
        row,
        headers,
        [
            "timestamp",
            "time",
            "created at",
            "submitted at"
        ]
    )

    if timestamp:

        return (
            f"{timestamp}|"
            f"ROW-{row['_row_number']}"
        )

    return (
        f"ROW-{row['_row_number']}"
    )


# ============================================================
# DETERMINE WHETHER ROW IS NEW
# ============================================================

def is_new_listing(
    row,
    headers
):

    status = get_column_value(
        row,
        headers,
        [
            "status",
            "processing status",
            "email status"
        ]
    )

    # If a Status column exists,
    # only process NEW / blank rows.

    if status:

        return status.lower() in [
            "new",
            "pending",
            "unprocessed"
        ]

    # No status column:
    # state.json controls duplicates.

    return True


# ============================================================
# GROQ AGENT
# ============================================================

def analyze_listing(
    row,
    headers
):

    prompt = f"""
You are the NeuraMind automated meeting/email agent.

You are given one row from a Google Sheet.

Your job is to understand the row and extract the
information needed to send an automated email.

GOOGLE SHEET HEADERS:

{json.dumps(headers, indent=2)}

GOOGLE SHEET ROW:

{json.dumps(row, indent=2)}

Return ONLY valid JSON.

Use exactly this structure:

{{
    "recipient_email": "",
    "owner_email": "",
    "person_name": "",
    "meeting_title": "",
    "meeting_time": "",
    "meeting_date": "",
    "meeting_link": "",
    "subject": "",
    "message": ""
}}

RULES:

1. Never invent an email address.

2. Never invent a meeting date.

3. Never invent a meeting time.

4. Never invent a Google Meet link.

5. If information is missing, return an empty string.

6. Find the recipient email from the row.

7. Find the meeting owner email from the row if available.

8. If owner_email is missing, leave it empty.

9. The Python application will use the configured
   DEFAULT_OWNER_EMAIL if owner_email is empty.

10. If the row already contains a meeting link,
    preserve it exactly.

11. If the row contains meeting date/time,
    preserve it accurately.

12. Create a concise professional email message.

13. Only use information that exists in the row.

14. Do not make assumptions about information that
    isn't present.

15. Do not output Markdown.

16. Return valid JSON only.
"""

    response = client.chat.completions.create(
        model="openai/gpt-oss-20b",
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ],
        temperature=0
    )

    result = (
        response
        .choices[0]
        .message
        .content
    )

    result = (
        result
        .replace(
            "```json",
            ""
        )
        .replace(
            "```",
            ""
        )
        .strip()
    )

    try:

        return json.loads(result)

    except json.JSONDecodeError:

        logger.error(
            "Groq returned invalid JSON:"
        )

        logger.error(result)

        raise


# ============================================================
# SEND EMAIL
# ============================================================

def send_email(
    recipient_email,
    owner_email,
    subject,
    message
):

    recipients = []

    if recipient_email:

        recipients.append(
            recipient_email
        )

    if owner_email:

        if owner_email not in recipients:

            recipients.append(
                owner_email
            )

    if not recipients:

        raise ValueError(
            "No email recipients found."
        )

    email = EmailMessage()

    email["From"] = SMTP_EMAIL

    email["To"] = ", ".join(
        recipients
    )

    email["Subject"] = subject

    email.set_content(
        message
    )

    logger.info(
        "Connecting to SMTP..."
    )

    with smtplib.SMTP_SSL(
        SMTP_HOST,
        SMTP_PORT
    ) as smtp:

        smtp.login(
            SMTP_EMAIL,
            SMTP_PASSWORD
        )

        smtp.send_message(
            email
        )

    logger.info(
        "Email sent successfully to: %s",
        ", ".join(recipients)
    )


# ============================================================
# UPDATE STATUS
# ============================================================

def update_status(
    service,
    sheet_title,
    headers,
    row_number,
    status
):

    status_column = find_column(
        headers,
        [
            "status",
            "processing status",
            "email status"
        ]
    )

    if not status_column:

        logger.info(
            "No Status column found. "
            "Skipping sheet status update."
        )

        return

    # Find column letter

    column_index = headers.index(
        status_column
    )

    column_number = (
        column_index + 1
    )

    # Convert number to Excel column

    column_letter = ""

    while column_number:

        column_number, remainder = divmod(
            column_number - 1,
            26
        )

        column_letter = (
            chr(65 + remainder)
            + column_letter
        )

    range_name = (
        f"'{sheet_title}'!"
        f"{column_letter}"
        f"{row_number}"
    )

    (
        service
        .spreadsheets()
        .values()
        .update(
            spreadsheetId=GOOGLE_SPREADSHEET_ID,
            range=range_name,
            valueInputOption="RAW",
            body={
                "values": [
                    [status]
                ]
            }
        )
        .execute()
    )

    logger.info(
        "Updated row %s status → %s",
        row_number,
        status
    )


# ============================================================
# PROCESS ONE LISTING
# ============================================================

def process_listing(
    row,
    headers,
    service,
    sheet_title,
    state
):

    listing_id = get_listing_id(
        row,
        headers
    )

    logger.info(
        "Checking listing: %s",
        listing_id
    )

    # --------------------------------------------------------
    # Duplicate protection
    # --------------------------------------------------------

    if listing_id in state["processed"]:

        logger.info(
            "Already processed: %s",
            listing_id
        )

        return

    # --------------------------------------------------------
    # Check NEW status
    # --------------------------------------------------------

    if not is_new_listing(
        row,
        headers
    ):

        logger.info(
            "Not a new listing: %s",
            listing_id
        )

        return

    logger.info(
        "Processing NEW listing: %s",
        listing_id
    )

    # --------------------------------------------------------
    # Groq
    # --------------------------------------------------------

    data = analyze_listing(
        row,
        headers
    )

    recipient_email = (
        data
        .get(
            "recipient_email",
            ""
        )
        .strip()
    )

    owner_email = (
        data
        .get(
            "owner_email",
            ""
        )
        .strip()
    )

    # --------------------------------------------------------
    # Default owner
    # --------------------------------------------------------

    if not owner_email:

        owner_email = (
            DEFAULT_OWNER_EMAIL
        )

    person_name = (
        data
        .get(
            "person_name",
            ""
        )
        .strip()
    )

    meeting_title = (
        data
        .get(
            "meeting_title",
            ""
        )
        .strip()
    )

    meeting_date = (
        data
        .get(
            "meeting_date",
            ""
        )
        .strip()
    )

    meeting_time = (
        data
        .get(
            "meeting_time",
            ""
        )
        .strip()
    )

    meeting_link = (
        data
        .get(
            "meeting_link",
            ""
        )
        .strip()
    )

    subject = (
        data
        .get(
            "subject",
            ""
        )
        .strip()
    )

    message = (
        data
        .get(
            "message",
            ""
        )
        .strip()
    )

    # --------------------------------------------------------
    # Validation
    # --------------------------------------------------------

    if not recipient_email:

        logger.error(
            "No recipient email found for %s",
            listing_id
        )

        return

    # --------------------------------------------------------
    # Build fallback subject
    # --------------------------------------------------------

    if not subject:

        if meeting_title:

            subject = (
                f"Meeting Information - "
                f"{meeting_title}"
            )

        else:

            subject = (
                "NeuraMind Meeting Information"
            )

    # --------------------------------------------------------
    # Build fallback message
    # --------------------------------------------------------

    if not message:

        message = (
            f"Hello"
        )

        if person_name:

            message += (
                f" {person_name}"
            )

        message += ",\n\n"

        message += (
            "Please find the meeting information below.\n\n"
        )

        if meeting_title:

            message += (
                f"Meeting: {meeting_title}\n"
            )

        if meeting_date:

            message += (
                f"Date: {meeting_date}\n"
            )

        if meeting_time:

            message += (
                f"Time: {meeting_time}\n"
            )

        if meeting_link:

            message += (
                f"Meeting Link: {meeting_link}\n"
            )

        message += (
            "\nRegards,\n"
            "NeuraMind"
        )

    # --------------------------------------------------------
    # Send email
    # --------------------------------------------------------

    send_email(
        recipient_email=recipient_email,
        owner_email=owner_email,
        subject=subject,
        message=message
    )

    # --------------------------------------------------------
    # Save processed state
    # --------------------------------------------------------

    state["processed"].append(
        listing_id
    )

    save_state(
        state
    )

    # --------------------------------------------------------
    # Update Google Sheet
    # --------------------------------------------------------

    update_status(
        service,
        sheet_title,
        headers,
        row["_row_number"],
        "PROCESSED"
    )

    logger.info(
        "Finished listing: %s",
        listing_id
    )


# ============================================================
# ONE AGENT CYCLE
# ============================================================

def run_cycle():

    logger.info(
        "Starting agent cycle..."
    )

    service = get_sheets_service()

    sheet_info = get_sheet_info(
        service
    )

    sheet_title = sheet_info[
        "title"
    ]

    logger.info(
        "Using sheet/tab: %s",
        sheet_title
    )

    headers, rows = read_sheet(
        service,
        sheet_title
    )

    logger.info(
        "Found %d rows.",
        len(rows)
    )

    if not headers:

        logger.warning(
            "No headers found."
        )

        return

    logger.info(
        "Columns: %s",
        headers
    )

    state = load_state()

    for row in rows:

        try:

            process_listing(
                row=row,
                headers=headers,
                service=service,
                sheet_title=sheet_title,
                state=state
            )

        except Exception as error:

            logger.exception(
                "Error processing row %s: %s",
                row.get("_row_number"),
                error
            )


# ============================================================
# MAIN HOURLY LOOP
# ============================================================

def run_agent():

    logger.info(
        "=========================================="
    )

    logger.info(
        "NeuraMind Automated Email Agent"
    )

    logger.info(
        "Starting..."
    )

    logger.info(
        "=========================================="
    )

    while True:

        try:

            run_cycle()

        except Exception as error:

            logger.exception(
                "Agent cycle failed: %s",
                error
            )

        logger.info(
            "Cycle complete."
        )

        logger.info(
            "Sleeping for 1 hour..."
        )

        time.sleep(
            LOOP_INTERVAL_SECONDS
        )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    run_agent()