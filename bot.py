import logging
import re
from datetime import datetime
import pytz
from google.oauth2 import service_account
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import os.path
import pickle
from telegram import Update
from telegram.ext import Application, MessageHandler, filters, CommandHandler, CallbackContext
from googleapiclient.errors import HttpError

# Enable logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Google Sheets setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '17iaYfvrJH5z7TFhfweQk-1K07WH0yl2vrcJqIDhsSfI'  # Replace with your Spreadsheet ID

creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)  # This is your OAuth info, keep it private
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('sheets', 'v4', credentials=creds)

# Store the spreadsheet ID and sheet name globally
current_spreadsheet_id = SPREADSHEET_ID
current_sheet_name = 'AugustVIP' # Default sheet name

def find_next_empty_row(sheet, range_name):
    """Find the next empty row in the specified range."""
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=current_spreadsheet_id,
            range=range_name
        ).execute()

        values = result.get('values', [])
        last_row = len(values) + 1  # The next row to be used
        return last_row
    except HttpError as err:
        logger.error(f"An error occurred: {err}")
        raise

def log_message_to_sheets(timestamp: str, bet_number: str, unit: str, odds: str) -> None:
    """Log a message to Google Sheets."""
    unit = float(unit.rstrip('u'))

    # Ensure odds have two decimal places
    odds = float(odds)
    formatted_odds = f"{odds:.2f}"

    # Fetch existing dates from the sheet
    existing_dates = []
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=current_spreadsheet_id,
            range=f'{current_sheet_name}!A:A'
        ).execute()
        existing_dates = [row[0] for row in result.get('values', [])]
    except HttpError as err:
        logger.error(f"An error occurred while fetching dates: {err}")

    is_new_date = timestamp not in existing_dates

    next_row = find_next_empty_row(service, f'{current_sheet_name}!A:D')
    range_to_append = f'{current_sheet_name}!A{next_row}:D{next_row}'

    values = [[timestamp, bet_number, unit, formatted_odds]]
    body = {
        'values': values
    }

    try:
        result = service.spreadsheets().values().update(
            spreadsheetId=current_spreadsheet_id,
            range=range_to_append,
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        logger.info(f"{result.get('updatedCells')} cells updated.")

        # Apply bold formatting if it's a new date
        if is_new_date:
            requests = [{
                "repeatCell": {
                    "range": {
                        "sheetId": get_sheet_id(current_spreadsheet_id, current_sheet_name),
                        "startRowIndex": next_row - 1,
                        "endRowIndex": next_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": 4
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "bold": True
                            }
                        }
                    },
                    "fields": "userEnteredFormat.textFormat.bold"
                }
            }]
            service.spreadsheets().batchUpdate(
                spreadsheetId=current_spreadsheet_id,
                body={"requests": requests}
            ).execute()

    except HttpError as err:
        logger.error(f"An error occurred: {err}")

def get_sheet_id(spreadsheet_id, sheet_name):
    """Get the sheet ID from the sheet name."""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        for sheet in spreadsheet['sheets']:
            if sheet['properties']['title'] == sheet_name:
                return sheet['properties']['sheetId']
    except HttpError as err:
        logger.error(f"An error occurred while fetching sheet ID: {err}")
    return None

async def handle_message(update: Update, context: CallbackContext) -> None:
    """Handle incoming messages and log if they match the specified format."""
    message = None

    if update.message:
        if update.message.text:
            message = update.message.text
        elif update.message.caption:
            message = update.message.caption
    elif update.channel_post:
        if update.channel_post.text:
            message = update.channel_post.text
        elif update.channel_post.caption:
            message = update.channel_post.caption

    if message:
        chat_id = update.message.chat_id if update.message else update.channel_post.chat_id

        pattern = r"#(\d+)\s+(\d*\.?\d+u)\s+@(\d+\.?\d*)"
        match = re.match(pattern, message)

        if match:
            bet_number, unit, odds = match.groups()

            # Ensure odds have two decimal places
            odds = float(odds)
            formatted_odds = f"{odds:.2f}"

            uk_tz = pytz.timezone('Europe/London')
            now = datetime.now(uk_tz)
            timestamp = now.strftime('%Y-%m-%d')

            log_message_to_sheets(timestamp, bet_number, unit, formatted_odds)

async def update_spreadsheet_id(update: Update, context: CallbackContext) -> None:
    """Update the spreadsheet ID."""
    global current_spreadsheet_id

    if context.args:
        new_spreadsheet_id = context.args[0]
        if len(new_spreadsheet_id) > 10:
            current_spreadsheet_id = new_spreadsheet_id
            await update.message.reply_text(f"Spreadsheet ID updated to: {current_spreadsheet_id}")
        else:
            await update.message.reply_text("Invalid Spreadsheet ID. Please provide a valid ID.")
    else:
        await update.message.reply_text("Please provide a new Spreadsheet ID.")

async def update_sheet_name(update: Update, context: CallbackContext) -> None:
    """Update the sheet name."""
    global current_sheet_name

    if context.args:
        new_sheet_name = context.args[0]
        if re.match(r'^[\w\s]+$', new_sheet_name):
            current_sheet_name = new_sheet_name
            await update.message.reply_text(f"Sheet name updated to: {current_sheet_name}")
        else:
            await update.message.reply_text("Invalid sheet name. Please provide a valid name.")
    else:
        await update.message.reply_text("Please provide a new sheet name.")

def main() -> None:
    """Start the bot."""
    application = Application.builder().token("7151714303:AAGEpX4t7lGin2YRu2bgu73G1Pkr1aHoVPY").build()

    # Add handlers
    application.add_handler(MessageHandler(filters.ALL & ~filters.COMMAND, handle_message))
    application.add_handler(CommandHandler('SetSpreadSheet', update_spreadsheet_id))
    application.add_handler(CommandHandler('SetSheetName', update_sheet_name))

    # Start the Bot
    application.run_polling()

if __name__ == '__main__':
    main()