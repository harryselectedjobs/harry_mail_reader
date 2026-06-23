import os
import time
import json
import requests
import logging
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv

load_dotenv()

# =========================
# ENV VARIABLES
# =========================

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REDIRECT_URI = os.getenv("REDIRECT_URI")

TOKENS_FILE = os.getenv("TOKENS_FILE", "tokens.json")
CHECK_INTERVAL = int(os.getenv("CHECK_INTERVAL", "8"))

EMAIL_EXISTS_API = os.getenv("EMAIL_EXISTS_API")
SAVE_TRANSCRIPT_API = os.getenv("SAVE_TRANSCRIPT_API")
DELETE_SEQUENCE_API = os.getenv("DELETE_SEQUENCE_API")
CRM_SEQUENCE_LEAD_API = os.getenv("CRM_SEQUENCE_LEAD_API")

# =========================
# LOGGING
# =========================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("mail_watcher.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

seen_email_ids = set()


def load_tokens():
    with open(TOKENS_FILE, "r") as f:
        return json.load(f)


def save_tokens(tokens):
    with open(TOKENS_FILE, "w") as f:
        json.dump(tokens, f, indent=2)


def refresh_access_token():
    tokens = load_tokens()

    refresh_token = tokens.get("refresh_token")

    if not refresh_token:
        raise Exception("No refresh token found. Run token generation again.")

    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

    data = {
        "grant_type": "refresh_token",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": refresh_token,
        "scope": "Mail.Read offline_access User.Read",
        "redirect_uri": REDIRECT_URI
    }

    response = requests.post(
        url,
        data=data,
        timeout=15
    )

    response.raise_for_status()

    new_tokens = response.json()

    save_tokens(new_tokens)

    logger.info("Access token refreshed successfully")

    return new_tokens["access_token"]


def fetch_latest_emails(access_token):
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    url = "https://graph.microsoft.com/v1.0/me/messages"

    since = (
        datetime.now(timezone.utc) - timedelta(hours=24)
    ).strftime("%Y-%m-%dT%H:%M:%SZ")

    params = {
        "$top": 20,
        "$orderby": "receivedDateTime desc",
        "$select": (
            "id,"
            "subject,"
            "from,"
            "receivedDateTime,"
            "isRead,"
            "bodyPreview"
        ),
        "$filter": f"receivedDateTime ge {since}"
    }

    response = requests.get(
        url,
        headers=headers,
        params=params,
        timeout=15
    )

    if response.status_code == 401:
        logger.info("Token expired. Refreshing access token.")

        access_token = refresh_access_token()

        headers = {
            "Authorization": f"Bearer {access_token}"
        }

        response = requests.get(
            url,
            headers=headers,
            params=params,
            timeout=15
        )

    response.raise_for_status()

    return response.json().get("value", []), access_token


def email_exists(email: str) -> bool:
    try:
        response = requests.get(
            f"{EMAIL_EXISTS_API}/{email}",
            timeout=15
        )

        response.raise_for_status()

        data = response.json()

        exists = data.get("exists", False)

        logger.info(
            f"Email existence check for {email}: {exists}"
        )

        return exists

    except Exception as e:
        logger.error(
            f"Failed email existence check for {email}: {e}"
        )
        return False


def save_transcript(payload: dict) -> bool:
    try:
        response = requests.post(
            SAVE_TRANSCRIPT_API,
            json=payload,
            headers={
                "Content-Type": "application/json"
            },
            timeout=15
        )

        response.raise_for_status()

        logger.info(
            f"Transcript saved successfully for "
            f"{payload['sender_email']}"
        )

        return True

    except Exception as e:
        logger.error(
            f"Failed saving transcript for "
            f"{payload['sender_email']}: {e}"
        )
        return False


def create_crm_sequence_lead(email: str) -> bool:
    try:
        response = requests.post(
            CRM_SEQUENCE_LEAD_API,
            json={"email": email},
            headers={
                "Content-Type": "application/json"
            },
            timeout=15
        )

        response.raise_for_status()

        logger.info(
            f"CRM sequence lead created for {email}"
        )

        return True

    except Exception as e:
        logger.error(
            f"Failed creating CRM sequence lead "
            f"for {email}: {e}"
        )
        return False


def delete_sequence_enrollment(email: str):
    try:
        response = requests.delete(
            f"{DELETE_SEQUENCE_API}/{email}",
            timeout=15
        )

        if response.status_code in [200, 204]:
            logger.info(
                f"Sequence enrollment deleted for {email}"
            )
        else:
            logger.warning(
                f"Delete sequence returned "
                f"{response.status_code} for {email}"
            )

    except Exception as e:
        logger.error(
            f"Failed deleting sequence enrollment "
            f"for {email}: {e}"
        )


def process_new_email(email):
    sender_email = (
        email.get("from", {})
        .get("emailAddress", {})
        .get("address", "")
    )

    receiver_email = "harry@selected.jobs"

    payload = {
        "sender_email": sender_email,
        "receiver_email": receiver_email,
        "subject": email.get("subject", ""),
        "body": email.get("bodyPreview", ""),
        "direction": "inbound",
        "sent_at": email.get(
            "receivedDateTime", ""
        )
    }

    logger.info(
        f"New email received from {sender_email}"
    )

    if not email_exists(sender_email):
        logger.info(
            f"Email not found in system. Ignoring "
            f"{sender_email}"
        )
        return

    if not save_transcript(payload):
        return

    create_crm_sequence_lead(sender_email)

    delete_sequence_enrollment(sender_email)

    logger.info(
        f"Finished processing inbound email "
        f"from {sender_email}"
    )


def run():
    logger.info("Mail watcher started")
    logger.info(
        f"Checking mailbox every "
        f"{CHECK_INTERVAL} seconds"
    )

    tokens = load_tokens()
    access_token = tokens["access_token"]

    while True:
        try:
            emails, access_token = fetch_latest_emails(
                access_token
            )

            new_emails = [
                email
                for email in emails
                if email["id"] not in seen_email_ids
            ]

            if new_emails:
                for email in new_emails:
                    seen_email_ids.add(email["id"])
                    process_new_email(email)
            else:
                logger.info("No new emails")

        except requests.exceptions.RequestException as e:
            logger.error(
                f"Network error: {e}"
            )

        except Exception as e:
            logger.error(
                f"Unexpected error: {e}"
            )

        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    run()