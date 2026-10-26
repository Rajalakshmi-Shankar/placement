import os
import requests
from dotenv import load_dotenv

load_dotenv()

url = os.getenv("MAKE_WEBHOOK_URL")

if not url:
    print("ERROR: Webhook URL missing from .env")
else:
    event = {
        "event": "review_completed",
        "candidate_id": "STU001",
        "task_id": "TASK001",
        "status": "Rework Required",
        "attempt": 1
    }

    try:
        response = requests.post(
            url,
            json=event,
            timeout=15
        )
        response.raise_for_status()
        print("SUCCESS: Event sent to Make.com!")

    except requests.RequestException as error:
        print("ERROR:", error)