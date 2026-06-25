import httpx
import json
import logging
from config import FIREBASE_FIRESTORE_URL

log = logging.getLogger("firestore_sync")

class FirestoreSync:
    def __init__(self, user_id: str, pc_name: str):
        self.user_id = user_id
        self.pc_name = pc_name

    def sync(self, data: dict):
        date_str = data.get("date")
        if not date_str:
            return

        # Structure: users/{user_id}/pcs/{pc_name}/days/{date}
        # In Firestore REST API, this is:
        url = f"{FIREBASE_FIRESTORE_URL}/users/{self.user_id}/pcs/{self.pc_name}/days/{date_str}"
        
        # Convert Python dict to Firestore Document format
        # This is a bit tedious for complex dicts, so we can just store it as a JSON string
        # to simplify, or properly format it.
        # Let's store it as a JSON string for simplicity, or we can use the proper format.
        # Better: just convert basic types.
        
        payload = {
            "fields": {
                "data_json": {"stringValue": json.dumps(data)}
            }
        }
        
        try:
            # We use patch to create or update the document
            patch_url = f"{url}?updateMask.fieldPaths=data_json"
            resp = httpx.patch(patch_url, json=payload, timeout=10.0)
            if resp.status_code not in (200, 201):
                log.error("Failed to sync to Firestore. Status: %s, Response: %s", resp.status_code, resp.text)
                raise Exception(f"Firestore sync failed: {resp.status_code}")
            log.info("Successfully synced data to Firestore for %s", date_str)
        except Exception as e:
            log.error("Exception during Firestore sync: %s", e)
            raise
