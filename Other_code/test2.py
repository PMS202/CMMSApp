import json
import requests

def send_teams_alert(webhook_url: str, title: str, message: str) -> bool:
    """
    Sends a formatted alert message to a Microsoft Teams channel via webhook.

    Args:
        webhook_url (str): The Teams incoming webhook URL.
        title (str): The title of the alert.
        message (str): The main message body.

    Returns:
        bool: True if sent successfully, False otherwise.
    """
    if not webhook_url.startswith("https://"):
        print("❌ Invalid webhook URL.")
        return False

    # Teams expects a JSON payload with 'text' or 'sections'
    payload = {
        "@type": "MessageCard",
        "@context": "https://schema.org/extensions",
        "summary": title,
        "themeColor": "0076D7",  # Blue
        "title": title,
        "text": message
    }

    try:
        response = requests.post(
            webhook_url,
            headers={"Content-Type": "application/json"},
            data=json.dumps(payload),
            timeout=30
        )
        if response.status_code == 200:
            print("✅ Alert sent successfully.")
            return True
        else:
            print(f"❌ Failed to send alert. HTTP {response.status_code}: {response.text}")
            return False
    except requests.exceptions.RequestException as e:
        print(f"❌ Error sending alert: {e}")
        return False


if __name__ == "__main__":
    # Replace with your actual Teams webhook URL
    TEAMS_WEBHOOK_URL = "https://default66d441884e30446d9ab1a4ae1962a1.b1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a9cee6f5f061445791bafe12b56136a5/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=6Wz-3UI5ibZb--T869Igbhnug9VEXWn6uZCNMSPy6mM"

    # Example usage
    send_teams_alert(
        TEAMS_WEBHOOK_URL,
        title="CMMS Daily Report",
        message="MISA25101999, here is your daily report on machine maintenance and performance. Please review the attached data and take necessary actions."
    )
