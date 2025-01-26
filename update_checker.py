import requests
import logging
from packaging import version
import webbrowser

class UpdateChecker:
    def __init__(self, current_version):
        self.current_version = current_version
        self.github_repo = "catalizcs/audio-switcher"
        self.latest_version = None
        self.download_url = None

    def check_for_updates(self):
        """Check GitHub for latest release version"""
        try:
            response = requests.get(
                f"https://api.github.com/repos/{self.github_repo}/releases/latest",
                timeout=5
            )
            if response.status_code == 200:
                data = response.json()
                self.latest_version = version.parse(data["tag_name"].lstrip("v"))
                self.download_url = data["html_url"]
                return version.parse(self.current_version) < self.latest_version
            return False
        except Exception as e:
            logging.error(f"Failed to check for updates: {e}")
            return False

    def open_download_page(self):
        """Open download page in browser"""
        if self.download_url:
            webbrowser.open(self.download_url)
