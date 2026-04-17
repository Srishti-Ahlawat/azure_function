"""
sharepoint_client.py — Download Excel files from SharePoint via Microsoft Graph API
-----------------------------------------------------------------------------------
Uses Managed Identity (DefaultAzureCredential) to authenticate.
The Function App's Managed Identity must have Sites.Read.All granted via
Microsoft Graph PowerShell (New-MgServicePrincipalAppRoleAssignment).

Required App Settings:
  SP_SITE_HOST     – e.g. "microsoftapc.sharepoint.com"
  SP_SITE_PATH     – e.g. "/teams/DigitalEmployee_PjM"
  SP_FOLDER_PATH   – e.g. "/Shared Documents/Delivery Quality Enablement Agent/TQP DataSet"
"""

import logging
import os

import requests
from azure.identity import DefaultAzureCredential

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class SharePointClient:
    """Downloads files from a SharePoint document library using Graph API."""

    def __init__(self):
        self.site_host = os.environ["SP_SITE_HOST"]
        self.site_path = os.environ["SP_SITE_PATH"]
        self.folder_path = os.environ.get("SP_FOLDER_PATH", "")
        self._credential = DefaultAzureCredential()
        self._site_id = None

    # ── Auth ─────────────────────────────────────────────────────────────────

    def _get_token(self) -> str:
        """Acquire token via Managed Identity for Microsoft Graph."""
        token = self._credential.get_token("https://graph.microsoft.com/.default")
        return token.token

    def _headers(self) -> dict:
        return {"Authorization": f"Bearer {self._get_token()}"}

    # ── Site discovery ───────────────────────────────────────────────────────

    def _get_site_id(self) -> str:
        """Resolve the SharePoint site ID from host + path."""
        if self._site_id:
            return self._site_id

        url = f"{GRAPH_BASE}/sites/{self.site_host}:{self.site_path}"
        resp = requests.get(url, headers=self._headers(), timeout=30)
        resp.raise_for_status()
        self._site_id = resp.json()["id"]
        logger.info("Resolved site ID: %s", self._site_id)
        return self._site_id

    # ── File download ────────────────────────────────────────────────────────

    def download_file(self, filename: str, local_path: str) -> None:
        """
        Download a single file from the SharePoint document library.

        Uses the /root:/{path}:/content endpoint which returns a 302 redirect
        to a pre-authenticated download URL.
        """
        site_id = self._get_site_id()
        # Build the item path inside the drive
        item_path = f"{self.folder_path}/{filename}".replace("//", "/")

        url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:{item_path}:/content"
        logger.info("Downloading: %s", url)

        resp = requests.get(url, headers=self._headers(), timeout=120, allow_redirects=True)
        resp.raise_for_status()

        with open(local_path, "wb") as f:
            f.write(resp.content)
        logger.info("Saved %d bytes -> %s", len(resp.content), local_path)

    def list_folder(self, folder_path: str = None) -> list:
        """List files in a SharePoint folder. Returns list of {name, id, size}."""
        site_id = self._get_site_id()
        path = folder_path or self.folder_path
        url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:{path}:/children"
        resp = requests.get(url, headers=self._headers(), timeout=30)
        resp.raise_for_status()
        items = resp.json().get("value", [])
        return [{"name": i["name"], "id": i["id"], "size": i.get("size", 0)} for i in items]
