"""
foundry_client.py — Upload text files to Azure AI Foundry vector store
----------------------------------------------------------------------
Uses the azure-ai-projects SDK with DefaultAzureCredential (Managed Identity).

Required App Settings:
  FOUNDRY_ENDPOINT   – e.g. "https://agentpjm-foundry-01.services.ai.azure.com/api/projects/agent-pjm-project"
  VECTOR_STORE_ID    – ID of the existing vector store to update
"""

import logging
import os
import glob
from pathlib import Path

from azure.ai.projects import AIProjectClient
from azure.identity import DefaultAzureCredential

logger = logging.getLogger(__name__)


class FoundryVectorStoreClient:
    """Manages file uploads to a Foundry agent vector store."""

    def __init__(self):
        self.endpoint = os.environ["FOUNDRY_ENDPOINT"]
        self.vector_store_id = os.environ["VECTOR_STORE_ID"]

        self._credential = DefaultAzureCredential()
        self._project = AIProjectClient(
            endpoint=self.endpoint,
            credential=self._credential,
        )
        self._openai = self._project.get_openai_client()

    def sync_files(self, output_dir: str) -> None:
        """
        Replace all files in the vector store with the new ones from output_dir.

        Strategy:
          1. List existing files in the vector store
          2. Delete all old files
          3. Upload all new .txt files from output_dir
        """
        vs_id = self.vector_store_id

        # ── Step 1: List existing files ──────────────────────────────────────
        existing_files = self._list_vector_store_files(vs_id)
        logger.info("Vector store %s has %d existing files.", vs_id, len(existing_files))

        # ── Step 2: Delete old files ─────────────────────────────────────────
        for file_entry in existing_files:
            file_id = file_entry.id
            try:
                self._openai.vector_stores.files.delete(
                    vector_store_id=vs_id,
                    file_id=file_id,
                )
                logger.info("  Deleted file: %s", file_id)
            except Exception:
                logger.warning("  Failed to delete file %s, continuing...", file_id, exc_info=True)

        # ── Step 3: Upload new files ─────────────────────────────────────────
        txt_files = sorted(glob.glob(os.path.join(output_dir, "*.txt")))
        logger.info("Uploading %d files to vector store %s...", len(txt_files), vs_id)

        for fpath in txt_files:
            fname = os.path.basename(fpath)
            try:
                with open(fpath, "rb") as fh:
                    result = self._openai.vector_stores.files.upload_and_poll(
                        vector_store_id=vs_id,
                        file=fh,
                    )
                logger.info("  Uploaded: %s -> %s", fname, result.id)
            except Exception:
                logger.error("  Failed to upload %s", fname, exc_info=True)
                raise

        logger.info("Sync complete: %d files uploaded.", len(txt_files))

    def _list_vector_store_files(self, vs_id: str) -> list:
        """List all files currently in a vector store."""
        files = []
        file_list = self._openai.vector_stores.files.list(vector_store_id=vs_id)
        for item in file_list:
            files.append(item)
        return files
