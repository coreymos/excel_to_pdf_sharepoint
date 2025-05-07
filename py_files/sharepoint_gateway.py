# sharepoint_gateway.py
import logging
from pathlib import Path
from typing import List
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.sharepoint.folders.folder import Folder
from azure.identity import DeviceCodeCredential  # updated import

from py_files.config import PUBLIC_GRAPH_CLIENT_ID

logger = logging.getLogger(__name__)

class SharePointGateway:
    """
    Encapsulates SharePoint Online operations.
    """
    def __init__(self, cfg: dict):
        self.tenant = cfg["tenant"].rstrip("/")
        self.site_name = cfg["site"].strip("/")
        self.root_folder = cfg["root_folder"].rstrip("/")
        self.output_folder = cfg["output_folder"].rstrip("/")

        client_id = cfg.get("auth", {}).get("client_id", PUBLIC_GRAPH_CLIENT_ID)
        tenant_id = cfg.get("auth", {}).get("tenant_id", "common")
        cred = DeviceCodeCredential(tenant_id=tenant_id, client_id=client_id)

        site_url = f"https://{self.tenant}/sites/{self.site_name}"
        self.ctx = ClientContext(site_url).with_credentials(cred)
        logger.info("Authenticated to %s", site_url)

    def list_immediate_subfolders(self) -> List[Folder]:
        folder = self.ctx.web.get_folder_by_server_relative_url(self.root_folder)
        self.ctx.load(folder.folders)
        self.ctx.execute_query()
        return list(folder.folders)

    def folder_has_pdf(self, rel_url: str) -> bool:
        folder = self.ctx.web.get_folder_by_server_relative_url(rel_url)
        self.ctx.load(folder.files)
        self.ctx.execute_query()
        return any(f.name.lower().endswith("_lease_leadpaint_xrf.pdf") for f in folder.files)

    def download_sources(self, rel_url: str, dest: Path) -> List[Path]:
        folder = self.ctx.web.get_folder_by_server_relative_url(rel_url)
        self.ctx.load(folder.files)
        self.ctx.execute_query()

        files: List[Path] = []
        for item in folder.files:
            if not item.name.lower().endswith((".xls", ".xlsx", ".csv")):
                continue
            local = dest / item.name
            with open(local, "wb") as fh:
                File.open_binary(self.ctx, item.serverRelativeUrl, fh)
            files.append(local)
            logger.info("Downloaded %s", item.serverRelativeUrl)
        return files

    def upload_pdf(self, pdf_path: Path) -> None:
        with open(pdf_path, "rb") as fh:
            File.save_binary(self.ctx, f"{self.output_folder}/{pdf_path.name}", fh)
        logger.info("Uploaded %s", pdf_path.name)