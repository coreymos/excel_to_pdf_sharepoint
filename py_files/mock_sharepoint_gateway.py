# mock_sharepoint_gateway.py
import logging
from pathlib import Path
from typing import List

from office365.sharepoint.folders.folder import Folder  # type: ignore

logger = logging.getLogger(__name__)

class MockSharePointGateway:
    """
    Mocks SharePoint operations by using a local filesystem directory structure.
    """
    def __init__(self, cfg: dict):
        self.local_root = Path(cfg.get('local_root', '.'))
        self.output_folder = Path(cfg.get('local_output', 'output'))
        self.output_folder.mkdir(parents=True, exist_ok=True)
        logger.info("Using local mock root: %s", self.local_root)

    def list_immediate_subfolders(self) -> List[Folder]:
        class DummyFolder:
            def __init__(self, path):
                self.name = path.name
                self.serverRelativeUrl = str(path)
        return [DummyFolder(p) for p in self.local_root.iterdir() if p.is_dir()]

    def folder_has_pdf(self, rel_url: str) -> bool:
        # Interpret rel_url as subdirectory under local_root
        folder = self.local_root / Path(rel_url).name
        return any(p.name.lower().endswith('_lease_leadpaint_xrf.pdf') for p in folder.rglob('*_lease_leadpaint_xrf.pdf'))

    def download_sources(self, rel_url: str, dest: Path) -> List[Path]:
        # Download only from the specific local folder
        src_folder = self.local_root / Path(rel_url).name
        paths: List[Path] = []
        for p in src_folder.iterdir():
            if p.suffix.lower() in ('.xls', '.xlsx', '.csv'):
                dst = dest / p.name
                dst.write_bytes(p.read_bytes())
                paths.append(dst)
        return paths

    def upload_pdf(self, pdf_path: Path) -> None:
        dst = self.output_folder / pdf_path.name
        dst.write_bytes(pdf_path.read_bytes())
        logger.info("Mock uploaded %s to %s", pdf_path.name, dst)