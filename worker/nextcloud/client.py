from __future__ import annotations
"""
Nextcloud WebDAV Client for TrueNAS SCALE integration.
Handles connection testing, file listing, SHA-256 calculation, and streaming download.
"""
import hashlib
import os
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple, Any
import requests

class NextcloudClient:
    def __init__(
        self,
        webdav_url: str,
        username: str,
        app_password: str,
        folder: str = "/ExcelImports"
    ):
        self.webdav_url = webdav_url.rstrip("/") + "/"
        self.username = username
        self.app_password = app_password
        self.folder = "/" + folder.strip("/")
        self.session = requests.Session()
        self.session.auth = (self.username, self.app_password)
        self.session.headers.update({
            "User-Agent": "NextcloudExcelSyncWorker/1.0",
            "Accept": "application/xml, text/xml, */*"
        })

    def _get_full_url(self, path: str = "") -> str:
        clean_path = path.lstrip("/")
        return f"{self.webdav_url}{clean_path}"

    def test_connection(self) -> Dict[str, any]:
        """
        Comprehensive connection test:
        1. URL reachable
        2. Authentication successful
        3. Target folder accessible
        4. List files verified
        """
        result = {
            "success": False,
            "url_reachable": False,
            "authenticated": False,
            "folder_accessible": False,
            "files_found": 0,
            "error": None
        }

        try:
            # 1 & 2: Test root WebDAV endpoint
            resp = self.session.request("PROPFIND", self.webdav_url, headers={"Depth": "0"}, timeout=10)
            if resp.status_code in (401, 403):
                result["url_reachable"] = True
                result["error"] = f"Authentication failed (HTTP {resp.status_code}). Check App Password."
                return result
            if resp.status_code not in (200, 207):
                result["error"] = f"Server returned unexpected status {resp.status_code}"
                return result

            result["url_reachable"] = True
            result["authenticated"] = True

            # 3: Check folder accessibility
            folder_url = self._get_full_url(self.folder.lstrip("/"))
            folder_resp = self.session.request("PROPFIND", folder_url, headers={"Depth": "1"}, timeout=15)

            if folder_resp.status_code == 404:
                result["error"] = f"Target folder '{self.folder}' does not exist in Nextcloud."
                return result

            if folder_resp.status_code not in (200, 207):
                result["error"] = f"Failed to access folder '{self.folder}' (HTTP {folder_resp.status_code})."
                return result

            result["folder_accessible"] = True
            files = self.list_excel_files()
            result["files_found"] = len(files)
            result["success"] = True
            return result

        except requests.exceptions.ConnectionError as e:
            result["error"] = f"Could not reach Nextcloud host: {str(e)}"
            return result
        except requests.exceptions.Timeout:
            result["error"] = "Connection timed out reaching Nextcloud."
            return result
        except Exception as e:
            result["error"] = f"Unexpected error during connection test: {str(e)}"
            return result

    def list_excel_files(self) -> List[Dict[str, any]]:
        """
        Scans target Nextcloud folder for .xlsx and .xlsm files via WebDAV PROPFIND.
        """
        folder_url = self._get_full_url(self.folder.lstrip("/"))
        body = """<?xml version="1.0" encoding="utf-8" ?>
        <D:propfind xmlns:D="DAV:">
            <D:prop>
                <D:displayname/>
                <D:getcontentlength/>
                <D:getlastmodified/>
                <D:resourcetype/>
                <D:getetag/>
            </D:prop>
        </D:propfind>"""

        headers = {"Depth": "1", "Content-Type": "application/xml; charset=utf-8"}
        resp = self.session.request("PROPFIND", folder_url, data=body, headers=headers, timeout=20)
        
        if resp.status_code not in (200, 207):
            return []

        excel_files = []
        try:
            root = ET.fromstring(resp.content)
            # WebDAV response namespaces
            ns = {"d": "DAV:"}
            for response in root.findall("d:response", ns):
                href = response.find("d:href", ns)
                if href is None or not href.text:
                    continue

                propstat = response.find("d:propstat", ns)
                if propstat is None:
                    continue
                prop = propstat.find("d:prop", ns)
                if prop is None:
                    continue

                resourcetype = prop.find("d:resourcetype", ns)
                # Skip collections (folders)
                if resourcetype is not None and len(resourcetype) > 0:
                    continue

                file_path = href.text
                filename = os.path.basename(file_path.rstrip("/"))
                
                # Check for excel extensions
                if not (filename.lower().endswith(".xlsx") or filename.lower().endswith(".xlsm")):
                    continue
                # Ignore temporary Excel lock files (~$filename.xlsx)
                if filename.startswith("~$"):
                    continue

                content_len = prop.find("d:getcontentlength", ns)
                file_size = int(content_len.text) if content_len is not None and content_len.text else 0

                last_mod = prop.find("d:getlastmodified", ns)
                last_modified = last_mod.text if last_mod is not None else ""

                etag = prop.find("d:getetag", ns)
                etag_val = etag.text.strip('"') if etag is not None and etag.text else ""

                excel_files.append({
                    "filename": filename,
                    "relative_path": file_path,
                    "file_size": file_size,
                    "last_modified": last_modified,
                    "etag": etag_val,
                    "url": self._get_full_url(file_path)
                })
        except Exception as e:
            print(f"[Nextcloud] Error parsing PROPFIND XML: {e}")

        return excel_files

    def download_file(self, file_path: str, destination_path: str) -> Tuple[bool, str, int]:
        """
        Streams file download to destination and computes SHA-256 hash on the fly.
        Returns: (success, sha256_hash, file_size_bytes)
        """
        os.makedirs(os.path.dirname(destination_path), exist_ok=True)
        download_url = self._get_full_url(file_path.lstrip("/"))

        sha256 = hashlib.sha256()
        total_bytes = 0

        try:
            with self.session.get(download_url, stream=True, timeout=60) as r:
                r.raise_for_status()
                with open(destination_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=65536):
                        if chunk:
                            f.write(chunk)
                            sha256.update(chunk)
                            total_bytes += len(chunk)

            return True, sha256.hexdigest(), total_bytes
        except Exception as e:
            print(f"[Nextcloud] Failed to download {file_path}: {e}")
            return False, "", 0
