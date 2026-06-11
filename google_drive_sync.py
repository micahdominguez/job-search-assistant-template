from __future__ import annotations

import json
from pathlib import Path
import re
from typing import Any


GOOGLE_DOC_MIME_TYPE = "application/vnd.google-apps.document"
DEFAULT_SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
]
DEFAULT_SHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]
DEFAULT_WORKSPACE_SCOPES = list(dict.fromkeys([*DEFAULT_SCOPES, *DEFAULT_SHEETS_SCOPES]))


def _import_google_clients() -> tuple[Any, ...]:
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials as UserCredentials
        from google.oauth2.service_account import Credentials as ServiceAccountCredentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
        from googleapiclient.errors import HttpError
    except ImportError as exc:  # pragma: no cover - environment-dependent
        raise RuntimeError(
            "Google Drive sync dependencies are missing. "
            "Install requirements-google-drive-sync.txt to enable Google Drive sync."
        ) from exc
    return ServiceAccountCredentials, UserCredentials, Request, InstalledAppFlow, build, HttpError


def extract_drive_id(url_or_id: str) -> str:
    value = (url_or_id or "").strip()
    if not value:
        raise ValueError("A Google Drive folder URL or file ID is required.")
    match = re.search(r"/folders/([A-Za-z0-9_-]+)", value)
    if match:
        return match.group(1)
    match = re.search(r"/d/([A-Za-z0-9_-]+)", value)
    if match:
        return match.group(1)
    if re.fullmatch(r"[A-Za-z0-9_-]{10,}", value):
        return value
    raise ValueError(f"Could not extract a Google Drive ID from: {url_or_id}")


def detect_google_credentials_type(credentials_json: Path) -> str:
    if not credentials_json.exists():
        raise FileNotFoundError(f"Google credentials not found: {credentials_json}")
    try:
        payload = json.loads(credentials_json.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Google credentials file is not valid JSON: {credentials_json}") from exc
    if payload.get("type") == "service_account":
        return "service_account"
    if "installed" in payload or "web" in payload:
        return "oauth_client"
    raise ValueError(
        "Unsupported Google credentials file. Expected a service-account JSON or OAuth client JSON."
    )


def _credentials_support_scopes(creds: Any, scopes: list[str]) -> bool:
    checker = getattr(creds, "has_scopes", None)
    if callable(checker):
        try:
            return bool(checker(scopes))
        except TypeError:
            pass
    granted = set(getattr(creds, "scopes", None) or getattr(creds, "granted_scopes", None) or [])
    return not granted or set(scopes).issubset(granted)


def _token_file_supports_scopes(token_path: Path, scopes: list[str]) -> bool:
    try:
        payload = json.loads(token_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return False
    granted = set(payload.get("scopes") or [])
    if not granted:
        return False
    return set(scopes).issubset(granted)


def load_google_credentials(
    credentials_json: Path,
    scopes: list[str] | None = None,
    *,
    token_json: Path | None = None,
) -> tuple[Any, str]:
    ServiceAccountCredentials, UserCredentials, Request, InstalledAppFlow, _, _ = _import_google_clients()
    scope_list = scopes or DEFAULT_SCOPES
    credentials_type = detect_google_credentials_type(credentials_json)
    if credentials_type == "service_account":
        creds = ServiceAccountCredentials.from_service_account_file(
            str(credentials_json),
            scopes=scope_list,
        )
        return creds, credentials_type

    token_path = token_json or credentials_json.with_name("google-oauth-token.json")
    creds = None
    if token_path.exists():
        if _token_file_supports_scopes(token_path, scope_list):
            creds = UserCredentials.from_authorized_user_file(str(token_path), scope_list)
            if creds and not _credentials_support_scopes(creds, scope_list):
                creds = None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as exc:
                message = str(exc).lower()
                recoverable_errors = (
                    "invalid_scope",
                    "invalid_grant",
                    "expired or revoked",
                    "token has been expired or revoked",
                )
                if not any(error in message for error in recoverable_errors):
                    raise
                creds = None
        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_json), scope_list)
            creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")
    return creds, credentials_type


def build_google_services(
    credentials_json: Path,
    scopes: list[str] | None = None,
    *,
    token_json: Path | None = None,
) -> tuple[Any, Any, str]:
    _, _, _, _, build, _ = _import_google_clients()
    creds, credentials_type = load_google_credentials(
        credentials_json,
        scopes=scopes,
        token_json=token_json,
    )
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    docs = build("docs", "v1", credentials=creds, cache_discovery=False)
    return drive, docs, credentials_type


def build_google_sheets_service(
    credentials_json: Path,
    scopes: list[str] | None = None,
    *,
    token_json: Path | None = None,
) -> tuple[Any, str]:
    _, _, _, _, build, _ = _import_google_clients()
    creds, credentials_type = load_google_credentials(
        credentials_json,
        scopes=scopes or DEFAULT_WORKSPACE_SCOPES,
        token_json=token_json,
    )
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
    return sheets, credentials_type


def is_storage_quota_exceeded_error(exc: Exception) -> bool:
    return "storageQuotaExceeded" in str(exc)


def _escape_drive_query(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def find_google_doc_in_folder(drive: Any, folder_id: str, title: str) -> dict[str, Any] | None:
    escaped_title = _escape_drive_query(title)
    query = (
        f"name = '{escaped_title}' and "
        f"'{folder_id}' in parents and "
        f"mimeType = '{GOOGLE_DOC_MIME_TYPE}' and trashed = false"
    )
    response = drive.files().list(
        q=query,
        pageSize=5,
        fields="files(id,name,webViewLink,parents,modifiedTime)",
        supportsAllDrives=True,
    ).execute()
    files = response.get("files") or []
    return files[0] if files else None


def create_google_doc_in_folder(drive: Any, folder_id: str, title: str) -> dict[str, Any]:
    metadata = {
        "name": title,
        "mimeType": GOOGLE_DOC_MIME_TYPE,
        "parents": [folder_id],
    }
    return drive.files().create(
        body=metadata,
        fields="id,name,webViewLink,parents,modifiedTime",
        supportsAllDrives=True,
    ).execute()


def replace_google_doc_text(docs: Any, document_id: str, text: str) -> None:
    document = docs.documents().get(documentId=document_id).execute()
    body_content = document.get("body", {}).get("content", [])
    end_index = body_content[-1].get("endIndex", 1) if body_content else 1
    requests: list[dict[str, Any]] = []
    delete_end_index = end_index - 1
    if delete_end_index > 1:
        requests.append(
            {
                "deleteContentRange": {
                    "range": {
                        "startIndex": 1,
                        "endIndex": delete_end_index,
                    }
                }
            }
        )
    if text:
        requests.append(
            {
                "insertText": {
                    "location": {"index": 1},
                    "text": text,
                }
            }
        )
    if requests:
        docs.documents().batchUpdate(
            documentId=document_id,
            body={"requests": requests},
        ).execute()


def upsert_google_doc_text(
    drive: Any,
    docs: Any,
    *,
    folder_id: str,
    title: str,
    text: str,
) -> dict[str, Any]:
    existing = find_google_doc_in_folder(drive, folder_id, title)
    file_info = existing or create_google_doc_in_folder(drive, folder_id, title)
    replace_google_doc_text(docs, file_info["id"], text)
    refreshed = drive.files().get(
        fileId=file_info["id"],
        fields="id,name,webViewLink,parents,modifiedTime",
        supportsAllDrives=True,
    ).execute()
    return refreshed
