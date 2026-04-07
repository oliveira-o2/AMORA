import base64
import hashlib
import json
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


ROOT_DIR = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT_DIR / "output" / "google_api"
WORKSPACE_FILE = OUTPUT_DIR / "workspace.json"
CLASP_AUTH_FILE = Path.home() / ".clasprc.json"
TOKEN_URL = "https://oauth2.googleapis.com/token"
DRIVE_API_BASE = "https://www.googleapis.com/drive/v3/files"
DRIVE_UPLOAD_BASE = "https://www.googleapis.com/upload/drive/v3/files"
USERINFO_URL = "https://openidconnect.googleapis.com/v1/userinfo"
WORKSPACE_TITLE = "AMORA Dashboard Workspace"
VERSIONS_INDEX_NAME = "versions_index.json"
GLOBAL_CONFIG_NAME = "global_config.json"
SNAPSHOT_PREFIX = "snapshot_"

_TOKEN_CACHE: Dict[str, Any] = {}


def _load_clasp_auth() -> Dict[str, Any]:
    if not CLASP_AUTH_FILE.exists():
        raise RuntimeError(f"Arquivo de autenticação não encontrado: {CLASP_AUTH_FILE}")
    return json.loads(CLASP_AUTH_FILE.read_text(encoding="utf-8"))


def _read_workspace_config() -> Dict[str, Any]:
    if not WORKSPACE_FILE.exists():
        return {}
    return json.loads(WORKSPACE_FILE.read_text(encoding="utf-8"))


def _write_workspace_config(payload: Dict[str, Any]) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    WORKSPACE_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _decode_jwt_payload(token: str) -> Dict[str, Any]:
    try:
        payload = token.split(".")[1]
        padded = payload + "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(padded.encode("utf-8")).decode("utf-8"))
    except Exception:
        return {}


def _get_access_token(force_refresh: bool = False) -> str:
    now_ms = int(time.time() * 1000)
    if not force_refresh and _TOKEN_CACHE.get("access_token") and _TOKEN_CACHE.get("expiry_date", 0) > now_ms + 60_000:
        return _TOKEN_CACHE["access_token"]

    auth = _load_clasp_auth()
    token_info = auth.get("token", {})
    if not force_refresh and token_info.get("access_token") and token_info.get("expiry_date", 0) > now_ms + 60_000:
        _TOKEN_CACHE.update(token_info)
        return token_info["access_token"]

    client = auth.get("oauth2ClientSettings", {})
    refresh_token = token_info.get("refresh_token")
    if not refresh_token:
        raise RuntimeError("Refresh token do Google não encontrado em .clasprc.json.")

    body = urlencode(
        {
            "client_id": client.get("clientId", ""),
            "client_secret": client.get("clientSecret", ""),
            "refresh_token": refresh_token,
            "grant_type": "refresh_token",
        }
    ).encode("utf-8")
    request = Request(TOKEN_URL, data=body, method="POST")
    request.add_header("Content-Type", "application/x-www-form-urlencoded")
    with urlopen(request, timeout=60) as response:
        data = json.loads(response.read().decode("utf-8"))

    access_token = data.get("access_token")
    expires_in = int(data.get("expires_in", 3600))
    if not access_token:
        raise RuntimeError("Falha ao renovar o token Google.")

    _TOKEN_CACHE.update(
        {
            "access_token": access_token,
            "expiry_date": now_ms + (expires_in * 1000),
            "refresh_token": refresh_token,
        }
    )
    return access_token


def _build_request(url: str, method: str, body: Optional[bytes] = None, content_type: Optional[str] = None) -> Request:
    request = Request(url, data=body, method=method.upper())
    request.add_header("Authorization", f"Bearer {_get_access_token(force_refresh=False)}")
    if content_type:
        request.add_header("Content-Type", content_type)
    return request


def _google_json_request(
    method: str,
    url: str,
    payload: Optional[Dict[str, Any]] = None,
    body: Optional[bytes] = None,
    content_type: Optional[str] = None,
    retry: bool = True,
) -> Dict[str, Any]:
    final_body = body if body is not None else (json.dumps(payload).encode("utf-8") if payload is not None else None)
    final_content_type = content_type or ("application/json; charset=utf-8" if payload is not None else None)
    request = _build_request(url, method, body=final_body, content_type=final_content_type)

    try:
        with urlopen(request, timeout=120) as response:
            raw = response.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except HTTPError as exc:
        raw = exc.read().decode("utf-8") if exc.fp else ""
        if exc.code == 401 and retry:
            _get_access_token(force_refresh=True)
            return _google_json_request(method, url, payload=payload, body=body, content_type=content_type, retry=False)
        if raw:
            try:
                error = json.loads(raw)
                message = error.get("error", {}).get("message")
                if message:
                    raise RuntimeError(message) from exc
            except json.JSONDecodeError:
                pass
        raise RuntimeError(f"Falha Google API ({exc.code}) em {url}.") from exc


def _google_text_request(method: str, url: str, retry: bool = True) -> str:
    request = _build_request(url, method)
    try:
        with urlopen(request, timeout=120) as response:
            return response.read().decode("utf-8")
    except HTTPError as exc:
        if exc.code == 401 and retry:
            _get_access_token(force_refresh=True)
            return _google_text_request(method, url, retry=False)
        raw = exc.read().decode("utf-8") if exc.fp else ""
        if raw:
            try:
                error = json.loads(raw)
                message = error.get("error", {}).get("message")
                if message:
                    raise RuntimeError(message) from exc
            except json.JSONDecodeError:
                pass
        raise RuntimeError(f"Falha Google API ({exc.code}) em {url}.") from exc


def _get_user_email() -> str:
    cached_email = _TOKEN_CACHE.get("email")
    if cached_email:
        return cached_email

    auth = _load_clasp_auth()
    id_token = auth.get("token", {}).get("id_token", "")
    email = _decode_jwt_payload(id_token).get("email")
    if email:
        _TOKEN_CACHE["email"] = email
        return email

    try:
        response = _google_json_request("GET", USERINFO_URL)
        email = response.get("email", "")
        _TOKEN_CACHE["email"] = email
        return email
    except Exception:
        return ""


def _drive_file_url(file_id: str) -> str:
    return f"https://drive.google.com/file/d/{file_id}/view"


def _drive_folder_url(folder_id: str) -> str:
    return f"https://drive.google.com/drive/folders/{folder_id}"


def _list_files(query: str, fields: str = "files(id,name,createdTime,modifiedTime,description)") -> list[Dict[str, Any]]:
    url = f"{DRIVE_API_BASE}?{urlencode({'q': query, 'fields': fields, 'pageSize': 1000, 'supportsAllDrives': 'true', 'includeItemsFromAllDrives': 'true'})}"
    response = _google_json_request("GET", url)
    return response.get("files", [])


def _get_file(file_id: str, fields: str = "id,name,mimeType,createdTime,modifiedTime") -> Dict[str, Any]:
    url = f"{DRIVE_API_BASE}/{file_id}?{urlencode({'fields': fields, 'supportsAllDrives': 'true'})}"
    return _google_json_request("GET", url)


def _download_json_file(file_id: str) -> Dict[str, Any]:
    url = f"{DRIVE_API_BASE}/{file_id}?alt=media&supportsAllDrives=true"
    raw = _google_text_request("GET", url)
    return json.loads(raw) if raw else {}


def _build_multipart_body(metadata: Dict[str, Any], content: str) -> tuple[bytes, str]:
    boundary = f"amora-boundary-{int(time.time() * 1000)}"
    parts = [
        f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{json.dumps(metadata, ensure_ascii=False)}\r\n",
        f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{content}\r\n",
        f"--{boundary}--\r\n",
    ]
    return "".join(parts).encode("utf-8"), f"multipart/related; boundary={boundary}"


def _create_json_file(name: str, parent_id: str, payload: Dict[str, Any], description: str = "") -> Dict[str, Any]:
    metadata = {
        "name": name,
        "parents": [parent_id],
        "mimeType": "application/json",
        "description": description,
    }
    body, content_type = _build_multipart_body(metadata, json.dumps(payload, ensure_ascii=False))
    url = f"{DRIVE_UPLOAD_BASE}?uploadType=multipart&supportsAllDrives=true"
    return _google_json_request("POST", url, body=body, content_type=content_type)


def _update_json_file(file_id: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    url = f"{DRIVE_UPLOAD_BASE}/{file_id}?uploadType=media&supportsAllDrives=true"
    return _google_json_request("PATCH", url, body=body, content_type="application/json; charset=utf-8")


def _find_file_by_name(parent_id: str, name: str) -> Optional[Dict[str, Any]]:
    escaped_name = name.replace("'", "\\'")
    files = _list_files(f"'{parent_id}' in parents and trashed = false and name = '{escaped_name}'")
    return files[0] if files else None


def _create_workspace_folder() -> Dict[str, Any]:
    folder = _google_json_request(
        "POST",
        DRIVE_API_BASE,
        payload={"name": WORKSPACE_TITLE, "mimeType": "application/vnd.google-apps.folder"},
    )
    folder_id = folder.get("id")
    if not folder_id:
        raise RuntimeError("Google Drive não retornou o ID da pasta operacional.")

    workspace = {
        "backend": "drive-json",
        "folderId": folder_id,
        "folderName": folder.get("name", WORKSPACE_TITLE),
        "folderUrl": _drive_folder_url(folder_id),
        "createdAt": datetime.utcnow().isoformat() + "Z",
    }
    _write_workspace_config(workspace)
    return workspace


def _get_or_create_workspace() -> Dict[str, Any]:
    workspace = _read_workspace_config()
    folder_id = workspace.get("folderId")
    if folder_id:
        try:
            folder = _get_file(folder_id)
            workspace["folderName"] = folder.get("name", workspace.get("folderName", WORKSPACE_TITLE))
            workspace["folderUrl"] = _drive_folder_url(folder_id)
            workspace["backend"] = "drive-json"
            _write_workspace_config(workspace)
            return workspace
        except Exception:
            pass
    return _create_workspace_folder()


def _ensure_support_file(workspace: Dict[str, Any], config_key: str, file_name: str, default_payload: Dict[str, Any]) -> str:
    file_id = workspace.get(config_key)
    if file_id:
        try:
            _get_file(file_id)
            return file_id
        except Exception:
            pass

    parent_id = workspace["folderId"]
    existing = _find_file_by_name(parent_id, file_name)
    if existing:
        file_id = existing["id"]
    else:
        created = _create_json_file(file_name, parent_id, default_payload)
        file_id = created["id"]

    workspace[config_key] = file_id
    _write_workspace_config(workspace)
    return file_id


def setup_workspace() -> Dict[str, Any]:
    workspace = _get_or_create_workspace()
    _ensure_support_file(workspace, "versionsIndexFileId", VERSIONS_INDEX_NAME, {"versions": []})
    _ensure_support_file(workspace, "globalConfigFileId", GLOBAL_CONFIG_NAME, {"config": {}, "cfopOverrides": []})
    _write_workspace_config(workspace)
    return get_workspace_status()


def _get_ready_workspace() -> Dict[str, Any]:
    setup_workspace()
    return _read_workspace_config()


def _read_versions_index(workspace: Dict[str, Any]) -> Dict[str, Any]:
    file_id = _ensure_support_file(workspace, "versionsIndexFileId", VERSIONS_INDEX_NAME, {"versions": []})
    return _download_json_file(file_id) or {"versions": []}


def _write_versions_index(workspace: Dict[str, Any], payload: Dict[str, Any]) -> None:
    file_id = _ensure_support_file(workspace, "versionsIndexFileId", VERSIONS_INDEX_NAME, {"versions": []})
    _update_json_file(file_id, payload)


def _read_global_config_file(workspace: Dict[str, Any]) -> Dict[str, Any]:
    file_id = _ensure_support_file(workspace, "globalConfigFileId", GLOBAL_CONFIG_NAME, {"config": {}, "cfopOverrides": []})
    return _download_json_file(file_id) or {"config": {}, "cfopOverrides": []}


def _write_global_config_file(workspace: Dict[str, Any], payload: Dict[str, Any]) -> None:
    file_id = _ensure_support_file(workspace, "globalConfigFileId", GLOBAL_CONFIG_NAME, {"config": {}, "cfopOverrides": []})
    _update_json_file(file_id, payload)


def get_workspace_status() -> Dict[str, Any]:
    workspace = _get_or_create_workspace()
    versions_index = _read_versions_index(workspace)
    versions = versions_index.get("versions") if isinstance(versions_index.get("versions"), list) else []
    global_config = _read_global_config_file(workspace)
    global_has_data = bool(global_config.get("config")) or bool(global_config.get("cfopOverrides"))

    sheets = [
        {"key": "workspace", "name": "drive_workspace", "exists": True, "headerIsReady": True, "rowCount": len(versions)},
        {"key": "versions", "name": "versions_index", "exists": True, "headerIsReady": True, "rowCount": len(versions)},
        {"key": "config", "name": "global_config", "exists": True, "headerIsReady": True, "rowCount": 1 if global_has_data else 0},
    ]
    return {
        "status": "ready",
        "spreadsheetId": workspace["folderId"],
        "spreadsheetName": workspace.get("folderName", WORKSPACE_TITLE),
        "spreadsheetUrl": workspace.get("folderUrl", _drive_folder_url(workspace["folderId"])),
        "sheetCount": len(sheets),
        "sheets": sheets,
    }


def get_api_info() -> Dict[str, Any]:
    workspace = setup_workspace()
    return {
        "status": "ready",
        "spreadsheetId": workspace["spreadsheetId"],
        "spreadsheetName": workspace["spreadsheetName"],
        "spreadsheetUrl": workspace["spreadsheetUrl"],
        "availableActions": {
            "get": ["ping", "workspace_status", "list_versions", "get_version", "get_global_config"],
            "post": ["setup_workspace", "save_snapshot", "save_global_config"],
        },
        "sheets": workspace["sheets"],
    }


def _build_version_id() -> str:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    suffix = hashlib.sha256(str(time.time()).encode("utf-8")).hexdigest()[:6].upper()
    return f"VER-{stamp}-{suffix}"


def _compute_hash(payload: Dict[str, Any]) -> str:
    return hashlib.sha256(json.dumps(payload, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()


def save_snapshot(input_payload: Dict[str, Any]) -> Dict[str, Any]:
    workspace = _get_ready_workspace()
    versions_index = _read_versions_index(workspace)
    versions = versions_index.get("versions") if isinstance(versions_index.get("versions"), list) else []

    sales = input_payload.get("sales") if isinstance(input_payload.get("sales"), list) else []
    stock = input_payload.get("stock") if isinstance(input_payload.get("stock"), list) else []
    summary = input_payload.get("summary") if isinstance(input_payload.get("summary"), dict) else {}
    created_at = datetime.utcnow().isoformat() + "Z"
    version_id = _build_version_id()
    payload_hash = _compute_hash(
        {
            "clientName": input_payload.get("clientName", ""),
            "sourceFileName": input_payload.get("sourceFileName", ""),
            "summary": summary,
            "config": input_payload.get("config", {}),
            "cfopOverrides": input_payload.get("cfopOverrides", []),
            "salesCount": len(sales),
            "stockCount": len(stock),
        }
    )

    snapshot_record = {
        "versionId": version_id,
        "createdAt": created_at,
        "clientName": input_payload.get("clientName", ""),
        "sourceFileName": input_payload.get("sourceFileName", ""),
        "sourceFormat": input_payload.get("sourceFormat", ""),
        "sheetName": input_payload.get("sheetName", ""),
        "coverageFrom": summary.get("coverageFrom", ""),
        "coverageTo": summary.get("coverageTo", ""),
        "salesCount": len(sales),
        "stockCount": len(stock),
        "totalRows": int(summary.get("totalRows") or 0),
        "ignoredRows": int(summary.get("ignoredRows") or 0),
        "unclassifiedRows": int(summary.get("unclassifiedRows") or 0),
        "parentVersionId": input_payload.get("parentVersionId", ""),
        "hash": payload_hash,
        "savedBy": _get_user_email(),
    }

    snapshot_payload = {
        "meta": snapshot_record,
        "config": input_payload.get("config", {}),
        "cfopOverrides": input_payload.get("cfopOverrides", []),
        "sales": sales,
        "stock": stock,
    }
    file_name = f"{SNAPSHOT_PREFIX}{version_id}.json"
    created = _create_json_file(file_name, workspace["folderId"], snapshot_payload, description=version_id)
    snapshot_record["fileId"] = created.get("id", "")
    snapshot_record["fileUrl"] = _drive_file_url(snapshot_record["fileId"]) if snapshot_record["fileId"] else ""

    versions.insert(0, snapshot_record)
    versions_index["versions"] = versions
    _write_versions_index(workspace, versions_index)
    return snapshot_record


def list_versions() -> list[Dict[str, Any]]:
    workspace = _get_ready_workspace()
    versions_index = _read_versions_index(workspace)
    versions = versions_index.get("versions") if isinstance(versions_index.get("versions"), list) else []
    return sorted(versions, key=lambda item: str(item.get("createdAt", "")), reverse=True)


def get_version(payload: Dict[str, Any]) -> Dict[str, Any]:
    version_id = str(payload.get("versionId", "")).strip()
    if not version_id:
        raise RuntimeError("Informe versionId para carregar o snapshot.")

    workspace = _get_ready_workspace()
    versions_index = _read_versions_index(workspace)
    versions = versions_index.get("versions") if isinstance(versions_index.get("versions"), list) else []
    version_meta = next((item for item in versions if item.get("versionId") == version_id), None)
    if not version_meta:
        raise RuntimeError(f"Versão não encontrada: {version_id}")

    file_id = version_meta.get("fileId")
    if not file_id:
        file_name = f"{SNAPSHOT_PREFIX}{version_id}.json"
        existing = _find_file_by_name(workspace["folderId"], file_name)
        if not existing:
            raise RuntimeError(f"Arquivo do snapshot não encontrado: {version_id}")
        file_id = existing["id"]

    snapshot = _download_json_file(file_id)
    if not snapshot:
        raise RuntimeError(f"Snapshot vazio: {version_id}")
    return snapshot


def get_global_config() -> Dict[str, Any]:
    workspace = _get_ready_workspace()
    config = _read_global_config_file(workspace)
    return {
        "config": config.get("config", {}) if isinstance(config, dict) else {},
        "cfopOverrides": config.get("cfopOverrides", []) if isinstance(config, dict) else [],
    }


def save_global_config(input_payload: Dict[str, Any]) -> Dict[str, Any]:
    workspace = _get_ready_workspace()
    updated_at = datetime.utcnow().isoformat() + "Z"
    payload = {
        "config": input_payload.get("config", {}),
        "cfopOverrides": input_payload.get("cfopOverrides", []),
        "updatedAt": updated_at,
        "updatedBy": _get_user_email(),
    }
    _write_global_config_file(workspace, payload)
    return {"updatedAt": updated_at}
