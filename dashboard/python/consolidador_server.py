import base64
import json
import logging
import tempfile
from datetime import date, datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import parse_qs, urlparse

import pandas as pd

from amora_consolidador import (
    APP_NAME,
    UserInputs,
    UserInterface,
    build_default_output_name,
    processar_arquivos,
)
from google_workspace_api import (
    get_api_info,
    get_global_config,
    get_version,
    get_workspace_status,
    list_versions,
    save_global_config,
    save_snapshot,
    setup_workspace,
)


HOST = "127.0.0.1"
PORT = 8765
MAX_REQUEST_BYTES = 200 * 1024 * 1024
ROOT_DIR = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT_DIR / "output" / "consolidador"

LOGGER = logging.getLogger("amora.consolidador.server")
LOG_FORMAT = "%(asctime)s | %(levelname)s | %(message)s"


class ApiUserInterface(UserInterface):
    def __init__(self) -> None:
        self.messages = []

    def _push(self, level: str, titulo: str, mensagem: str) -> None:
        self.messages.append({"level": level, "title": titulo, "message": mensagem})

    def info(self, titulo: str, mensagem: str) -> None:
        self._push("info", titulo, mensagem)

    def warning(self, titulo: str, mensagem: str) -> None:
        self._push("warning", titulo, mensagem)

    def error(self, titulo: str, mensagem: str) -> None:
        self._push("error", titulo, mensagem)

    def ask_open_files(self, titulo: str):
        raise RuntimeError(f"{titulo}: seleção manual não suportada pela API.")

    def ask_open_file(self, titulo: str):
        raise RuntimeError(f"{titulo}: seleção manual não suportada pela API.")

    def ask_save_file(self, titulo: str, default_name: str):
        raise RuntimeError(f"{titulo}: caminho de saída deve ser informado pela API. Sugestão: {default_name}")


def sanitize_scalar(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, Path):
        return str(value)
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, (int, float, bool, str)):
        return value
    return str(value)


def dataframe_to_records(df: pd.DataFrame) -> list[dict[str, Any]]:
    if df is None or df.empty:
        return []
    records: list[dict[str, Any]] = []
    for row in df.to_dict(orient="records"):
        records.append({str(key): sanitize_scalar(value) for key, value in row.items()})
    return records


def safe_output_name(name: Optional[str]) -> str:
    candidate = (name or "").strip() or build_default_output_name()
    safe = "".join(ch for ch in candidate if ch.isalnum() or ch in ("-", "_", ".", " ")).strip()
    if not safe:
        safe = build_default_output_name()
    if not safe.lower().endswith(".xlsx"):
        safe = f"{safe}.xlsx"
    return safe


def safe_upload_name(name: Optional[str], default_name: str) -> str:
    raw_name = Path((name or "").strip() or default_name).name
    stem = "".join(ch for ch in Path(raw_name).stem if ch.isalnum() or ch in ("-", "_", ".", " ")).strip()
    suffix = "".join(ch for ch in Path(raw_name).suffix if ch.isalnum() or ch == ".").lower()

    if not stem:
        fallback = Path(default_name)
        stem = "".join(ch for ch in fallback.stem if ch.isalnum() or ch in ("-", "_", ".", " ")).strip() or "arquivo"
        suffix = suffix or fallback.suffix.lower()

    return f"{stem}{suffix or Path(default_name).suffix.lower()}"


def decode_file_payload(file_payload: Optional[Dict[str, Any]], temp_dir: Path) -> Optional[Path]:
    if not file_payload:
        return None
    content = file_payload.get("contentBase64") or file_payload.get("content_base64")
    if not content:
        return None
    file_name = safe_upload_name(file_payload.get("name"), "arquivo.xlsx")
    file_path = temp_dir / file_name
    file_path.write_bytes(base64.b64decode(content))
    return file_path


def enrich_payload_meta(payload: Dict[str, Any], output_excel: Path, output_json: Path) -> Dict[str, Any]:
    meta = payload.setdefault("meta", {})
    meta.setdefault("sourceFileName", output_excel.name)
    meta.setdefault("sourceFormat", "Consolidador Python")
    meta.setdefault("sheetName", "Consolidado gerado")
    if "generated_at" in meta and "generatedAt" not in meta:
        meta["generatedAt"] = meta["generated_at"]
    if "coverage_from" in meta and "coverageFrom" not in meta:
        meta["coverageFrom"] = meta["coverage_from"]
    if "coverage_to" in meta and "coverageTo" not in meta:
        meta["coverageTo"] = meta["coverage_to"]
    meta.setdefault("outputExcel", str(output_excel))
    meta.setdefault("outputJson", str(output_json))
    return payload


def process_request_payload(data: Dict[str, Any]) -> Dict[str, Any]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="amora_consolidator_") as temp_dir_raw:
        temp_dir = Path(temp_dir_raw)

        nfe_specs = data.get("nfeFiles") or data.get("nfe_files") or []
        nfe_paths = [decode_file_payload(spec, temp_dir) for spec in nfe_specs]
        nfe_paths = [str(path) for path in nfe_paths if path]
        if not nfe_paths:
            raise ValueError("Selecione ao menos um relatório de NFe para executar o consolidator.")

        vendas_item_path = decode_file_payload(data.get("vendasItemFile") or data.get("vendas_item_file"), temp_dir)
        lista_vendas_path = decode_file_payload(data.get("listaVendasFile") or data.get("lista_vendas_file"), temp_dir)

        output_name = safe_output_name(data.get("outputName") or data.get("output_name"))
        output_path = OUTPUT_DIR / output_name

        ui = ApiUserInterface()
        result = processar_arquivos(
            UserInputs(
                arquivos_nfe=nfe_paths,
                caminho_vendas_item=str(vendas_item_path) if vendas_item_path else None,
                caminho_lista_vendas=str(lista_vendas_path) if lista_vendas_path else None,
                caminho_saida=output_path,
            ),
            ui,
        )

        if result is None:
            raise RuntimeError("O consolidator não retornou resultado.")

        payload = json.loads(result.caminho_json.read_text(encoding="utf-8"))
        payload = enrich_payload_meta(payload, result.caminho_saida, result.caminho_json)

        return {
            "ok": True,
            "app": APP_NAME,
            "messages": ui.messages,
            "outputExcel": str(result.caminho_saida),
            "outputJson": str(result.caminho_json),
            "payload": payload,
            "metrics": dataframe_to_records(result.metricas),
            "inconsistencies": dataframe_to_records(result.inconsistencias),
        }


class ConsolidatorRequestHandler(BaseHTTPRequestHandler):
    server_version = "AmoraConsolidatorServer/2.0"

    def end_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
        super().end_headers()

    def log_message(self, fmt: str, *args: Any) -> None:
        LOGGER.info("%s - %s", self.address_string(), fmt % args)

    def _send_json(self, status: int, payload: Dict[str, Any]) -> None:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _read_json_body(self) -> Dict[str, Any]:
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            raise ValueError("Body vazio.")
        if content_length > MAX_REQUEST_BYTES:
            raise ValueError("Payload excede o limite suportado pelo servidor local.")
        raw = self.rfile.read(content_length)
        return json.loads(raw.decode("utf-8"))

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        if path in ("/health", "/api/health"):
            self._send_json(200, {"ok": True, "status": "healthy", "host": HOST, "port": PORT})
            return

        if path != "/api/google":
            self._send_json(404, {"ok": False, "error": "Rota não encontrada."})
            return

        try:
            params = parse_qs(parsed.query or "")
            action = (params.get("action") or ["ping"])[0]
            payload_raw = (params.get("payload") or ["{}"])[0]
            payload = json.loads(payload_raw or "{}")

            if action == "ping":
                self._send_json(200, {"ok": True, "data": get_api_info()})
            elif action == "workspace_status":
                self._send_json(200, {"ok": True, "data": get_workspace_status()})
            elif action == "list_versions":
                self._send_json(200, {"ok": True, "versions": list_versions()})
            elif action == "get_version":
                self._send_json(200, {"ok": True, "data": get_version(payload)})
            elif action == "get_global_config":
                self._send_json(200, {"ok": True, "data": get_global_config()})
            else:
                self._send_json(400, {"ok": False, "error": f"Ação GET não suportada: {action}"})
        except Exception as exc:
            LOGGER.exception("Falha na API Google local: %s", exc)
            self._send_json(500, {"ok": False, "error": str(exc)})

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        try:
            data = self._read_json_body()

            if path == "/api/consolidar":
                result = process_request_payload(data)
                self._send_json(200, result)
                return

            if path == "/api/google":
                action = data.get("action") or "ping"
                if action == "setup_workspace":
                    self._send_json(200, {"ok": True, "data": setup_workspace()})
                elif action == "save_snapshot":
                    self._send_json(200, {"ok": True, "version": save_snapshot(data)})
                elif action == "save_global_config":
                    self._send_json(200, {"ok": True, "data": save_global_config(data)})
                else:
                    self._send_json(400, {"ok": False, "error": f"Ação POST não suportada: {action}"})
                return

            self._send_json(404, {"ok": False, "error": "Rota não encontrada."})
        except Exception as exc:
            LOGGER.exception("Falha no servidor local: %s", exc)
            self._send_json(500, {"ok": False, "error": str(exc)})


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
    server = ThreadingHTTPServer((HOST, PORT), ConsolidatorRequestHandler)
    LOGGER.info("Servidor local disponível em http://%s:%s", HOST, PORT)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        LOGGER.info("Encerrando servidor local.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
