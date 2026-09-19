import io
import json
import logging
import os
import re

import azure.functions as func
import requests
from azure.identity import DefaultAzureCredential
from docxtpl import DocxTemplate
from jinja2 import Environment

GRAPH = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"

SHEET_NAME = os.environ.get("SHEET_NAME", "Sheet1")

# token -> where to read it from in the Excel workbook.
# Keys contain no spaces so they match the [[token]] placeholders in the Word MASTER.
MAPPING = {
    "client_name": {"sheet": SHEET_NAME, "cell": "B5"},
}

app = func.FunctionApp()


def graph_token() -> str:
    credential = DefaultAzureCredential()
    return credential.get_token(GRAPH_SCOPE).token


def read_fields(token: str, site_id: str, excel_item_id: str) -> dict:
    headers = {"Authorization": f"Bearer {token}"}
    data = {}
    for key, loc in MAPPING.items():
        url = (
            f"{GRAPH}/sites/{site_id}/drive/items/{excel_item_id}"
            f"/workbook/worksheets('{loc['sheet']}')/range(address='{loc['cell']}')"
        )
        resp = requests.get(url, headers=headers, timeout=60)
        resp.raise_for_status()
        # `text` is the formatted display value, so dates/currency keep their Excel formatting.
        data[key] = resp.json()["text"][0][0]
    return data


def download(token: str, site_id: str, item_id: str) -> tuple[bytes, dict]:
    headers = {"Authorization": f"Bearer {token}"}
    meta = requests.get(f"{GRAPH}/sites/{site_id}/drive/items/{item_id}", headers=headers, timeout=60)
    meta.raise_for_status()
    content = requests.get(
        f"{GRAPH}/sites/{site_id}/drive/items/{item_id}/content", headers=headers, timeout=120
    )
    content.raise_for_status()
    return content.content, meta.json()


def fill_template(master_bytes: bytes, data: dict) -> bytes:
    doc = DocxTemplate(io.BytesIO(master_bytes))
    env = Environment(variable_start_string="[[", variable_end_string="]]", autoescape=False)
    doc.render(data, jinja_env=env)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def safe_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("_")
    return cleaned or "Client"


def upload_new(token: str, site_id: str, parent_id: str, filename: str, content: bytes) -> dict:
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }
    url = f"{GRAPH}/sites/{site_id}/drive/items/{parent_id}:/{filename}:/content"
    resp = requests.put(url, headers=headers, data=content, timeout=300)
    resp.raise_for_status()
    return resp.json()


def setting(body: dict, name: str) -> str:
    value = body.get(name) or os.environ.get(name)
    if not value:
        raise ValueError(f"Missing required setting '{name}' (env var or request body)")
    return value


@app.route(route="generate", methods=["POST"], auth_level=func.AuthLevel.FUNCTION)
def generate(req: func.HttpRequest) -> func.HttpResponse:
    try:
        try:
            body = req.get_json()
        except ValueError:
            body = {}
        if not isinstance(body, dict):
            body = {}

        site_id = setting(body, "SHAREPOINT_SITE_ID")
        excel_item_id = setting(body, "EXCEL_ITEM_ID")
        word_item_id = setting(body, "WORD_ITEM_ID")

        token = graph_token()

        data = read_fields(token, site_id, excel_item_id)
        master_bytes, master_meta = download(token, site_id, word_item_id)
        rendered = fill_template(master_bytes, data)

        parent_id = master_meta["parentReference"]["id"]
        filename = f"{safe_filename(data.get('client_name', ''))}_Proposal.docx"
        if filename == master_meta.get("name"):
            raise ValueError("Output filename would overwrite the MASTER template")
        item = upload_new(token, site_id, parent_id, filename, rendered)

        return func.HttpResponse(
            json.dumps({"status": "ok", "url": item.get("webUrl"), "fields": data}),
            status_code=200,
            mimetype="application/json",
        )
    except Exception as exc:  # noqa: BLE001
        logging.exception("generate failed")
        return func.HttpResponse(
            json.dumps({"status": "error", "error": str(exc)}),
            status_code=500,
            mimetype="application/json",
        )
