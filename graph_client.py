import os
import io
import aiohttp
import pdfplumber
import docx
import openpyxl
from azure.identity.aio import ClientSecretCredential
from dotenv import load_dotenv

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("MICROSOFT_APP_ID")
CLIENT_SECRET = os.getenv("MICROSOFT_APP_PASSWORD")
SHAREPOINT_HOSTNAME = os.getenv("SHAREPOINT_HOSTNAME")
SHAREPOINT_SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME")
SHAREPOINT_FOLDER = os.getenv("SHAREPOINT_FOLDER_PATH", "IT Playbooks")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

async def get_token():
    credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = await credential.get_token("https://graph.microsoft.com/.default")
    await credential.close()
    return token.token

async def get_site_id():
    token = await get_token()
    url = f"{GRAPH_BASE}/sites/{SHAREPOINT_HOSTNAME}:/sites/{SHAREPOINT_SITE_NAME}"
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers={"Authorization": f"Bearer {token}"}) as resp:
            data = await resp.json()
            return data.get("id")

async def list_playbooks():
    token = await get_token()
    site_id = await get_site_id()
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{SHAREPOINT_FOLDER}:/children"
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers={"Authorization": f"Bearer {token}"}) as resp:
            data = await resp.json()
            files = []
            for item in data.get("value", []):
                name = item.get("name", "")
                if not item.get("folder"):
                    files.append({
                        "name": name,
                        "id": item.get("id"),
                        "size": item.get("size"),
                        "lastModified": item.get("lastModifiedDateTime")
                    })
            return files

async def search_playbooks(query: str):
    token = await get_token()
    site_id = await get_site_id()
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{SHAREPOINT_FOLDER}:/search(q='{query}')"
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers={"Authorization": f"Bearer {token}"}) as resp:
            data = await resp.json()
            files = []
            for item in data.get("value", []):
                name = item.get("name", "")
                if not item.get("folder"):
                    files.append({
                        "name": name,
                        "id": item.get("id"),
                    })
            return files

async def read_file_content(file_id: str, file_name: str) -> str:
    token = await get_token()
    site_id = await get_site_id()
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/items/{file_id}/content"
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers={"Authorization": f"Bearer {token}"}, allow_redirects=True) as resp:
            content = await resp.read()

    if file_name.endswith(".pdf"):
        return _extract_pdf(content)
    elif file_name.endswith(".docx"):
        return _extract_docx(content)
    elif file_name.endswith(".xlsx"):
        return _extract_xlsx(content)
    return ""

def _extract_pdf(content: bytes) -> str:
    text = []
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text.append(t)
    return "\n".join(text)

def _extract_docx(content: bytes) -> str:
    doc = docx.Document(io.BytesIO(content))
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

def _extract_xlsx(content: bytes) -> str:
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    lines = []
    for sheet in wb.worksheets:
        lines.append(f"Sheet: {sheet.title}")
        for row in sheet.iter_rows(values_only=True):
            row_text = " | ".join([str(c) for c in row if c is not None])
            if row_text.strip():
                lines.append(row_text)
    return "\n".join(lines)

async def list_onenote_notebooks() -> list:
    token = await get_token()
    url = f"{GRAPH_BASE}/sites/{await get_site_id()}/onenote/notebooks"
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers={"Authorization": f"Bearer {token}"}) as resp:
            data = await resp.json()
            return [{"id": nb.get("id"), "name": nb.get("displayName")} for nb in data.get("value", [])]

async def list_onenote_pages(notebook_id: str = None) -> list:
    token = await get_token()
    if notebook_id:
        url = f"{GRAPH_BASE}/sites/{await get_site_id()}/onenote/notebooks/{notebook_id}/sections"
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers={"Authorization": f"Bearer {token}"}) as resp:
                sections_data = await resp.json()
        pages = []
        for section in sections_data.get("value", []):
            section_id = section.get("id")
            section_name = section.get("displayName")
            pages_url = f"{GRAPH_BASE}/sites/{await get_site_id()}/onenote/sections/{section_id}/pages"
            async with aiohttp.ClientSession() as session:
                async with session.get(pages_url, headers={"Authorization": f"Bearer {token}"}) as resp:
                    pages_data = await resp.json()
            for page in pages_data.get("value", []):
                pages.append({
                    "id": page.get("id"),
                    "title": page.get("title"),
                    "section": section_name
                })
        return pages
    else:
        url = f"{GRAPH_BASE}/sites/{await get_site_id()}/onenote/pages"
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers={"Authorization": f"Bearer {token}"}) as resp:
                data = await resp.json()
                return [{"id": p.get("id"), "title": p.get("title")} for p in data.get("value", [])]

async def read_onenote_page(page_id: str) -> str:
    token = await get_token()
    url = f"{GRAPH_BASE}/sites/{await get_site_id()}/onenote/pages/{page_id}/content"
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers={"Authorization": f"Bearer {token}"}) as resp:
            html_content = await resp.text()
    return _extract_text_from_html(html_content)

async def search_onenote(query: str) -> list:
    token = await get_token()
    url = f"{GRAPH_BASE}/sites/{await get_site_id()}/onenote/pages?$search={query}"
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers={"Authorization": f"Bearer {token}"}) as resp:
            data = await resp.json()
            return [{"id": p.get("id"), "title": p.get("title")} for p in data.get("value", [])]

def _extract_text_from_html(html: str) -> str:
    import re
    clean = re.sub(r'<style[^>]*>.*?</style>', '', html, flags=re.DOTALL)
    clean = re.sub(r'<script[^>]*>.*?</script>', '', clean, flags=re.DOTALL)
    clean = re.sub(r'<br\s*/?>', '\n', clean)
    clean = re.sub(r'<p[^>]*>', '\n', clean)
    clean = re.sub(r'<[^>]+>', '', clean)
    clean = re.sub(r'\n{3,}', '\n\n', clean)
    return clean.strip()