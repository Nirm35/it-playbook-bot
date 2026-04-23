import os
import json
from openai import AsyncAzureOpenAI
from graph_client import (list_playbooks, search_playbooks, read_file_content,
                          list_onenote_notebooks, list_onenote_pages,
                          read_onenote_page, search_onenote)
from dotenv import load_dotenv

load_dotenv()

client = AsyncAzureOpenAI(
    api_key=os.getenv("AZURE_OPENAI_KEY"),
    azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
    api_version="2024-02-01"
)

DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4o")

SYSTEM_PROMPT = """אתה סוכן IT מועיל של חברת Valinor Israel.
תפקידך לעזור לצוות IT למצוא מידע מתוך Playbooks פנימיים השמורים ב-SharePoint.
ענה תמיד בעברית אלא אם המשתמש פונה באנגלית.
כשאתה מביא מידע מקובץ, ציין את שם הקובץ כמקור.
אם לא מצאת מידע רלוונטי, אמור זאת בצורה ברורה."""

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "list_playbooks",
            "description": "מחזיר רשימה של כל קבצי ה-Playbook הזמינים ב-SharePoint",
            "parameters": {"type": "object", "properties": {}, "required": []}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "search_playbooks",
            "description": "מחפש קבצי Playbook לפי מילות מפתח",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "מילות החיפוש"}
                },
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_file",
            "description": "קורא את התוכן המלא של קובץ Playbook ספציפי",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_id": {"type": "string", "description": "מזהה הקובץ"},
                    "file_name": {"type": "string", "description": "שם הקובץ כולל סיומת"}
                },
                "required": ["file_id", "file_name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "list_onenote_notebooks",
            "description": "מחזיר רשימה של כל ה-Notebooks ב-OneNote",
            "parameters": {"type": "object", "properties": {}, "required": []}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "list_onenote_pages",
            "description": "מחזיר רשימת דפים מ-OneNote, אופציונלית לפי Notebook",
            "parameters": {
                "type": "object",
                "properties": {
                    "notebook_id": {"type": "string", "description": "מזהה ה-Notebook (אופציונלי)"}
                },
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_onenote_page",
            "description": "קורא את התוכן של דף OneNote ספציפי",
            "parameters": {
                "type": "object",
                "properties": {
                    "page_id": {"type": "string", "description": "מזהה הדף"}
                },
                "required": ["page_id"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "search_onenote",
            "description": "מחפש דפים ב-OneNote לפי מילות מפתח",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "מילות החיפוש"}
                },
                "required": ["query"]
            }
        }
    }
]

async def run_agent(user_message: str) -> str:
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_message}
    ]

    for _ in range(5):
        response = await client.chat.completions.create(
            model=DEPLOYMENT,
            messages=messages,
            tools=TOOLS,
            tool_choice="auto"
        )

        msg = response.choices[0].message

        if msg.tool_calls:
            messages.append(msg)
            for tool_call in msg.tool_calls:
                result = await _execute_tool(tool_call.function.name, tool_call.function.arguments)
                messages.append({
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "content": json.dumps(result, ensure_ascii=False)
                })
        else:
            return msg.content

    return "מצטער, לא הצלחתי לעבד את הבקשה. נסה שוב."

async def _execute_tool(name: str, arguments: str) -> any:
    args = json.loads(arguments)
    if name == "list_playbooks":
        return await list_playbooks()
    elif name == "search_playbooks":
        return await search_playbooks(args["query"])
    elif name == "read_file":
        content = await read_file_content(args["file_id"], args["file_name"])
        return {"file_name": args["file_name"], "content": content[:8000]}
    elif name == "list_onenote_notebooks":
        return await list_onenote_notebooks()
    elif name == "list_onenote_pages":
        return await list_onenote_pages(args.get("notebook_id"))
    elif name == "read_onenote_page":
        content = await read_onenote_page(args["page_id"])
        return {"content": content[:8000]}
    elif name == "search_onenote":
        return await search_onenote(args["query"])
    return {"error": "unknown tool"}