import os
import json
from openai import AsyncAzureOpenAI
from graph_client import list_files, list_folder, search_files, read_file_content
from dotenv import load_dotenv

load_dotenv()

client = AsyncAzureOpenAI(
    api_key=os.getenv("AZURE_OPENAI_KEY"),
    azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
    api_version="2024-02-01"
)

DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4o")

SYSTEM_PROMPT = """אתה סוכן IT מועיל של חברת Valinor Israel.
תפקידך לעזור לצוות IT למצוא מידע מתוך קבצים פנימיים השמורים ב-SharePoint.
ענה תמיד בעברית אלא אם המשתמש פונה באנגלית.
כשאתה מביא מידע מקובץ, ציין את שם הקובץ כמקור.
אם לא מצאת מידע רלוונטי, אמור זאת בצורה ברורה.

## איך לענות על שאלות:
1. כשמשתמש שואל שאלה — השתמש ב-search_files כדי לחפש קבצים רלוונטיים לפי מילות מפתח מהשאלה.
2. אם החיפוש מצא קבצים — השתמש ב-read_file כדי לקרוא את התוכן שלהם.
3. אם החיפוש לא מצא — השתמש ב-list_files כדי לראות מה יש בתיקייה, ואם יש תת-תיקייה רלוונטית השתמש ב-list_folder כדי לחפש בתוכה.
4. ענה על השאלה לפי התוכן שקראת — לא לפי שם הקובץ בלבד.
5. אל תבקש מהמשתמש שם קובץ — אתה אמור למצוא את המידע בעצמך.
6. אם המשתמש שואל על נושא כללי כמו "budget" — חפש ותקרא מספר קבצים רלוונטיים.
7. זכור את ההקשר מהשיחה — אם המשתמש שואל שאלת המשך, התייחס למה שדובר קודם."""

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "list_files",
            "description": "מחזיר רשימה של כל הקבצים והתיקיות בתיקייה הראשית ב-SharePoint. השתמש בזה כדי לראות מה זמין.",
            "parameters": {"type": "object", "properties": {}, "required": []}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "list_folder",
            "description": "נכנס לתת-תיקייה ומחזיר את הקבצים שבתוכה. השתמש כשאתה רואה תיקייה רלוונטית ברשימה.",
            "parameters": {
                "type": "object",
                "properties": {
                    "folder_name": {"type": "string", "description": "שם התת-תיקייה בדיוק כפי שהוא מופיע ברשימה"}
                },
                "required": ["folder_name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "search_files",
            "description": "מחפש קבצים בכל ה-SharePoint לפי מילות מפתח. מחפש בשמות קבצים ובתוכן. השתמש בזה כצעד ראשון בכל שאלה.",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "מילות החיפוש באנגלית או בעברית"}
                },
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_file",
            "description": "קורא את התוכן המלא של קובץ (PDF, Word, Excel, CSV, TXT). השתמש בזה אחרי שמצאת קובץ רלוונטי.",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_id": {"type": "string", "description": "מזהה הקובץ מהחיפוש או מרשימת הקבצים"},
                    "file_name": {"type": "string", "description": "שם הקובץ כולל סיומת"}
                },
                "required": ["file_id", "file_name"]
            }
        }
    }
]


async def run_agent(user_message: str, history: list = None) -> str:
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT}
    ]

    # הוסף היסטוריית שיחה (בלי ההודעה הנוכחית שכבר נוספה)
    if history and len(history) > 1:
        for msg in history[:-1]:
            messages.append({"role": msg["role"], "content": msg["content"]})

    messages.append({"role": "user", "content": user_message})

    for _ in range(10):
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
    if name == "list_files":
        return await list_files()
    elif name == "list_folder":
        return await list_folder(args["folder_name"])
    elif name == "search_files":
        return await search_files(args["query"])
    elif name == "read_file":
        content = await read_file_content(args["file_id"], args["file_name"])
        return {"file_name": args["file_name"], "content": content[:8000]}
    return {"error": "unknown tool"}