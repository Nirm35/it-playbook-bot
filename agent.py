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

SYSTEM_PROMPT = """אתה סוכן IT מומחה של חברת Valinor Israel (שנקראת גם Comm-IT).
תפקידך לעזור לצוות למצוא מידע מתוך מאות קבצים פנימיים השמורים ב-SharePoint.
ענה תמיד בעברית אלא אם המשתמש פונה באנגלית.

## כללי עבודה חשובים:

### חיפוש מידע:
1. כשמשתמש שואל שאלה — חשוב על כמה מילות מפתח רלוונטיות (בעברית וגם באנגלית) והרץ חיפושים עם search_files.
   לדוגמה: אם שואלים "מה מדיניות הסיסמאות?" → חפש "password policy", "סיסמאות", "password"
2. אם חיפוש אחד לא מצא — נסה מילים אחרות. לפחות 2-3 חיפושים שונים לפני שמוותרים.
3. אחרי שמצאת קבצים רלוונטיים — תמיד קרא את התוכן שלהם עם read_file לפני שאתה עונה.
4. אם חיפוש לא עובד — השתמש ב-list_files לראות תיקיות, ואז list_folder להיכנס לתיקיות רלוונטיות.

### קריאת קבצי Excel:
- בקבצי Excel, שורות סיכום מסומנות עם ⭐ — שים לב אליהן כי הן מכילות סיכומים חשובים.
- כל גיליון (Sheet) מייצג בדרך כלל חודש או נושא — ציין מאיזה גיליון הגיע המידע.
- הכותרות בשורה הראשונה מציינות את שמות העמודות — השתמש בהן כדי לזהות את הנתונים.

### מתן תשובות:
- ענה על השאלה לפי התוכן שקראת מהקבצים — לא לפי שמות קבצים.
- בסוף כל תשובה ציין בדיוק מאיזה קובץ הגיע המידע בפורמט:
  📄 מקור: [שם הקובץ המלא] | גיליון: [שם הגיליון אם רלוונטי] | 🔗 קישור: [הלינק לקובץ]
- אם מצאת מידע מכמה קבצים, ציין את כולם בנפרד עם הלינקים שלהם.
- אם יש כמה קבצים דומים (למשל budget 2024, 2025, 2026) — קרא את כולם וציין לגבי כל פריט מידע מאיזה שנה/קובץ הוא הגיע.
- כשהמשתמש מבקש לינק לקובץ — תן לו את הלינק מהשדה link שחוזר מ-read_file.
- אם לא מצאת — אמור בבירור שלא נמצא מידע ותציע מה אפשר לנסות.

### אל תעשה:
- אל תבקש מהמשתמש שם קובץ — אתה אמור למצוא לבד.
- אל תענה רק "מצאתי קובץ X" — תקרא אותו ותענה על השאלה.
- אל תוותר אחרי חיפוש אחד — נסה מילים אחרות.

### זיכרון שיחה:
- זכור את ההקשר מהשיחה. אם המשתמש שואל שאלת המשך, התייחס למה שדובר קודם."""

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
            "description": "מחפש קבצים בכל ה-SharePoint לפי מילות מפתח — מחפש בשמות קבצים ובתוכן. נסה כמה חיפושים עם מילים שונות בעברית ובאנגלית.",
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
            "description": "קורא את התוכן המלא של קובץ (PDF, Word, Excel, CSV, TXT). מחזיר גם לינק לקובץ. חובה להשתמש בזה אחרי שמצאת קובץ רלוונטי — אל תענה בלי לקרוא את התוכן.",
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

    # הוסף היסטוריית שיחה
    if history and len(history) > 1:
        for msg in history[:-1]:
            messages.append({"role": msg["role"], "content": msg["content"]})

    messages.append({"role": "user", "content": user_message})

    for _ in range(15):
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
        result = await read_file_content(args["file_id"], args["file_name"])
        return {
            "file_name": args["file_name"],
            "content": result["content"][:16000],
            "link": result["webUrl"]
        }
    return {"error": "unknown tool"}