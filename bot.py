import os
import aiohttp
from botbuilder.core import ActivityHandler, TurnContext, MessageFactory
from botbuilder.schema import ChannelAccount
from agent import run_agent
from dotenv import load_dotenv
 
load_dotenv()
 
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("MICROSOFT_APP_ID")
CLIENT_SECRET = os.getenv("MICROSOFT_APP_PASSWORD")
ALLOWED_GROUP_ID = os.getenv("ALLOWED_GROUP_ID")
 
async def get_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(url, data=data) as resp:
            result = await resp.json()
            return result.get("access_token")
 
async def is_user_authorized(user_aad_id: str) -> bool:
    if not ALLOWED_GROUP_ID:
        return True
    token = await get_token()
    url = f"https://graph.microsoft.com/v1.0/groups/{ALLOWED_GROUP_ID}/members"
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers={"Authorization": f"Bearer {token}"}) as resp:
            data = await resp.json()
            members = data.get("value", [])
            return any(m.get("id") == user_aad_id for m in members)
 
class ITPlaybookBot(ActivityHandler):
 
    async def on_message_activity(self, turn_context: TurnContext):
        user_id = turn_context.activity.from_property.aad_object_id
 
        if user_id and not await is_user_authorized(user_id):
            await turn_context.send_activity(
                MessageFactory.text("מצטער, אין לך הרשאה להשתמש בבוט זה.")
            )
            return
 
        user_message = turn_context.activity.text
        if not user_message or not user_message.strip():
            await turn_context.send_activity(
                MessageFactory.text("שלום! אני בוט ה-IT Playbooks. שלח לי שאלה ואעזור לך למצוא את המידע הרלוונטי.")
            )
            return
 
        await turn_context.send_activity(MessageFactory.text("מחפש... רגע אחד"))
 
        try:
            answer = await run_agent(user_message)
            await turn_context.send_activity(MessageFactory.text(answer))
        except Exception as e:
            await turn_context.send_activity(
                MessageFactory.text(f"אירעה שגיאה: {str(e)}")
            )
 
    async def on_members_added_activity(self, members_added: list[ChannelAccount], turn_context: TurnContext):
        for member in members_added:
            if member.id != turn_context.activity.recipient.id:
                await turn_context.send_activity(
                    MessageFactory.text("שלום! אני בוט ה-IT Playbooks של Valinor Israel. איך אוכל לעזור?")
                )