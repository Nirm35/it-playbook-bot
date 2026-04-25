import os
from fastapi import FastAPI, Request, Response
from botbuilder.core import BotFrameworkAdapter, BotFrameworkAdapterSettings
from botbuilder.schema import Activity
from bot import ITPlaybookBot
from dotenv import load_dotenv
 
load_dotenv()
 
SETTINGS = BotFrameworkAdapterSettings(
    app_id=os.getenv("MICROSOFT_APP_ID"),
    app_password=os.getenv("MICROSOFT_APP_PASSWORD"),
    channel_auth_tenant=os.getenv("TENANT_ID")
)
 
adapter = BotFrameworkAdapter(SETTINGS)
bot = ITPlaybookBot()
app = FastAPI()
 
@app.get("/")
async def health():
    return {"status": "IT Playbook Bot is running"}
 
@app.post("/api/messages")
async def messages(request: Request):
    body = await request.json()
    activity = Activity().deserialize(body)
    auth_header = request.headers.get("Authorization", "")
 
    response = await adapter.process_activity(activity, auth_header, bot.on_turn)
    if response:
        return Response(content=str(response.body), status_code=response.status)
    return Response(status_code=201)
 
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)