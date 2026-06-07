"""
Discord Bot for natural-language expense/income logging.
Listens on channels: 老婆私帳, 老公私帳, 公帳
"""

import json
import os
import re
from datetime import date
from pathlib import Path

import anthropic
import discord
import requests
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent.parent / ".env")

DISCORD_TOKEN     = os.environ["DISCORD_TOKEN"]
NOTION_TOKEN      = os.environ["NOTION_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]

NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Content-Type": "application/json",
    "Notion-Version": "2022-06-28",
}

CHANNEL_MAP = {
    "老婆私帳": ("wife",  os.environ["DB_WIFE"]),
    "老公私帳": ("hub",   os.environ["DB_HUB"]),
    "公帳":     ("joint", os.environ["DB_JOINT"]),
}

claude = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

intents = discord.Intents.default()
intents.message_content = True
bot = discord.Client(intents=intents)


def parse_transaction(text: str, account: str) -> dict | None:
    today = date.today().isoformat()
    is_joint = account == "joint"

    categories = (
        '["家用","旅遊基金","學習教育","投資","其他雜支"]'
        if is_joint
        else '["餐飲","交通","娛樂","帳單","購物","醫療","其他"]'
    )
    type_field   = "" if is_joint else '"type": "支出" or "收入",'
    income_src   = "" if is_joint else '"income_source": one of ["薪資","美股","其他"] (null if 支出),'
    who_paid     = '"who_paid": one of ["老婆","老公","共同"],' if is_joint else ""

    system = f"""You parse Chinese financial messages for a family finance app.
Today: {today}
Return ONLY valid JSON with:
- item_name: short description
- date: YYYY-MM-DD (default today)
- amount: positive number
{type_field}
- category: one of {categories}
{income_src}
- payment: one of ["信用卡","金融卡","現金","街口","Line Pay","VISA"] (guess or default "信用卡")
{who_paid}
- note: extra info or ""

If not a financial transaction, return {{"error": "not a transaction"}}"""

    msg = claude.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=300,
        system=system,
        messages=[{"role": "user", "content": text}],
    )
    raw = msg.content[0].text.strip()
    raw = re.sub(r"```[a-z]*\n?", "", raw).replace("```", "").strip()
    data = json.loads(raw)
    return None if "error" in data else data


def write_to_notion(db_id: str, data: dict, account: str) -> None:
    props: dict = {
        "項目名稱": {"title": [{"text": {"content": data["item_name"]}}]},
        "日期":     {"date": {"start": data["date"]}},
        "金額":     {"number": data["amount"]},
        "來源":     {"select": {"name": "手動輸入"}},
    }
    if data.get("note"):
        props["備注"] = {"rich_text": [{"text": {"content": data["note"]}}]}

    if account in ("wife", "hub"):
        props["類型"]     = {"select": {"name": data["type"]}}
        props["類別"]     = {"select": {"name": data["category"]}}
        props["付款方式"] = {"select": {"name": data.get("payment", "信用卡")}}
        if data.get("type") == "收入" and data.get("income_source"):
            props["收入來源"] = {"select": {"name": data["income_source"]}}
    else:
        props["類別"]     = {"select": {"name": data["category"]}}
        props["付款方式"] = {"select": {"name": data.get("payment", "信用卡")}}
        props["誰付的"]   = {"select": {"name": data.get("who_paid", "共同")}}

    r = requests.post(
        "https://api.notion.com/v1/pages",
        headers=NOTION_HEADERS,
        json={"parent": {"database_id": db_id}, "properties": props},
    )
    r.raise_for_status()


def format_confirm(data: dict, account: str) -> str:
    is_income = data.get("type") == "收入"
    icon = "💰" if is_income else "💸"
    lines = [
        f"{icon} **{data['item_name']}** 已記錄！",
        f"📅 {data['date']}　💴 ${data['amount']:,}",
        f"🏷️ {data['category']}　💳 {data.get('payment', '-')}",
    ]
    if account in ("wife", "hub"):
        lines.append(f"{'收入' if is_income else '支出'}" +
                     (f" · {data['income_source']}" if is_income and data.get('income_source') else ""))
    else:
        lines.append(f"付款人：{data.get('who_paid', '共同')}")
    if data.get("note"):
        lines.append(f"📝 {data['note']}")
    return "\n".join(lines)


@bot.event
async def on_ready():
    print(f"✅ Bot ready: {bot.user}")


@bot.event
async def on_message(message: discord.Message):
    if message.author.bot:
        return
    channel_name = message.channel.name
    if channel_name not in CHANNEL_MAP:
        return

    account, db_id = CHANNEL_MAP[channel_name]
    text = message.content.strip()
    if not text:
        return

    async with message.channel.typing():
        try:
            data = parse_transaction(text, account)
            if data is None:
                await message.reply("⚠️ 無法辨識為帳務記錄，請重新輸入\n範例：`吃飯 250 信用卡` 或 `薪資入帳 48000`")
                return
            write_to_notion(db_id, data, account)
            await message.reply(format_confirm(data, account))
        except json.JSONDecodeError:
            await message.reply("❌ 解析失敗，請稍後再試")
        except Exception as e:
            await message.reply(f"❌ 錯誤：{e}")


bot.run(DISCORD_TOKEN)
