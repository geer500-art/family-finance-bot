"""
Discord Bot for natural-language expense/income logging.
Listens on channels: 老婆私帳, 老公私帳, 公帳

Every text message goes straight to Claude (tool use). Claude decides whether to
log a transaction, cancel the last entry, confirm a pending import, analyze
spending, or just chat — there is no keyword/regex routing for text.

File attachments (PDF / Excel / CSV / image) are parsed into a preview and
held as a pending import until the user confirms via chat.
"""

import calendar
import csv
import io
import json
import os
import re
from datetime import date, datetime
from pathlib import Path

import anthropic
import discord
import openpyxl
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

last_page_id: dict[str, str] = {}
chat_history: dict[str, list] = {}  # channel_name -> message history
pending_imports: dict[str, dict] = {}  # channel_name -> {transactions, file_type, db_id}


# ── Notion helpers ──────────────────────────────────────────────

def query_notion(db_id: str, filter_body: dict = None) -> list[dict]:
    pages, cursor = [], None
    while True:
        body = {}
        if filter_body:
            body["filter"] = filter_body
        if cursor:
            body["start_cursor"] = cursor
        r = requests.post(
            f"https://api.notion.com/v1/databases/{db_id}/query",
            headers=NOTION_HEADERS,
            json=body,
        )
        r.raise_for_status()
        data = r.json()
        pages.extend(data.get("results", []))
        if not data.get("has_more"):
            break
        cursor = data.get("next_cursor")
    return pages


def prop_num(page, name):
    return page.get("properties", {}).get(name, {}).get("number") or 0

def prop_select(page, name):
    s = page.get("properties", {}).get(name, {}).get("select")
    return s["name"] if s else ""

def prop_title(page, name):
    t = page.get("properties", {}).get(name, {}).get("title", [])
    return t[0]["plain_text"] if t else ""

def prop_date(page, name):
    d = page.get("properties", {}).get(name, {}).get("date")
    return d["start"] if d else ""


def write_to_notion(db_id: str, data: dict, account: str) -> str:
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
    return r.json()["id"]


def delete_notion_page(page_id: str) -> None:
    requests.patch(
        f"https://api.notion.com/v1/pages/{page_id}",
        headers=NOTION_HEADERS,
        json={"archived": True},
    ).raise_for_status()


# ── Analysis ────────────────────────────────────────────────────

def fetch_month_data(db_id: str, year: int, month: int) -> list[dict]:
    last_day = calendar.monthrange(year, month)[1]
    start = f"{year}-{month:02d}-01"
    end   = f"{year}-{month:02d}-{last_day}"
    return query_notion(db_id, {
        "and": [
            {"property": "日期", "date": {"on_or_after": start}},
            {"property": "日期", "date": {"on_or_before": end}},
        ]
    })


def build_analysis_context(pages: list[dict], account: str) -> str:
    """Summarise Notion pages into a text block for Claude to analyse."""
    rows = []
    for p in pages:
        name = prop_title(p, "項目名稱")
        d    = prop_date(p,  "日期")
        amt  = prop_num(p,   "金額")
        cat  = prop_select(p, "類別")
        typ  = prop_select(p, "類型") if account != "joint" else "支出"
        pay  = prop_select(p, "付款方式")
        rows.append(f"{d} | {name} | {typ} | {cat} | ${amt:,.0f} | {pay}")
    return "\n".join(rows) if rows else "（無資料）"


def fetch_period_context(period: str, db_id: str, account: str) -> str:
    today = date.today()
    year, month = today.year, today.month
    if period == "last_month":
        month -= 1
        if month == 0:
            month, year = 12, year - 1
    pages = fetch_month_data(db_id, year, month)
    context = build_analysis_context(pages, account)
    month_label = f"{year}年{month}月"
    return f"{month_label} 帳務資料（日期 | 項目 | 類型 | 類別 | 金額 | 付款方式）：\n{context}"


# ── File parsers ──────────────────────────────────────────────

def parse_tabular_transactions(rows_text: str, account: str) -> list[dict]:
    today = date.today().isoformat()
    is_joint = account == "joint"
    categories = (
        '["家用","旅遊基金","學習教育","投資","其他雜支"]'
        if is_joint else
        '["餐飲","交通","娛樂","帳單","購物","醫療","其他"]'
    )
    type_field = "" if is_joint else '"type": "支出" or "收入",'
    income_src = "" if is_joint else '"income_source": one of ["薪資","美股","其他"] (null if 支出),'
    who_paid   = '"who_paid": one of ["老婆","老公","共同"],' if is_joint else ""

    system = f"""You are a bank statement parser for a Taiwanese family finance app.
Extract ALL transactions from the tabular data below and return a JSON array.
Today: {today}
Each item must have:
- item_name: short description
- date: YYYY-MM-DD
- amount: positive number
{type_field}
- category: one of {categories}
{income_src}
- payment: one of ["信用卡","金融卡","現金","街口","Line Pay","VISA"]
{who_paid}
- note: extra info or ""

Return ONLY a valid JSON array. If no transactions found, return []"""

    msg = claude.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        system=system,
        messages=[{"role": "user", "content": rows_text}],
    )
    raw = re.sub(r"```[a-z]*\n?", "", msg.content[0].text.strip()).replace("```", "").strip()
    return json.loads(raw)


def excel_to_text(file_bytes: bytes) -> str:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    lines = []
    for row in ws.iter_rows(values_only=True):
        cells = [str(c) if c is not None else "" for c in row]
        if any(c.strip() for c in cells):
            lines.append("\t".join(cells))
    return "\n".join(lines)


def csv_to_text(file_bytes: bytes) -> str:
    text = file_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    return "\n".join("\t".join(row) for row in reader if any(row))


def pdf_to_text(pdf_bytes: bytes, password: str = "") -> str:
    """Extract text from PDF using pypdf, with optional password decryption."""
    import pypdf
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    if reader.is_encrypted:
        result = reader.decrypt(password)
        if result == pypdf.PasswordType.NOT_DECRYPTED:
            raise ValueError("PDF 密碼錯誤，無法解密")
    lines = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            lines.append(text)
    return "\n".join(lines)


def parse_pdf_transactions(pdf_bytes: bytes, account: str, password: str = "") -> list[dict]:
    text = pdf_to_text(pdf_bytes, password)
    if not text.strip():
        raise ValueError("PDF 無法提取文字，可能是掃描版圖片 PDF")
    return parse_tabular_transactions(text, account)


def parse_images_transactions(images: list[tuple[bytes, str]], account: str) -> list[dict]:
    """images: list of (image_bytes, filename). All images are sent in one call
    so transactions split across multiple screenshots are extracted together."""
    import base64
    media_type_map = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "webp": "image/webp"}
    today = date.today().isoformat()
    is_joint = account == "joint"
    categories = (
        '["家用","旅遊基金","學習教育","投資","其他雜支"]'
        if is_joint else
        '["餐飲","交通","娛樂","帳單","購物","醫療","其他"]'
    )
    type_field = "" if is_joint else '"type": "支出" or "收入",'
    income_src = "" if is_joint else '"income_source": one of ["薪資","美股","其他"] (null if 支出),'
    who_paid   = '"who_paid": one of ["老婆","老公","共同"],' if is_joint else ""
    system = f"""You are a bank statement parser for a Taiwanese family finance app.
You will receive {len(images)} screenshot(s), possibly continuations of the same statement.
Extract ALL transactions from ALL images and return a single combined JSON array.
Today: {today}
Each item must have:
- item_name: short description (use 備注/對方帳號 as hint)
- date: YYYY-MM-DD
- amount: positive number (支出 from 支出 column, 收入 from 存入 column)
{type_field}
- category: one of {categories}
{income_src}
- payment: one of ["信用卡","金融卡","現金","街口","Line Pay","VISA"]
{who_paid}
- note: extra info or ""
Return ONLY a valid JSON array covering every transaction in every image. If none found, return []"""

    content = []
    for image_bytes, filename in images:
        ext = filename.rsplit(".", 1)[-1].lower()
        media_type = media_type_map.get(ext, "image/png")
        b64 = base64.standard_b64encode(image_bytes).decode("utf-8")
        content.append({"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}})
    content.append({"type": "text", "text": "請提取以上所有截圖中的所有交易記錄，合併成一個 JSON 陣列回傳。"})

    msg = claude.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        system=system,
        messages=[{"role": "user", "content": content}],
    )
    raw = re.sub(r"```[a-z]*\n?", "", msg.content[0].text.strip()).replace("```", "").strip()
    return json.loads(raw)


# ── Format helpers ──────────────────────────────────────────────

def format_confirm(data: dict, account: str) -> str:
    is_income = data.get("type") == "收入"
    icon = "💰" if is_income else "💸"
    lines = [
        f"{icon} **{data['item_name']}** 已記錄！",
        f"📅 {data['date']}　💴 ${data['amount']:,}",
        f"🏷️ {data['category']}　💳 {data.get('payment', '-')}",
    ]
    if account in ("wife", "hub"):
        lines.append(
            f"{'收入' if is_income else '支出'}" +
            (f" · {data['income_source']}" if is_income and data.get("income_source") else "")
        )
    else:
        lines.append(f"付款人：{data.get('who_paid', '共同')}")
    if data.get("note"):
        lines.append(f"📝 {data['note']}")
    return "\n".join(lines)


def format_import_preview(transactions: list[dict], account: str, file_type: str) -> str:
    lines = [f"📄 從 **{file_type}** 找到 **{len(transactions)}** 筆交易，整理如下：\n"]
    for i, tx in enumerate(transactions, 1):
        amt = tx.get("amount", 0)
        sign = "+" if tx.get("type") == "收入" else "-"
        extra = f"（{tx.get('category','')}）" if account in ("wife", "hub") else f"（{tx.get('who_paid','共同')}）"
        lines.append(f"{i}. {tx.get('date','')} {tx.get('item_name','')} {sign}${amt:,} {extra}")
    lines.append("\n跟我說「確認」就會記錄，或告訴我要改哪一筆／取消。")
    return "\n".join(lines)


# ── Claude tools ──────────────────────────────────────────────

def build_tools(account: str) -> list[dict]:
    is_joint = account == "joint"
    categories = (
        ["家用", "旅遊基金", "學習教育", "投資", "其他雜支"]
        if is_joint else
        ["餐飲", "交通", "娛樂", "帳單", "購物", "醫療", "其他"]
    )
    properties = {
        "item_name": {"type": "string", "description": "項目簡述"},
        "date": {"type": "string", "description": "YYYY-MM-DD，沒提到就用今天"},
        "amount": {"type": "number", "description": "金額，正數"},
        "category": {"type": "string", "enum": categories},
        "payment": {"type": "string", "enum": ["信用卡", "金融卡", "現金", "街口", "Line Pay", "VISA"]},
        "note": {"type": "string", "description": "額外備註，沒有就空字串"},
    }
    required = ["item_name", "date", "amount", "category"]
    if is_joint:
        properties["who_paid"] = {"type": "string", "enum": ["老婆", "老公", "共同"]}
    else:
        properties["type"] = {"type": "string", "enum": ["支出", "收入"]}
        properties["income_source"] = {"type": "string", "enum": ["薪資", "美股", "其他"]}
        required.append("type")

    return [
        {
            "name": "log_transaction",
            "description": "記一筆收入或支出到記帳系統。當使用者描述一筆消費或收入時呼叫（例如「吃飯185」「薪水入帳5萬」）。",
            "input_schema": {"type": "object", "properties": properties, "required": required},
        },
        {
            "name": "cancel_last",
            "description": "取消/刪除上一筆已記錄的交易，或捨棄目前等待確認的匯入內容。當使用者說取消、刪除上一筆時呼叫。",
            "input_schema": {"type": "object", "properties": {}},
        },
        {
            "name": "confirm_pending_import",
            "description": "當使用者確認要把先前匯入預覽（PDF/Excel/CSV/圖片）的交易寫入記帳系統時呼叫，例如使用者說「確認」「好，記錄」。",
            "input_schema": {"type": "object", "properties": {}},
        },
        {
            "name": "analyze_spending",
            "description": "查詢某段期間的帳務原始資料，用於回答花了多少、結餘多少、類別統計等問題。呼叫後你會拿到原始交易清單，自己整理分析。",
            "input_schema": {
                "type": "object",
                "properties": {
                    "period": {"type": "string", "enum": ["this_month", "last_month"], "description": "要查詢的月份"},
                },
                "required": ["period"],
            },
        },
    ]


async def run_tool(name: str, tool_input: dict, account: str, channel_name: str, db_id: str) -> str:
    if name == "log_transaction":
        try:
            page_id = write_to_notion(db_id, tool_input, account)
            last_page_id[channel_name] = page_id
            return json.dumps({"success": True, "confirmation_text": format_confirm(tool_input, account)}, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False)

    if name == "cancel_last":
        if channel_name in pending_imports:
            pending_imports.pop(channel_name)
            return json.dumps({"success": True, "message": "已取消等待確認的匯入，沒有寫入任何資料"}, ensure_ascii=False)
        page_id = last_page_id.get(channel_name)
        if not page_id:
            return json.dumps({"success": False, "message": "找不到上一筆記錄，無法取消"}, ensure_ascii=False)
        try:
            delete_notion_page(page_id)
            last_page_id.pop(channel_name, None)
            return json.dumps({"success": True, "message": "已刪除上一筆記錄"}, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False)

    if name == "confirm_pending_import":
        pending = pending_imports.pop(channel_name, None)
        if not pending:
            return json.dumps({"success": False, "message": "目前沒有等待確認的匯入"}, ensure_ascii=False)
        transactions, file_type, pdb_id = pending["transactions"], pending["file_type"], pending["db_id"]
        ok, fail = 0, 0
        for tx in transactions:
            try:
                page_id = write_to_notion(pdb_id, tx, account)
                last_page_id[channel_name] = page_id
                ok += 1
            except Exception:
                fail += 1
        return json.dumps({"success": True, "imported": ok, "failed": fail, "file_type": file_type}, ensure_ascii=False)

    if name == "analyze_spending":
        period = tool_input.get("period", "this_month")
        try:
            context = fetch_period_context(period, db_id, account)
            return context
        except Exception as e:
            return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False)

    return json.dumps({"error": f"unknown tool {name}"}, ensure_ascii=False)


SYSTEM_TEMPLATE = """你是「富貴號」——楊家的私人財務顧問兼生活助理，個性聰明、親切、有點幽默。
你在 {account_label} 的記帳頻道服務，用繁體中文回答，語氣自然像朋友聊天。

你有以下工具，請主動判斷並呼叫，不要只是嘴上說要做：
- log_transaction：使用者描述一筆消費或收入時呼叫，幫他記到 Notion
- cancel_last：使用者要取消、刪除上一筆記錄，或捨棄剛才匯入預覽時呼叫
- confirm_pending_import：使用者確認要寫入先前匯入預覽的交易時呼叫
- analyze_spending：使用者問花費、結餘、類別統計等問題時呼叫，拿到原始資料後自己整理分析回答（條列式，重點數字加粗，最後給一句建議）

其他情況（純聊天、生活問題、理財建議、計算等）直接自然回答，不要呼叫工具。
回答盡量完整，不要為了精簡而省略重要資訊；純聊天就輕鬆自然。"""


async def handle_chat(text: str, account: str, channel_name: str, db_id: str) -> str:
    account_label = {"wife": "老婆", "hub": "老公", "joint": "公帳"}.get(account, "")
    history = chat_history.setdefault(channel_name, [])
    history.append({"role": "user", "content": text})
    if len(history) > 30:
        history[:] = history[-30:]

    tools = build_tools(account)
    system = SYSTEM_TEMPLATE.format(account_label=account_label)

    for _ in range(4):
        msg = claude.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2048,
            system=system,
            tools=tools,
            messages=history,
        )
        history.append({"role": "assistant", "content": msg.content})

        if msg.stop_reason != "tool_use":
            reply = "".join(b.text for b in msg.content if b.type == "text")
            return reply or "（沒有內容）"

        tool_results = []
        for block in msg.content:
            if block.type != "tool_use":
                continue
            result_text = await run_tool(block.name, block.input, account, channel_name, db_id)
            tool_results.append({
                "type": "tool_result",
                "tool_use_id": block.id,
                "content": result_text,
            })
        history.append({"role": "user", "content": tool_results})

    return "處理時遇到問題，請再試一次。"


# ── Bot events ──────────────────────────────────────────────────

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

    async with message.channel.typing():

        # ── File attachment (PDF / Excel / CSV / Image) → preview ──
        supported = [a for a in message.attachments
                     if a.filename.lower().endswith((".pdf", ".xlsx", ".xls", ".csv",
                                                     ".png", ".jpg", ".jpeg", ".webp"))]
        if supported:
            password = ""
            pw_match = re.search(r"密碼[是:：\s]+(\S+)", text)
            if pw_match:
                password = pw_match.group(1)

            names = "、".join(a.filename for a in supported)
            status_msg = await message.reply(f"📄 正在讀取 **{names}**，請稍候…")
            try:
                image_atts = [a for a in supported if a.filename.lower().endswith((".png", ".jpg", ".jpeg", ".webp"))]
                other_atts = [a for a in supported if a not in image_atts]

                transactions: list[dict] = []
                file_types: list[str] = []

                if image_atts:
                    images = [(await a.read(), a.filename) for a in image_atts]
                    transactions += parse_images_transactions(images, account)
                    file_types.append("圖片")

                for att in other_atts:
                    fname = att.filename.lower()
                    file_bytes = await att.read()
                    if fname.endswith(".pdf"):
                        transactions += parse_pdf_transactions(file_bytes, account, password)
                        file_types.append("PDF")
                    elif fname.endswith((".xlsx", ".xls")):
                        transactions += parse_tabular_transactions(excel_to_text(file_bytes), account)
                        file_types.append("Excel")
                    else:
                        transactions += parse_tabular_transactions(csv_to_text(file_bytes), account)
                        file_types.append("CSV")

                file_type = "、".join(dict.fromkeys(file_types))

                if not transactions:
                    await status_msg.edit(content="⚠️ 找不到任何交易記錄，請確認檔案格式正確。")
                    return

                pending_imports[channel_name] = {"transactions": transactions, "file_type": file_type, "db_id": db_id}
                await status_msg.edit(content=format_import_preview(transactions, account, file_type))

                history = chat_history.setdefault(channel_name, [])
                history.append({"role": "user", "content": f"[系統：使用者上傳了 {file_type}，已解析出 {len(transactions)} 筆交易等待確認]"})
                history.append({"role": "assistant", "content": "好，我已經整理好預覽了，等使用者確認。"})

            except json.JSONDecodeError:
                await status_msg.edit(content="❌ 解析失敗，可能格式不支援\n若為加密 PDF 請附上：`密碼是XXXX`")
            except Exception as e:
                await status_msg.edit(content=f"❌ 錯誤：{e}")
            return

        if not text:
            return

        # ── Everything else goes straight to Claude (tool use) ──
        try:
            reply = await handle_chat(text, account, channel_name, db_id)
            await message.reply(reply)
        except Exception as e:
            await message.reply(f"❌ 錯誤：{e}")


bot.run(DISCORD_TOKEN)
