"""
Discord Bot for natural-language expense/income logging.
Listens on channels: 老婆私帳, 老公私帳, 公帳

Features:
- Natural language transaction parsing via Claude
- Delete last entry: 取消 / 刪除上一筆
- PDF / Excel / CSV bank statement import
- Spending analysis: 分析 / 這個月花了多少 / 類別統計
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

ANALYSIS_KEYWORDS = [
    "分析", "統計", "花了多少", "支出", "收入", "這個月", "上個月",
    "本月", "今年", "比較", "類別", "帳單", "總結", "報告", "幫我看",
    "消費", "結餘", "剩多少", "存了多少",
]

def is_analysis_request(text: str) -> bool:
    return any(k in text for k in ANALYSIS_KEYWORDS)


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


async def do_analysis(text: str, db_id: str, account: str) -> str:
    today = date.today()
    year, month = today.year, today.month

    # Detect if asking about last month
    if "上個月" in text or "上月" in text:
        month -= 1
        if month == 0:
            month, year = 12, year - 1

    pages = fetch_month_data(db_id, year, month)
    context = build_analysis_context(pages, account)
    month_label = f"{year}年{month}月"

    msg = claude.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        system=f"""你是一個家庭財務分析助理，用繁體中文回答，語氣親切簡潔。
分析以下 {month_label} 的帳務資料，回答用戶的問題。
格式：條列式，重點數字加粗，最後給一句建議。
帳務資料（日期 | 項目 | 類型 | 類別 | 金額 | 付款方式）：
{context}""",
        messages=[{"role": "user", "content": text}],
    )
    return f"📊 **{month_label} 帳務分析**\n\n{msg.content[0].text}"


# ── Claude parsers ──────────────────────────────────────────────

def parse_transaction(text: str, account: str) -> dict | None:
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
    raw = re.sub(r"```[a-z]*\n?", "", msg.content[0].text.strip()).replace("```", "").strip()
    data = json.loads(raw)
    return None if "error" in data else data


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


# ── Free chat ───────────────────────────────────────────────────

async def free_chat(text: str, account: str, channel_name: str) -> str:
    account_label = {"wife": "老婆", "hub": "老公", "joint": "公帳"}.get(account, "")
    history = chat_history.setdefault(channel_name, [])
    history.append({"role": "user", "content": text})
    # keep last 20 turns to avoid token overflow
    if len(history) > 20:
        history[:] = history[-20:]
    msg = claude.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        system=f"""你是「富貴號」——楊家的私人財務顧問兼生活助理，個性聰明、親切、有點幽默。
你在 {account_label} 的記帳頻道服務，但你的能力不限於記帳。

你可以：
- 幫忙分析家庭財務狀況、規劃預算、解讀對帳單
- 解答任何生活問題、計算、查詢、翻譯
- 提供理財建議（台灣視角：定存、ETF、美股、保險等）
- 閒聊、說笑話、回答各種奇怪問題

回答原則：
- 用繁體中文，語氣自然像朋友聊天
- 該詳細就詳細，不要為了簡潔而省掉重要資訊
- 數字、步驟用條列式；純聊天就自然回應
- 有記憶，能接續上下文對話""",
        messages=history,
    )
    reply = msg.content[0].text
    history.append({"role": "assistant", "content": reply})
    return reply


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


def is_cancel_command(text: str) -> bool:
    return any(k in text for k in ["取消", "刪除上一筆", "刪掉上一筆", "undo", "cancel", "刪除", "取消上一筆"])


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

        # ── File attachment (PDF / Excel / CSV) ────────────────
        supported = [a for a in message.attachments
                     if a.filename.lower().endswith((".pdf", ".xlsx", ".xls", ".csv"))]
        if supported:
            att = supported[0]
            fname = att.filename.lower()
            password = ""
            pw_match = re.search(r"密碼[是:：\s]+(\S+)", text)
            if pw_match:
                password = pw_match.group(1)

            status_msg = await message.reply(f"📄 正在讀取 **{att.filename}**，請稍候…")
            try:
                file_bytes = await att.read()
                if fname.endswith(".pdf"):
                    transactions = parse_pdf_transactions(file_bytes, account, password)
                    file_type = "PDF"
                elif fname.endswith((".xlsx", ".xls")):
                    transactions = parse_tabular_transactions(excel_to_text(file_bytes), account)
                    file_type = "Excel"
                else:
                    transactions = parse_tabular_transactions(csv_to_text(file_bytes), account)
                    file_type = "CSV"

                if not transactions:
                    await status_msg.edit(content="⚠️ 找不到任何交易記錄，請確認檔案格式正確。")
                    return

                await status_msg.edit(content=f"📄 找到 **{len(transactions)}** 筆記錄，匯入中…")
                ok, fail = 0, 0
                for tx in transactions:
                    try:
                        page_id = write_to_notion(db_id, tx, account)
                        last_page_id[channel_name] = page_id
                        ok += 1
                    except Exception:
                        fail += 1

                summary = f"✅ {file_type} 匯入完成！成功 **{ok}** 筆"
                if fail:
                    summary += f"，失敗 {fail} 筆"
                await status_msg.edit(content=summary)

            except json.JSONDecodeError:
                await status_msg.edit(content="❌ 解析失敗，可能格式不支援\n若為加密 PDF 請附上：`密碼是XXXX`")
            except Exception as e:
                await status_msg.edit(content=f"❌ 錯誤：{e}")
            return

        if not text:
            return

        # ── Cancel / undo ────────────────────────────────────────
        if is_cancel_command(text):
            page_id = last_page_id.get(channel_name)
            if not page_id:
                await message.reply("⚠️ 找不到上一筆記錄，無法取消。")
                return
            try:
                delete_notion_page(page_id)
                last_page_id.pop(channel_name, None)
                await message.reply("🗑️ 上一筆記錄已刪除！")
            except Exception as e:
                await message.reply(f"❌ 刪除失敗：{e}")
            return

        # ── Analysis ─────────────────────────────────────────────
        if is_analysis_request(text):
            try:
                result = await do_analysis(text, db_id, account)
                await message.reply(result)
            except Exception as e:
                await message.reply(f"❌ 分析失敗：{e}")
            return

        # ── Normal transaction or free chat ─────────────────────
        try:
            data = parse_transaction(text, account)
            if data is None:
                # Not a transaction — treat as free chat with Claude
                reply = await free_chat(text, account, channel_name)
                await message.reply(reply)
                return
            page_id = write_to_notion(db_id, data, account)
            last_page_id[channel_name] = page_id
            await message.reply(format_confirm(data, account))
        except json.JSONDecodeError:
            await message.reply("❌ 解析失敗，請稍後再試")
        except Exception as e:
            await message.reply(f"❌ 錯誤：{e}")


bot.run(DISCORD_TOKEN)
