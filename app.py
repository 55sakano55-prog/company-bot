import os
import tempfile
from flask import Flask, request, abort
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import Configuration, ApiClient, MessagingApi, ReplyMessageRequest, TextMessage
from linebot.v3.webhooks import MessageEvent, TextMessageContent
from invoice import build_invoice
from payslip import build_payslip

app = Flask(__name__)
CHANNEL_SECRET = os.environ["LINE_CHANNEL_SECRET"]
CHANNEL_ACCESS_TOKEN = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]
configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)
state = {}

@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"

@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    user_id = event.source.user_id
    text = event.message.text.strip()
    triggers = ["作成をお願いします", "お願いします", "請求書", "給料明細", "作成"]
    if user_id not in state and any(t in text for t in triggers):
        state[user_id] = {"flow": None, "step": 0, "data": {}}
        reply(event, menu_message())
        return
    if user_id not in state:
        return
    s = state[user_id]
    if s['flow'] is None:
        if text in ["1", "①", "請求書", "請求書作成"]:
            s["flow"] = "invoice"
            s["step"] = 1
            reply(event, "【請求書作成】\n①宛先（会社名）を教えてください。\n例：〇〇建設株式会社")
        elif text in ["2", "②", "給料明細", "給料明細作成"]:
            s["flow"] = "payslip"
            s["step"] = 1
            reply(event, "【給料明細作成】\n①氏名を教えてください。")
        elif text in ["0", "キャンセル"]:
            del state[user_id]
            reply(event, "キャンセルしました。")
        else:
            reply(event, "「1」または「2」で選択してください。\n1：請求書\n2：給料明細\n0：キャンセル")
        return
    if s["flow"] == "invoice":
        handle_invoice_flow(event, s, text)
    elif s["flow"] == "payslip":
        handle_payslip_flow(event, s, text)

INVOICE_STEPS = [
    (1, "①宛先（会社名）を教えてください。\n例：〇〇建設株式会社", "client"),
    (2, "②発行日を教えてください。\n例：2026年4月3日", "date"),
    (3, "③現場名を教えてください。", "site"),
    (4, "④支払期限を教えてください。\n例：2026年4月30日", "due_date"),
    (5, "⑤工事内容を入力してください。\n形式：工事名,数量,単価\n（終わったら「完了」と送信）", "items"),
    (6, "⑥備考があれば入力してください。（なければ「なし」）", "note"),
]

def handle_invoice_flow(event, s, text):
    step = s['step']
    if step == 5:
        if text == "完了":
            if not s["data"].get("items"):
                reply(event, "工事内容が入力されていません。")
                return
            s["step"] = 6
            reply(event, INVOICE_STEPS[5][1])
        else:
            parts = [p.strip() for p in text.replace("、", ",").split(",")]
            if len(parts) != 3:
                reply(event, "形式が正しくありません。\n工事名,数量,単価 で入力してください。")
                return
            try:
                name, qty, price = parts[0], float(parts[1]), float(parts[2])
            except ValueError:
                reply(event, "数量・単価は数字で入力してください。")
                return
            s["data"].setdefault("items", []).append({"name": name, "qty": qty, "price": price})
            count = len(s["data"]["items"])
            reply(event, f"{count}件目を登録しました✅\n追加がある場合は続けて入力、終わったら「完了」と送信してください。")
        return
    if step == 6:
        s["data"]["note"] = "" if text == "なし" else text
        reply(event, "請求書を作成中です...⏳")
        try:
            filepath = build_invoice(s["data"])
            send_file(event, filepath)
        except Exception as e:
            reply(event, f"エラーが発生しました：{e}")
        del state[event.source.user_id]
        return
    key = INVOICE_STEPS[step - 1][2]
    s["data"][key] = text
    s["step"] += 1
    reply(event, INVOICE_STEPS[s["step"] - 1][1])

PAYSLIP_STEPS = [
    (1, "①氏名を教えてください。", "name"),
    (2, "②社員番号を教えてください。", "emp_no"),
    (3, "③対象年月を教えてください。\n例：2026年4月", "period"),
    (4, "④就業日数を入力してください。", "work_days"),
    (5, "⑤出勤日数を入力してください。", "attend_days"),
    (6, "⑥労働時間を入力してください。", "work_hours"),
    (7, "⑦残業時間を入力してください。（なければ0）", "overtime"),
    (8, "⑧休日出勤日数を入力してください。（なければ0）", "holiday_work"),
    (9, "⑨有給消化日数を入力してください。（なければ0）", "paid_leave"),
    (10, "⑩基本給を入力してください。\n例：216000", "base_salary"),
    (11, "⑪残業手当を入力してください。（なければ0）", "overtime_pay"),
    (12, "⑫深夜勤務手当を入力してください。（なければ0）", "night_pay"),
    (13, "⑬役員報酬を入力してください。（なければ0）", "director_pay"),
    (14, "⑭厚生年金保険料を入力してください。", "pension"),
    (15, "⑮雇用保険料を入力してください。", "employment_ins"),
    (16, "⑯健康保険料を入力してください。", "health_ins"),
    (17, "⑰所得税を入力してください。", "income_tax"),
    (18, "⑱家賃控除があれば入力してください。（なければ0）", "rent"),
]

def handle_payslip_flow(event, s, text):
    step = s['step']
    idx = step - 1
    if step >= 4:
        try:
            val = float(text)
        except ValueError:
            reply(event, "数字で入力してください。")
            return
        s["data"][PAYSLIP_STEPS[idx][2]] = val
    else:
        s["data"][PAYSLIP_STEPS[idx][2]] = text
    if step == len(PAYSLIP_STEPS):
        reply(event, "給料明細を作成中です...⏳")
        try:
            filepath = build_payslip(s["data"])
            send_file(event, filepath)
        except Exception as e:
            reply(event, f"エラーが発生しました：{e}")
        del state[event.source.user_id]
        return
    s["step"] += 1
    reply(event, PAYSLIP_STEPS[s["step"] - 1][1])

def reply(event, text_or_msg):
    with ApiClient(configuration) as api_client:
        api = MessagingApi(api_client)
        if isinstance(text_or_msg, str):
            msg = TextMessage(type='text', text=text_or_msg)
        else:
            msg = text_or_msg
        api.reply_message(ReplyMessageRequest(reply_token=event.reply_token, messages=[msg]))

def send_file(event, filepath):
    base_url = os.environ.get('BASE_URL', '')
    fname = os.path.basename(filepath)
    url = f'{base_url}/files/{fname}'
    reply(event, f'📎 ファイルのダウンロードはこちら：\n{url}')

def menu_message():
    return ('何を作成しますか？番号を送信してください。\n\n1️⃣  請求書\n2️⃣  給料明細\n\n0：キャンセル')

TEMP_DIR = tempfile.gettempdir()

@app.route('/files/<filename>')
def serve_file(filename):
    from flask import send_from_directory
    return send_from_directory(TEMP_DIR, filename, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
