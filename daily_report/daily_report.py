import imaplib
import email
import email.header
import os
import json
import datetime
import requests
import pandas as pd
from io import BytesIO
from openai import OpenAI

EMAIL_HOST = "imap.exmail.qq.com"
EMAIL_PORT = 993
EMAIL_USER = os.environ.get("EMAIL_USER", "")
EMAIL_PASS = os.environ.get("EMAIL_PASS", "")
WEBHOOK_URL = os.environ.get("WEBHOOK_URL", "")
API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
API_BASE = "https://wolfai.top/v1"

ATTACHMENT_KEYWORDS = {"zhongkebao": "中课包-用户行为"}

MILESTONES = {
    "周一": {"指标": "登录率20%/深访80%", "渠道": "55%-60%登录/15个有效深访", "进校": "70%-75%登录/10个有效深访",
             "动作": ["①13:00发送家访邀约话术，邀约时间到21:00", "②群发登录话术节点15:00/19:30/20:30"]},
    "周二": {"指标": "登录率10%/深访90%", "渠道": "60%-65%登录/15个有效深访", "进校": "75%-80%登录/10个有效深访",
             "动作": ["①13:00发送家访邀约话术，邀约时间到21:00", "②群发登录话术节点19:30/20:30"]},
    "周三": {"指标": "登录率", "渠道": "20%-30%", "进校": "30%",
             "动作": ["①消息日清日结", "②12:00之后进量消息10分钟之内回复", "③群发节点19:00/20:30"]},
    "周四": {"指标": "登录率", "渠道": "30%-40%", "进校": "30%-40%",
             "动作": ["①消息日清日结", "②12:00之后进量消息10分钟之内回复", "③群发节点19:00/20:30"]},
    "周五": {"指标": "登录率", "渠道": "40%-50%", "进校": "50%-55%",
             "动作": ["①消息日清日结", "②12:00之后进量消息10分钟之内回复", "③群发节点19:00/20:30/21:30"]},
    "周六": {"指标": "登录率", "渠道": "50%-55%", "进校": "55%-60%",
             "动作": ["①群发节点7:00/19:00/20:30/21:30", "②消息日清日结", "③12:00之后消息10分钟内回复"]},
    "周天": {"指标": "登录率30%/深访70%", "渠道": "55%-60%登录/10个有效深访", "进校": "60%-70%登录/10个有效深访",
             "动作": ["①8:00发送家访邀约话术，邀约时间到19:00", "②群发登录节点19:30/20:30/21:30"]},
}

def get_weekday(today):
    return ["周一","周二","周三","周四","周五","周六","周天"][today.weekday()]

def decode_filename(raw):
    parts = email.header.decode_header(raw)
    result = ""
    for part, charset in parts:
        if isinstance(part, bytes):
            result += part.decode(charset or "utf-8", errors="replace")
        else:
            result += str(part)
    return result

def fetch_excel(days_back=2):
    results = {}
    mail = imaplib.IMAP4_SSL(EMAIL_HOST, EMAIL_PORT)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("INBOX")
    since = (datetime.date.today() - datetime.timedelta(days=days_back)).strftime("%d-%b-%Y")
    _, ids = mail.search(None, f"SINCE {since}")
    ids = ids[0].split()
    print(f"[搜索] {len(ids)} 封邮件")
    for uid in reversed(ids[-50:]):
        _, data = mail.fetch(uid, "(RFC822)")
        msg = email.message_from_bytes(data[0][1])
        for part in msg.walk():
            raw_fn = part.get_filename()
            if not raw_fn:
                continue
            fn = decode_filename(raw_fn)
            for key, name in ATTACHMENT_KEYWORDS.items():
                if name in fn and fn.endswith((".xlsx", ".xls")) and key not in results:
                    df = pd.read_excel(BytesIO(part.get_payload(decode=True)))
                    results[key] = {"name": name, "df": df}
                    print(f"[成功] {fn} ({len(df)}行)")
    mail.logout()
    return results

def parse_df(df):
    df.iloc[:, 0] = df.iloc[:, 0].ffill()
    df.columns = df.columns.str.strip()
    latest = df.iloc[:, 0].dropna().iloc[-1]
    rows = df[df.iloc[:, 0] == latest].copy()
    records = []
    for _, row in rows.iterrows():
        r = {"分组": f"{row.get('进校/非进校','')} {row.get('高低龄','')}".strip()}
        for col in ["销售 - 有效好友","登录率","课前深访率","lec1完课率","lec2完课率","lec3完课率","DM单点击率（UV）","销售好友转化率-营期内"]:
            val = row.get(col, 0) or 0
            r[col] = int(val) if col == "销售 - 有效好友" else (f"{float(val):.1%}" if val else "0.0%")
        records.append(r)
    return records

def call_ai(data, today):
    client = OpenAI(api_key=API_KEY, base_url=API_BASE)
    wd = get_weekday(today)
    ms = MILESTONES.get(wd, {})
    prompt = f"""西瓜创客AI销售日报 - {today.strftime('%m月%d日')}（{wd}）

数据：
{json.dumps(data, ensure_ascii=False, indent=2)}

今日指标：{ms.get('指标','')}
目标 渠道：{ms.get('渠道','')} / 进校：{ms.get('进校','')}

今日必做：
{chr(10).join(ms.get('动作',[]))}

生成日报（管理者13:00午会用）：
1. 一句话整体进展
2. 各分组对比目标，加预警标注
3. 今日必做动作
4. 一句重点提示

不超过300字。"""
    resp = client.chat.completions.create(
        model="claude-3-5-sonnet-20241022",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=1000
    )
    return resp.choices[0].message.content

def push(content):
    r = requests.post(WEBHOOK_URL, json={"msgtype":"text","text":{"content":content}}, timeout=10).json()
    print("[推送成功]" if r.get("errcode") == 0 else f"[推送失败] {r}")

def main():
    today = datetime.date.today()
    print(f"[开始] {today}")
    data = fetch_excel()
    if not data:
        push(f"【AI日报】{today.strftime('%m月%d日')} 未收到BI数据邮件，请手动检查")
        return
    all_data = {item["name"]: parse_df(item["df"]) for item in data.values()}
    report = call_ai(all_data, today)
    header = f"【AI销售日报】{today.strftime('%m月%d日')} · {get_weekday(today)}\n{'='*25}\n"
    push(header + report)

if __name__ == "__main__":
    main()
