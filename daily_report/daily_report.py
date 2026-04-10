import imaplib
import email
import email.header
import os
import json
import datetime
import requests
import pandas as pd
from io import BytesIO
from anthropic import Anthropic

EMAIL_HOST = "imap.exmail.qq.com"
EMAIL_PORT = 993
EMAIL_USER = os.environ.get("EMAIL_USER", "")
EMAIL_PASS = os.environ.get("EMAIL_PASS", "")

# 邮件主题关键词 -> 卡片名
CARD_SUBJECTS = {
    "中课包-用户行为": "中课包-用户行为",
}

WEBHOOK_URL = os.environ.get("WEBHOOK_URL", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

def get_weekday_label(today):
    weekdays = ["周一", "周二", "周三", "周四", "周五", "周六", "周天"]
    return weekdays[today.weekday()]

def get_campaign_day(open_date, today):
    return (today - open_date).days + 1

DAILY_MILESTONES = {
    "周三": {"指标": "登录率", "渠道目标": "20%-30%", "进校目标": "30%",
             "动作": ["①消息日清日结", "②12:00之后进量消息10分钟之内回复", "③群发节点19:00/20:30"]},
    "周四": {"指标": "登录率", "渠道目标": "30%-40%", "进校目标": "30%-40%",
             "动作": ["①消息日清日结", "②12:00之后进量消息10分钟之内回复", "③群发节点19:00/20:30"]},
    "周五": {"指标": "登录率", "渠道目标": "40%-50%", "进校目标": "50%-55%",
             "动作": ["①消息日清日结", "②12:00之后进量消息10分钟之内回复", "③群发节点19:00/20:30/21:30"]},
    "周六": {"指标": "登录率", "渠道目标": "50%-55%", "进校目标": "55%-60%",
             "动作": ["①群发节点7:00/19:00/20:30/21:30", "②消息日清日结", "③12:00之后消息10分钟内回复"]},
    "周天": {"指标": "登录率30%/深访70%", "渠道目标": "55%-60%登录/10个有效深访", "进校目标": "60%-70%登录/10个有效深访",
             "动作": ["①8:00发送家访邀约话术，邀约时间到19:00", "②群发登录节点19:30/20:30/21:30"]},
    "周一": {"指标": "登录率20%/深访80%", "渠道目标": "55%-60%登录/15个有效深访", "进校目标": "70%-75%登录/10个有效深访",
             "动作": ["①13:00发送家访邀约话术，邀约时间到21:00", "②群发登录话术节点15:00/19:30/20:30"]},
    "周二": {"指标": "登录率10%/深访90%", "渠道目标": "60%-65%登录/15个有效深访", "进校目标": "75%-80%登录/10个有效深访",
             "动作": ["①13:00发送家访邀约话术，邀约时间到21:00", "②群发登录话术节点19:30/20:30"]},
}

def decode_subject(raw_subject):
    parts = email.header.decode_header(raw_subject)
    result = ""
    for part, charset in parts:
        if isinstance(part, bytes):
            result += part.decode(charset or "utf-8", errors="replace")
        else:
            result += part
    return result

def fetch_excel_attachments(days_back=2):
    results = {}
    mail = imaplib.IMAP4_SSL(EMAIL_HOST, EMAIL_PORT)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("INBOX")
    
    since_date = (datetime.date.today() - datetime.timedelta(days=days_back)).strftime("%d-%b-%Y")
    print(f"[搜索] 查找 {since_date} 之后的邮件")
    
    _, msg_ids = mail.search(None, f"SINCE {since_date}")
    ids = msg_ids[0].split()
    print(f"[搜索] 找到 {len(ids)} 封邮件")
    
    for keyword, subject_cn in CARD_SUBJECTS.items():
        found_id = None
        for uid in reversed(ids[-50:]):
            _, msg_data = mail.fetch(uid, "(BODY[HEADER.FIELDS (SUBJECT)])")
            raw_header = msg_data[0][1]
            msg_header = email.message_from_bytes(raw_header)
            subj = decode_subject(msg_header.get("Subject", ""))
            print(f"[检查] 主题: {subj}")
            if subject_cn in subj:
                found_id = uid
                print(f"[匹配] 找到: {subj}")
                break
        
        if not found_id:
            print(f"[警告] 未找到含 {subject_cn} 的邮件")
            continue
        
        _, msg_data = mail.fetch(found_id, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])
        for part in msg.walk():
            fn = part.get_filename() or ""
            if fn.endswith(".xlsx") or fn.endswith(".xls"):
                payload = part.get_payload(decode=True)
                df = pd.read_excel(BytesIO(payload))
                results[keyword] = {"name": subject_cn, "df": df}
                print(f"[成功] 读取附件: {fn} ({len(df)}行)")
                break
    
    mail.logout()
    return results

def parse_behavior_df(df):
    df.iloc[:, 0] = df.iloc[:, 0].ffill()
    df.columns = df.columns.str.strip()
    latest = df.iloc[:, 0].dropna().iloc[-1]
    period_df = df[df.iloc[:, 0] == latest].copy()
    records = []
    for _, row in period_df.iterrows():
        r = {"分组": f"{row.get('进校/非进校','')} {row.get('高低龄','')}".strip()}
        for col in ["销售 - 有效好友", "登录率", "课前深访率", "lec1完课率", "lec2完课率", "lec3完课率", "DM单点击率（UV）", "销售好友转化率-营期内"]:
            val = row.get(col, 0) or 0
            if col == "销售 - 有效好友":
                r[col] = int(val)
            else:
                r[col] = f"{float(val):.1%}" if val else "0.0%"
        records.append(r)
    return records

def analyze_with_claude(data, today, open_date):
    client = Anthropic(api_key=ANTHROPIC_API_KEY)
    weekday = get_weekday_label(today)
    campaign_day = get_campaign_day(open_date, today)
    milestone = DAILY_MILESTONES.get(weekday, {})
    prompt = f"""你是西瓜创客AI业务线的销售运营分析师。今天是{today.strftime('%m月%d日')}（{weekday}），营期第{campaign_day}天。

数据（最新营期，按进校/非进校 × 高低龄分组）：
{json.dumps(data, ensure_ascii=False, indent=2)}

今日核心指标：{milestone.get('指标','')}
目标 - 渠道：{milestone.get('渠道目标','')} / 进校：{milestone.get('进校目标','')}

今日必做动作：
{chr(10).join(milestone.get('动作',[]))}

请生成销售日报（管理者13:00午会前用）：
1. 一句话说整体进展
2. 各分组对比目标，⚠️预警（低于目标）或✅达标
3. 今日必做动作直接列出
4. 一句重点提示

不超过300字，简洁直接。"""
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        messages=[{"role": "user", "content": prompt}]
    )
    return response.content[0].text

def send_to_wecom(content):
    payload = {"msgtype": "text", "text": {"content": content}}
    resp = requests.post(WEBHOOK_URL, json=payload, timeout=10)
    result = resp.json()
    print("[推送成功]" if result.get("errcode") == 0 else f"[推送失败] {result}")

def main():
    today = datetime.date.today()
    print(f"[开始] {today} 日报生成")
    attachments = fetch_excel_attachments(days_back=2)
    if not attachments:
        send_to_wecom(f"【AI日报】{today.strftime('%m月%d日')} ⚠️ 未收到BI数据邮件，请手动检查")
        return
    all_data = {}
    for key, item in attachments.items():
        all_data[item["name"]] = parse_behavior_df(item["df"])
    open_date = today - datetime.timedelta(days=today.weekday() + 2)
    report = analyze_with_claude(all_data, today, open_date)
    header = f"【AI销售日报】{today.strftime('%m月%d日')} · {get_weekday_label(today)}\n{'='*25}\n"
    send_to_wecom(header + report)

if __name__ == "__main__":
    main()
