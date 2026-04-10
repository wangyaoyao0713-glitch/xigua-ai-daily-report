"""
西瓜创客 AI业务线 · 销售日报自动化
流程：企业微信邮箱 → 读Excel附件 → Claude API分析 → 群机器人推送
"""

import imaplib
import email
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

CARD_SUBJECTS = {
    "中课包-用户行为": "中课包-用户行为",
    "目标达成率": "目标达成率",
}

WEBHOOK_URL = os.environ.get("WEBHOOK_URL", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

def get_campaign_day(open_date, today):
    return (today - open_date).days + 1

def get_weekday_label(today):
    weekdays = ["周一", "周二", "周三", "周四", "周五", "周六", "周天"]
    return weekdays[today.weekday()]

DAILY_MILESTONES = {
    "周三": {"指标": "登录率", "渠道目标": "20%-30%", "进校目标": "30%",
             "动作": ["①消息日清日结", "②12:00之后进量消息10分钟之内回复", "③群发节点19:00/20:30"]},
    "周四": {"指标": "登录率", "渠道目标": "30%-40%", "进校目标": "30%-40%",
             "动作": ["①消息日清日结", "②12:00之后进量消息10分钟之内回复", "③群发节点19:00/20:30"]},
    "周五": {"指标": "登录率", "渠道目标": "40%-50%", "进校目标": "50%-55%",
             "动作": ["①消息日清日结", "②12:00之后进量消息10分钟之内回复", "③群发节点19:00/20:30/21:30"]},
    "周六": {"指标": "登录率", "渠道目标": "50%-55%", "进校目标": "55%-60%",
             "动作": ["①群发节点7:00/19:00/20:30/21:30", "②消息日清日结", "③12:00之后消息10分钟内回复"]},
    "周天": {"指标": "登录率30%/深访[小灶课]70%", "渠道目标": "55%-60%登录/10个有效深访", "进校目标": "60%-70%登录/10个有效深访",
             "动作": ["①8:00发送家访邀约话术，邀约时间到19:00", "②群发登录节点19:30/20:30/21:30"]},
    "周一": {"指标": "登录率20%/深访80%", "渠道目标": "55%-60%登录/15个有效深访", "进校目标": "70%-75%登录/10个有效深访",
             "动作": ["①13:00发送家访邀约话术，邀约时间到21:00", "②群发登录话术节点15:00/19:30/20:30"]},
    "周二": {"指标": "登录率10%/深访90%", "渠道目标": "60%-65%登录/15个有效深访", "进校目标": "75%-80%登录/10个有效深访",
             "动作": ["①13:00发送家访邀约话术，邀约时间到21:00", "②群发登录话术节点19:30/20:30"]},
}

def fetch_excel_attachments(subject_keywords, days_back=1):
    results = {}
    mail = imaplib.IMAP4_SSL(EMAIL_HOST, EMAIL_PORT)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("INBOX")
    since_date = (datetime.date.today() - datetime.timedelta(days=days_back)).strftime("%d-%b-%Y")
    for keyword in subject_keywords:
        _, msg_ids = mail.search(None, f'(SINCE {since_date} SUBJECT "{keyword}")')
        ids = msg_ids[0].split()
        if not ids:
            print(f"[警告] 未找到主题含 {keyword} 的邮件")
            continue
        _, msg_data = mail.fetch(ids[-1], "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])
        for part in msg.walk():
            if part.get_content_type() in [
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel", "application/octet-stream"
            ]:
                filename = part.get_filename() or ""
                if filename.endswith(".xlsx") or filename.endswith(".xls"):
                    payload = part.get_payload(decode=True)
                    df = pd.read_excel(BytesIO(payload))
                    results[keyword] = df
                    print(f"[成功] 读取附件: {filename}")
    mail.logout()
    return results

def parse_behavior_df(df, target_date):
    df.iloc[:, 0] = df.iloc[:, 0].ffill()
    df.columns = df.columns.str.strip()
    mask = df.iloc[:, 0].astype(str).str.contains(target_date[:10], na=False)
    period_df = df[mask].copy()
    if period_df.empty:
        period_df = df[df.iloc[:, 0] == df.iloc[:, 0].dropna().iloc[-1]]
    records = []
    for _, row in period_df.iterrows():
        records.append({
            "进校/非进校": row.get("进校/非进校", ""),
            "高低龄": row.get("高低龄", ""),
            "有效好友数": row.get("销售 - 有效好友", 0),
            "登录率": f"{row.get('登录率', 0):.1%}",
            "课前深访率": f"{row.get('课前深访率', 0):.1%}",
            "课前直播参与率": f"{row.get('课前直播参与率', 0):.1%}",
            "lec1完课率": f"{row.get('lec1完课率', 0):.1%}",
            "lec2完课率": f"{row.get('lec2完课率', 0):.1%}",
            "lec3完课率": f"{row.get('lec3完课率', 0):.1%}",
            "DM单点击率": f"{row.get('DM单点击率（UV）', 0):.1%}",
            "好友转化率": f"{row.get('销售好友转化率-营期内', 0):.1%}",
        })
    return records

def analyze_with_claude(data, today, open_date):
    client = Anthropic(api_key=ANTHROPIC_API_KEY)
    weekday = get_weekday_label(today)
    campaign_day = get_campaign_day(open_date, today)
    milestone = DAILY_MILESTONES.get(weekday, {})
    prompt = f"""你是西瓜创客AI业务线的销售运营分析师。今天是{today.strftime('%m月%d日')}（{weekday}），当前营期第{campaign_day}天。

今日数据（当前营期，按进校/非进校 × 高低龄分组）：
{json.dumps(data, ensure_ascii=False, indent=2)}

今日关键指标：{milestone.get('指标', '')}
Milestone目标 - 渠道：{milestone.get('渠道目标', '')} / 进校：{milestone.get('进校目标', '')}

今日必做动作：
{chr(10).join(milestone.get('动作', []))}

请生成销售日报（给管理者，13:00午会前对齐）：
1. 开头一句话说整体进展
2. 分组数据对比milestone，标注⚠️（低于目标）或✅（达标）
3. 今日必做动作直接列出
4. 一句重点提示

风格简洁，不超过300字。"""
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
    if result.get("errcode") == 0:
        print("[成功] 日报已推送至企业微信")
    else:
        print(f"[失败] 推送失败: {result}")

def main():
    today = datetime.date.today()
    print(f"[开始] {today} 日报生成")
    attachments = fetch_excel_attachments(list(CARD_SUBJECTS.keys()), days_back=1)
    if not attachments:
        send_to_wecom(f"【AI日报】{today.strftime('%m月%d日')} ⚠️ 数据获取失败，请手动检查BI看板")
        return
    open_date = today - datetime.timedelta(days=today.weekday() + 2)
    all_data = {}
    for keyword, df in attachments.items():
        all_data[keyword] = parse_behavior_df(df, today.strftime("%Y-%m-%d"))
    report = analyze_with_claude(all_data, today, open_date)
    header = f"【AI销售日报】{today.strftime('%m月%d日')} · {get_weekday_label(today)}\n{'='*25}\n"
    send_to_wecom(header + report)

if __name__ == "__main__":
    main()
