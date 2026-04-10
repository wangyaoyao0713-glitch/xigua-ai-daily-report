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


ATTACHMENT_KEYWORDS = {
    "中课包-用户行为": "中课包-用户行为",
}


WEBHOOK_URL = os.environ.get("WEBHOOK_URL", "")
API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
API_BASE = "https://wolfai.top/v1"


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
