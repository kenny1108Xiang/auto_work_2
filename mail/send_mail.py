import smtplib
import os
from email.mime.text import MIMEText
from email.header import Header
from dotenv import load_dotenv
import logging
import asyncio
from concurrent.futures import ThreadPoolExecutor


def render_email_html(summary_data: dict) -> str:
    """
    產生精美、相容度高的郵件 HTML（無 emoji）。
    - 專業 UX/UI 配色與排版
    - 高對比，深/淺色模式皆清楚
    - 仍以 table 為骨架提高相容度
    """
    submitted_days = summary_data.get("submitted_days", [])
    submitted_days_str = "、".join(submitted_days) if submitted_days else "（無）"

    reasons = summary_data.get("reasons", {}) or {}
    all_success = bool(summary_data.get("all_success", False))
    successful_day_names = summary_data.get("successful_day_names", [])
    failed_tasks = summary_data.get("failed_tasks", [])

    preheader = f"本次提交：{len(successful_day_names)} 成功, {len(failed_tasks)} 失敗"

    # 理由列表 HTML
    reasons_items = []
    if reasons.get("sat"):
        reasons_items.append(
            f"<li><span class='chip chip-day'>星期六</span><span class='reason'>{reasons['sat']}</span></li>"
        )
    if reasons.get("sun"):
        reasons_items.append(
            f"<li><span class='chip chip-day'>星期日</span><span class='reason'>{reasons['sun']}</span></li>"
        )
    reasons_html = ""
    if reasons_items:
        reasons_html = """
        <tr>
          <td class="section">
            <div class="section-title">請假理由</div>
            <ul class="reasons">
              {items}
            </ul>
          </td>
        </tr>
        """.format(items="\n".join(reasons_items))

    # 執行結果 HTML
    if all_success:
        result_html = """
        <div class="status">
          <span class="badge success">全數成功</span>
          <p class="hint">已成功提交所有指定表單。</p>
        </div>
        """
    else:
        # 成功部分
        success_items = []
        if successful_day_names:
            for day_name in successful_day_names:
                success_items.append(
                    f"<li class='success-item'>{day_name}</li>"
                )
            success_list_html = "<ul class='success-list'>{}</ul>".format("\n".join(success_items))
        else:
            success_list_html = "<p class='hint'>無</p>"
        
        # 失敗部分
        failed_items = []
        for task in failed_tasks:
            day_name = task.get('day_name', '未知表單')
            status = task.get('status', 'unknown')
            
            if status == 'closed':
                reason_text = "表單已關閉或名額已滿"
            elif status == 'prep_failed':
                reason_text = "資料準備失敗 (URL/欄位錯誤)"
            elif status == 'submission_failed':
                reason_text = "提交失敗 (網路或伺服器錯誤)"
            else:
                reason_text = "未知失敗"
            
            failed_items.append(
                f"<li class='failure-item'><span class='chip chip-day'>{day_name}</span><span class='failure-reason'>{reason_text}</span></li>"
            )
        
        failed_list_html = "<ul class='failures-list'>{}</ul>".format("\n".join(failed_items))

        result_html = f"""
        <div class="status">
          <span class="badge failure">未全數成功</span>
          <div class="result-section">
            <p class="result-label">成功部分：</p>
            {success_list_html}
          </div>
          <div class="result-section" style="margin-top:16px;">
            <p class="result-label">失敗部分：</p>
            {failed_list_html}
            <p class="hint" style="margin-top:12px;">請查看程式的日誌輸出以了解詳細錯誤原因。</p>
          </div>
        </div>
        """

    return f"""\
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="x-apple-disable-message-reformatting">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="color-scheme" content="light dark">
  <title>表單提交總結報告</title>
  <style>
    /* -------- 基礎重置 -------- */
    body,table,td,p,span,a {{ margin:0; padding:0; }}
    img {{ border:0; line-height:100%; outline:none; text-decoration:none; max-width:100%; }}
    a {{ text-decoration:none; }}
    
    /* Gmail 特殊重置 */
    u + .body {{ background:#F6F8FC; }}
    * {{ -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; }}

    /* -------- 淺色主題（預設） -------- */
    body {{
      background:#F6F8FC;
      color:#111827; /* 高對比正文 */
      font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,"Noto Sans TC","PingFang TC","Microsoft JhengHei",sans-serif;
      line-height:1.65;
    }}
    .wrapper {{ width:100%; padding:28px 12px; }}
    .container {{
      width:100%; max-width:720px; margin:0 auto; background:#FFFFFF;
      border:1px solid #E5E7EB; border-radius:12px; overflow:hidden;
    }}
    .header {{
      padding:24px 28px;
      background:#FFFFFF;
      border-bottom:1px solid #E5E7EB;
    }}
    .eyebrow {{
      font-size:12px; letter-spacing:.08em; text-transform:uppercase;
      color:#2563EB; font-weight:700; margin-bottom:6px;
    }}
    .title {{
      font-size:22px; font-weight:800; color:#0F172A; /* 更深的標題色 */
    }}
    .content {{ padding:8px 28px 28px 28px; }}
    .section {{ padding:18px 0; }}
    .section + .section {{ border-top:1px solid #EEF2F7; }}
    .section-title {{
      font-size:15px; font-weight:800; color:#0F172A; margin-bottom:10px;
      padding-left:10px; border-left:3px solid #2563EB; /* 清楚的視覺錨點 */
    }}

    /* 資訊卡：提高對比 */
    .card {{
      background:#F9FAFB;
      border:1px solid #E5E7EB;
      border-radius:10px;
      padding:14px 16px;
    }}
    .meta p {{ margin:0 0 6px 0; color:#111827; }}
    .meta b {{ color:#0F172A; }}

    /* 膠囊標籤、狀態徽章 - Gmail 優化版 */
    .chip {{
      display:inline-block !important; /* Gmail 需要 !important */
      padding:8px 12px !important;
      border-radius:999px !important;
      font-size:12px !important;
      font-weight:700 !important;
      border:1px solid #C7D2FE !important;
      background:#EEF2FF !important;
      color:#1E3A8A !important;
      margin-right:8px !important;
      vertical-align:middle !important;
      text-align:center !important;
      white-space:nowrap !important;
      line-height:16px !important; /* 明確指定行高 */
      min-height:32px !important; /* 確保有足夠高度 */
      box-sizing:border-box !important;
    }}
    .chip-day {{ min-width:56px !important; }}

    .badge {{
      display:inline-block; padding:7px 12px; border-radius:999px;
      font-size:13px; font-weight:800; border:1px solid transparent;
    }}
    .badge.success {{ color:#065F46; background:#ECFDF5; border-color:#A7F3D0; }}
    .badge.failure {{ color:#7F1D1D; background:#FEF2F2; border-color:#FECACA; }}

    .status {{ display:block; }}
    .hint {{ color:#4B5563; font-size:13px; margin-top:6px; }}

    /* 理由列表：條列與留白 - Gmail 優化版 */
    .reasons {{ list-style:none !important; padding-left:0 !important; margin:0 !important; }}
    .reasons li {{
      display:block !important; /* Gmail 用 block 更穩定 */
      background:#F9FAFB !important;
      border:1px solid #E5E7EB !important;
      border-radius:10px !important;
      padding:12px !important;
      margin-bottom:8px !important;
      color:#111827 !important;
      overflow:hidden !important;
    }}
    .reasons .chip {{ 
      float:left !important; /* 使用 float 代替 flex */
      margin-right:10px !important;
      margin-bottom:0 !important;
    }}
    .reasons .reason {{ 
      display:block !important;
      overflow:hidden !important;
      line-height:32px !important; /* 與標籤高度一致 */
      min-height:32px !important;
    }}

    /* 失敗列表 - Gmail 優化版 */
    .failures-list {{ list-style:none !important; padding:10px 0 0 0 !important; margin:0 !important; }}
    .failure-item {{
        display:block !important; /* Gmail 用 block 更穩定 */
        background:#FEF2F2 !important;
        border:1px solid #FECACA !important;
        border-radius:10px !important;
        padding:12px !important;
        margin-bottom:8px !important;
        overflow:hidden !important;
    }}
    .failure-item .chip-day {{
        background:#FEE2E2 !important;
        border-color:#FCA5A5 !important;
        color:#991B1B !important;
        float:left !important;
        margin-right:10px !important;
    }}
    .failure-reason {{ 
      color:#991B1B !important;
      font-weight:700 !important;
      font-size:13px !important;
      line-height:32px !important; /* 與標籤高度一致 */
      min-height:32px !important;
      display:block !important;
      overflow:hidden !important;
    }}

    /* 成功列表 - Gmail 優化版 */
    .success-list {{ 
      list-style:none !important;
      padding:0 !important;
      margin:8px 0 0 0 !important;
    }}
    .success-item {{
        display:inline-block !important;
        background:#D1FAE5 !important;
        border:1px solid #6EE7B7 !important;
        border-radius:999px !important;
        padding:8px 14px !important;
        margin:0 12px 12px 0 !important; /* 右邊和下邊留間距 */
        font-size:12px !important;
        font-weight:700 !important;
        color:#065F46 !important;
        text-align:center !important;
        min-width:56px !important;
        line-height:16px !important;
        min-height:32px !important;
        box-sizing:border-box !important;
        white-space:nowrap !important;
        vertical-align:top !important;
    }}
    .chip-success {{ background:#D1FAE5 !important; border-color:#6EE7B7 !important; color:#065F46 !important; }}

    /* 結果區塊標籤 */
    .result-section {{ margin-top:8px; }}
    .result-label {{ font-weight:700; font-size:14px; color:#0F172A; margin-bottom:8px; }}

    .footer {{
      padding:16px 28px; border-top:1px solid #E5E7EB; color:#6B7280; font-size:12px; text-align:center;
      background:#FAFAFA;
    }}

    /* -------- 深色模式：手動指定避免被客戶端自動反色影響對比 -------- */
    @media (prefers-color-scheme: dark) {{
      body {{ background:#0B0F14 !important; color:#E5E7EB !important; }}
      .container {{ background:#0F1720 !important; border-color:#1F2937 !important; }}
      .header {{ background:#0F1720 !important; border-bottom-color:#1F2937 !important; }}
      .title {{ color:#F3F4F6 !important; }}
      .section + .section {{ border-top-color:#1F2937 !important; }}
      .section-title {{ color:#F3F4F6 !important; border-left-color:#3B82F6 !important; }}
      .card {{ background:#111827 !important; border-color:#334155 !important; }}
      .meta p {{ color:#E5E7EB !important; }}
      .meta b {{ color:#FFFFFF !important; }}
      .chip {{ background:#1E3A8A !important; border-color:#3B82F6 !important; color:#93C5FD !important; }}
      .reasons li {{ background:#1F2937 !important; border-color:#374151 !important; color:#E5E7EB !important; }}
      .reasons .reason {{ color:#E5E7EB !important; }}
      .footer {{ background:#0F1720 !important; border-top-color:#1F2937 !important; color:#9CA3AF !important; }}
      .hint {{ color:#9CA3AF !important; }}
      .badge.success {{ color:#86EFAC !important; background:#052E1A !important; border-color:#14532d !important; }}
      .badge.failure {{ color:#FCA5A5 !important; background:#2A0B0B !important; border-color:#7f1d1d !important; }}
      
      .failure-item {{ background:#2A0B0B !important; border-color:#7f1d1d !important; }}
      .failure-item .chip-day {{ background:#450a0a !important; border-color:#991B1B !important; color:#FCA5A5 !important; }}
      .failure-reason {{ color:#FCA5A5 !important; }}
      
      .success-item {{ 
        background:#064E3B !important; border-color:#065F46 !important; color:#86EFAC !important;
      }}
      .chip-success {{ background:#064E3B !important; border-color:#065F46 !important; color:#86EFAC !important; }}
      
      .result-label {{ color:#F3F4F6 !important; }}
    }}

    /* -------- 小螢幕微調（手機版專屬優化） -------- */
    @media screen and (max-width:520px) {{
      .header, .content, .footer {{ padding-left:18px !important; padding-right:18px !important; }}
      .title {{ font-size:20px !important; }}
      
      /* 手機版不需要額外調整，因為已經針對 Gmail 優化 */
    }}

    /* 收件匣摘要（隱藏） */
    .preheader {{
      display:none !important; visibility:hidden; opacity:0; color:transparent; height:0; width:0;
      overflow:hidden; mso-hide:all;
    }}
  </style>
</head>
<body>
  <span class="preheader">{preheader}</span>
  <table role="presentation" class="wrapper" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td align="center">
        <table role="presentation" class="container" cellpadding="0" cellspacing="0" width="100%">
          <tr>
            <td class="header">
              <div class="eyebrow">Google Form Auto-Filler</div>
              <div class="title">表單提交總結報告</div>
            </td>
          </tr>
          <tr>
            <td class="content">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td class="section">
                    <div class="section-title">提交詳情</div>
                    <div class="card meta">
                      <p><b>本次提交的表單：</b>{submitted_days_str}</p>
                    </div>
                  </td>
                </tr>
                {reasons_html}
                <tr>
                  <td class="section">
                    <div class="section-title">執行結果</div>
                    {result_html}
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td class="footer">
              這是一封自動發送的郵件，請勿直接回覆。
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>"""


async def send_email_to_single_recipient_async(recipient_email, sender_email, app_password, subject, body):
    """
    異步發送郵件給單一收件人（使用 Python 3.14 的 async 特性）。
    
    Args:
        recipient_email (str): 收件人郵箱
        sender_email (str): 寄件人郵箱
        app_password (str): 應用程式密碼
        subject (str): 郵件主旨
        body (str): 郵件內容 (HTML)
    
    Returns:
        tuple: (收件人郵箱, 是否成功)
    """
    loop = asyncio.get_event_loop()
    try:
        # 在執行緒池中執行同步的 SMTP 操作
        def send_sync():
            msg = MIMEText(body, 'html', 'utf-8')
            msg['From'] = f'自動填寫劃假表單 <{sender_email}>'
            msg['To'] = recipient_email
            msg['Subject'] = Header(subject, 'utf-8')
            
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                server.login(sender_email, app_password)
                server.sendmail(sender_email, [recipient_email], msg.as_string())
        
        await loop.run_in_executor(None, send_sync)
        logging.info(f"✓ 郵件已成功發送至：{recipient_email}")
        return (recipient_email, True)
    except Exception as e:
        logging.error(f"✗ 發送至 {recipient_email} 失敗：{e}")
        return (recipient_email, False)


def send_email_to_single_recipient(recipient_email, sender_email, app_password, subject, body):
    """
    發送郵件給單一收件人（同步版本，保留向後相容）。
    
    Args:
        recipient_email (str): 收件人郵箱
        sender_email (str): 寄件人郵箱
        app_password (str): 應用程式密碼
        subject (str): 郵件主旨
        body (str): 郵件內容 (HTML)
    
    Returns:
        tuple: (收件人郵箱, 是否成功)
    """
    try:
        msg = MIMEText(body, 'html', 'utf-8')
        msg['From'] = f'自動填寫劃假表單 <{sender_email}>'
        msg['To'] = recipient_email
        msg['Subject'] = Header(subject, 'utf-8')
        
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, app_password)
            server.sendmail(sender_email, [recipient_email], msg.as_string())
        
        logging.info(f"✓ 郵件已成功發送至：{recipient_email}")
        return (recipient_email, True)
    except Exception as e:
        logging.error(f"✗ 發送至 {recipient_email} 失敗：{e}")
        return (recipient_email, False)


def send_summary_email(summary_data):
    """
    發送總結報告郵件。
    - 如果只有 1 個收件人：同步發送
    - 如果有多個收件人：非同步並發送（加速）

    Args:
        summary_data (dict): 包含報告內容的字典。
    
    Returns:
        bool: 是否成功發送郵件。
    """
    # 從不同路徑載入環境變數
    load_dotenv(dotenv_path='mail/mail_key.env')       # 讀取 mail/ 目錄下的 mail_key.env
    load_dotenv(dotenv_path='mail/mail_settings.env')  # 讀取 mail/ 目錄下的 mail_settings.env
    
    sender_email = os.getenv("SENDER_EMAIL")
    recipient_email_str = os.getenv("RECIPIENT_EMAIL")  # 可能包含多個郵箱（逗號分隔）
    app_password = os.getenv("KEY")

    if not all([sender_email, recipient_email_str, app_password]):
        logging.error("郵件設定不完整，請檢查 mail_key.env 和 mail_settings.env 的設定。")
        if not sender_email:
            logging.error("-> 缺少 SENDER_EMAIL")
        if not recipient_email_str:
            logging.error("-> 缺少 RECIPIENT_EMAIL")
        if not app_password:
            logging.error("-> 缺少 KEY")
        return False

    # 解析收件人列表（支持逗號或分號分隔，並去除空白）
    recipient_emails = [email.strip() for email in recipient_email_str.replace(';', ',').split(',') if email.strip()]
    
    if not recipient_emails:
        logging.error("收件人郵箱列表為空，請檢查 RECIPIENT_EMAIL 設定。")
        return False
    
    logging.info(f"收件人列表：{', '.join(recipient_emails)}")

    # --- 建立郵件內容（HTML，無 emoji） ---
    subject = "自動填寫工作劃假表單總結報告"
    body = render_email_html(summary_data)

    # --- 判斷發送方式 ---
    if len(recipient_emails) == 1:
        # 只有 1 個收件人：同步發送
        logging.info(f"單一收件人，使用同步發送...")
        try:
            msg = MIMEText(body, 'html', 'utf-8')
            msg['From'] = f'自動填寫劃假表單 <{sender_email}>'
            msg['To'] = recipient_emails[0]
            msg['Subject'] = Header(subject, 'utf-8')
            
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                server.login(sender_email, app_password)
                server.sendmail(sender_email, recipient_emails, msg.as_string())
            
            logging.info(f"郵件發送成功，已發送至：{recipient_emails[0]}")
            return True
        except smtplib.SMTPAuthenticationError:
            logging.error("郵件發送失敗：SMTP 驗證錯誤。請檢查 SENDER_EMAIL 與 KEY 是否正確。")
            return False
        except Exception as e:
            logging.error(f"郵件發送時發生未預期的錯誤：{e}")
            return False
    else:
        # 多個收件人：使用 Python 3.14 的異步並發發送
        logging.info(f"多個收件人 ({len(recipient_emails)} 位)，使用異步並發送（Python 3.14）...")
        
        async def send_all_emails():
            """使用 asyncio.TaskGroup（Python 3.14 特性）並發發送所有郵件"""
            results = []
            
            # Python 3.14 的 TaskGroup 提供更好的錯誤處理和資源管理
            async with asyncio.TaskGroup() as tg:
                tasks = [
                    tg.create_task(
                        send_email_to_single_recipient_async(
                            email, sender_email, app_password, subject, body
                        )
                    )
                    for email in recipient_emails
                ]
            
            # 收集所有結果
            for task in tasks:
                results.append(await task)
            
            return results
        
        # 執行異步發送
        try:
            results = asyncio.run(send_all_emails())
        except* Exception as eg:  # Python 3.11+ 的 ExceptionGroup 語法
            logging.error(f"發送過程中發生多個錯誤：{eg}")
            results = []
        
        # 統計結果
        success_count = sum(1 for _, success in results if success)
        failed_count = len(results) - success_count
        failed_emails = [email for email, success in results if not success]
        
        # 顯示結果
        logging.info("=" * 60)
        logging.info("郵件發送總結")
        logging.info(f"成功：{success_count} 封")
        logging.info(f"失敗：{failed_count} 封")
        if failed_emails:
            logging.error(f"失敗的郵箱：{', '.join(failed_emails)}")
        logging.info("=" * 60)
        
        return failed_count == 0
