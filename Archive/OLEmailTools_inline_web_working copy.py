# ==============================================
#      OLEmailTools_inline.py  (Gmail-compatible 2-image inline send)
# ==============================================

import tkinter as tk
from tkinter import filedialog, messagebox
import smtplib
import os
import re
import time
from email.header import Header
from email.utils import formataddr
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
import threading
import requests  # make sure this is installed

# ================== CONFIG (hardcoded as requested) ==================
EMAIL_ADDRESS = "supakarn.than@gmail.com"
EMAIL_PASSWORD = "meinbpegstmyzxoj"  # Gmail App Password recommended

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_SSL_PORT = 465
SMTP_TIMEOUT = 60
PER_EMAIL_DELAY = 1.0
MAX_SEND_RETRIES = 3
REQUIRED_COLUMNS = {"email", "business_name"}


HARDCODED_PDF_PATH  = "https://raw.githubusercontent.com/supakarnamy/EmailTools/images/Openlink%20Introduction%20V3_compressed.pdf"
HARDCODED_IMAGE_PATH = "https://raw.githubusercontent.com/supakarnamy/EmailTools/images/oplmail1.png"
HARDCODED_IMAGE_PATH_2 = "https://raw.githubusercontent.com/supakarnamy/EmailTools/images/oplmail2.png"



LOG_BASENAME = "email_sender.log"

def sanitize_business_name(name: str) -> str:
    return re.sub(r"\W+", "", str(name)).lower()

def create_html(business_name: str) -> str:
    username = sanitize_business_name(business_name)
    signup_link = f"https://www.openlink.co/auth/signup?username={username}"
    return f"""
    <html>
      <body style="font-family:Arial, sans-serif; line-height:1.6; color:#111;">
        <p>สวัสดีค่ะ เอมี่จากทีม <strong>openlink</strong> นะคะ</p>

        <p>
          ทางทีมเห็นว่าร้าน <strong>{business_name}</strong> <u>มีหลายช่องทาง</u>
          จึงอยากแนะนำ <strong>openlink</strong> แพลตฟอร์มที่ช่วยรวมทุกช่องทางและข้อมูลสำคัญไว้ในลิงก์เดียว
        </p>

        <p>
          แชร์ได้ทุกโซเชียล ทำให้ลูกค้าหาเจอง่าย สั่งซื้อได้ง่ายขึ้น
          หลายๆ ร้านใช้แล้วมียอดขายเพิ่มขึ้น และช่วยทำให้แบรนด์เป็นที่จดจำ
        </p>

        <p>
          <strong>บริการพิเศษเฉพาะเดือนนี้</strong><br>
          <strong>Invite Package:</strong> openlink ฟรีบริการสร้างและตกแต่งเพจ ใช้งานได้ทันที เพียงตอบกลับอีเมลนี้ภายใน 31 สิงหาคมนี้
        </p>

        <p>
          เริ่มใช้งานได้เองทันที สมัครผ่าน
          <a href="{signup_link}" target="_blank" rel="noopener noreferrer">{signup_link}</a>
        </p>

        <div style="text-align:center;margin:16px 0;">
          <img src="{HARDCODED_IMAGE_PATH}" alt="Openlink example 1" style="max-width:100%;height:auto;" />
        </div>

        <div style="text-align:center;margin:16px 0;">
          <img src="{HARDCODED_IMAGE_PATH_2}" alt="Openlink example 2" style="max-width:100%;height:auto;" />
        </div>

        <p>เอมี่ได้แนบเอกสารแนะนำและวิธีการใช้งานไว้ในอีเมลนี้ หากมีคำถามเพิ่มเติม สามารถตอบกลับอีเมลนี้ได้เลยค่ะ</p>

        <p>
          Supakarn Thanyakriengkrai<br>
          Business Development Team, openlink<br>
          Bangkok, Thailand<br>
          (+66)88-501-5035
        </p>
      </body>
    </html>
    """



def build_message(recipient_email: str, business_name: str) -> MIMEMultipart:
    html_body = create_html(business_name)

    msg = MIMEMultipart()
    msg['From'] = formataddr(("openlink BD (Amy)", EMAIL_ADDRESS))
    msg['Subject'] = str(Header("เริ่มใช้งาน openlink Premium ฟรี 3 เดือน", "utf-8"))
    msg['To'] = recipient_email
    msg.attach(MIMEText(html_body, "html", _charset="utf-8"))

    # Download and attach PDF from URL
    if HARDCODED_PDF_PATH.startswith("http"):
        pdf_data = requests.get(HARDCODED_PDF_PATH).content
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_data)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="Openlink_Introduction.pdf"')
        msg.attach(part)

    return msg


# (everything below is unchanged – GUI, threading, send logic, etc.)
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# [ EXACT copy from your script, included here so the file is self-contained: ]
import socket, ssl
def get_tls_context():
    ctx = ssl.create_default_context()
    return ctx

def smtp_connect():
    ctx = get_tls_context()
    last_exc = None
    try:
        srv = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_SSL_PORT, context=ctx, timeout=SMTP_TIMEOUT)
        srv.ehlo(); srv.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        return srv
    except Exception as e:
        last_exc = e
    try:
        srv = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=SMTP_TIMEOUT)
        srv.ehlo(); srv.starttls(context=ctx); srv.ehlo(); srv.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        return srv
    except Exception as e:
        raise RuntimeError(str(e))

def send_one_with_retries(recipient, business_name, log_path):
    attempts = 0
    last_err = None
    while attempts < MAX_SEND_RETRIES:
        attempts += 1
        try:
            server = smtp_connect()
            try:
                msg = build_message(recipient, business_name)
                server.send_message(msg)
                return None
            finally:
                try: server.quit()
                except Exception: pass
        except (smtplib.SMTPServerDisconnected, smtplib.SMTPException, socket.error) as e:
            last_err = f"{e}"
            time.sleep(min(2 ** attempts, 8))
    # log failure
    try:
        with open(log_path, "a", encoding="utf-8") as lf:
            lf.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {recipient}: {last_err}\n")
    except Exception:
        pass
    return last_err or "unknown error"

# ========== GUI code from your original file ==========
def ui_set_status(text: str):
    status_label.config(text=text)

def start_email_thread():
    start_btn.config(state='disabled')
    ui_set_status("📤 Preparing to send…")
    threading.Thread(target=send_batch, daemon=True).start()

def read_recipients(path: str):
    import pandas as pd
    _, ext = os.path.splitext(path.lower())
    df = pd.read_csv(path) if ext == ".csv" else pd.read_excel(path)
    missing = REQUIRED_COLUMNS - set(map(str.lower, df.columns))
    if missing:
        raise RuntimeError("Excel/CSV missing columns: " + ", ".join(sorted(missing)))
    cols = {c.lower(): c for c in df.columns}
    return df[cols['email']].tolist(), df[cols['business_name']].tolist()

def send_batch():
    if not app.file_path:
        messagebox.showerror("Error", "Please select a recipient file first.")
        start_btn.config(state='normal'); return
    emails, names = read_recipients(app.file_path)
    total = len(emails)
    successes, failures, failed_rows = 0, 0, []
    log_path = os.path.join(os.path.dirname(app.file_path) or os.getcwd(), LOG_BASENAME)
    for i, (email, biz) in enumerate(zip(emails, names), start=1):
        ui_set_status(f"📨 Sending to {email} ({i}/{total})")
        try:
            if not re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", email): raise ValueError("Invalid email format")
            err = send_one_with_retries(email, biz, log_path)
            if err is None:
                successes += 1
            else:
                failures += 1
                failed_rows.append({"row": i, "email": email, "business_name": biz, "error": err})
        except Exception as ex:
            failures += 1
            failed_rows.append({"row": i, "email": email, "business_name": biz, "error": str(ex)})
        time.sleep(PER_EMAIL_DELAY)
    if failed_rows:
        import pandas as pd
        report = os.path.join(os.path.dirname(app.file_path) or os.getcwd(), "failed_sends.csv")
        pd.DataFrame(failed_rows).to_csv(report, index=False)
        messagebox.showwarning("Partial", f"Sent {successes} | Failed {failures}\nSee: {report}")
    else:
        messagebox.showinfo("✅ Done", f"All {successes} emails sent!\nLog: {log_path}")
    ui_set_status("✅ Ready"); start_btn.config(state='normal')

def select_file():
    fp = filedialog.askopenfilename(filetypes=[("Excel/CSV Files","*.xlsx *.csv")])
    if fp: app.file_path = fp; file_label.config(text=f"✔️ Loaded: {os.path.basename(fp)}")

app = tk.Tk()
app.title("Openlink Email Sender")
app.geometry("600x360")
app.file_path = None

tk.Label(app,text="📧 Openlink Email Outreach Tool",font=("Arial",16)).pack(pady=10)
tk.Button(app,text="📁 Upload Recipients (.xlsx or .csv)",command=select_file).pack(pady=5)
file_label = tk.Label(app,text="No file selected",fg="gray"); file_label.pack()
start_btn = tk.Button(app,text="🚀 Start Sending Emails",command=start_email_thread,bg="#4CAF50",
                      fg="black",font=("Arial",12)); start_btn.pack(pady=10)
status_label = tk.Label(app,text="✅ Ready",fg="blue"); status_label.pack(pady=10)
tk.Label(app,text="by Supakarn @ Openlink",fg="gray").pack(side="bottom",pady=5)

app.mainloop()