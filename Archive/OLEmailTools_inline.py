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

HARDCODED_PDF_PATH = "C:/Users/Taninlimpabandh/Downloads/openlink/gtm_playbook.pdf"

# Images (PNG, Gmail inline)
HARDCODED_IMAGE_PATH = "C:/Users/Taninlimpabandh/Downloads/openlink/oplmail1.png"
IMAGE_DIR = os.path.dirname(HARDCODED_IMAGE_PATH) if HARDCODED_IMAGE_PATH else ""
HARDCODED_IMAGE_PATH_2 = os.path.join(IMAGE_DIR, "oplmail2.png") if IMAGE_DIR else ""

LOG_BASENAME = "email_sender.log"

def sanitize_business_name(name: str) -> str:
    return re.sub(r"\W+", "", str(name)).lower()

def create_email_body(business_name: str, include_second_image: bool) -> str:
    username = sanitize_business_name(business_name)
    signup_link = f"https://www.openlink.co/auth/signup?username={username}"

    second_img_block = (
        """
        <div style="text-align:center;margin:16px 0;">
          <img src="cid:embedded_image_2" width="800" />
        </div>
        """ if include_second_image else ""
    )

    return f"""
    <html><body style="font-family:Arial, sans-serif; line-height:1.6;">
      <p>สวัสดีค่ะ เอมี่จากทีม openlink นะคะ<br>
      ทางทีมเห็นว่าร้าน <strong>{business_name}</strong> มีเมนูและสไตล์ที่น่าสนใจ จึงอยากแนะนำ <strong>openlink</strong> แพลตฟอร์มที่ช่วยรวมข้อมูลสำคัญไว้ในลิงก์เดียว แชร์ได้ทุกโซเชียล ให้ลูกค้าหาเจอง่ายขึ้น สั่งซื้อง่ายขึ้น หลายร้านใช้แล้วมียอดขายเพิ่มขึ้น ช่วยสร้างแบรนด์ให้เป็นที่จดจำ</p>

      <p>ตอนนี้เรามีแคมเปญพิเศษสำหรับร้านค้าชั้นนำ<br>
      <strong>Invite Package: ใช้งาน openlink Premium ฟรี 3 เดือน</strong><br>
      พร้อมบริการ Onboarding มูลค่า 3,990 บาท สำหรับร้านค้าที่เริ่มใช้งานภายในเดือนนี้ค่ะ</p>

      <p><strong>แนะนำ openlink และวิธีใช้งาน</strong><br>
      ดูรายละเอียดฟีเจอร์ ตัวอย่างเว็บไซต์ และขั้นตอนการใช้งานได้ในเอกสารที่แนบไปได้เลยค่ะ</p>

      <p><strong>เริ่มต้นใช้งานทันที</strong><br>
      หลังจากกดยืนยันอีเมลแล้วสามารถเข้าใช้งานในชื่อร้านของคุณได้เลยค่ะ:<br>
      <a href="{signup_link}">{signup_link}</a></p>

      <div style="text-align:center;margin:16px 0;">
        <img src="cid:embedded_image" width="800" />
      </div>

      <p><strong>ตัวอย่างร้านค้าที่ใช้งานจริง:</strong><br>
      <a href="https://www.openlink.co/burgerkingthailand">https://www.openlink.co/burgerkingthailand</a><br>
      <a href="https://www.openlink.co/maguro">https://www.openlink.co/maguro</a><br>
      หรือสำรวจเพิ่มเติมที่ <a href="https://www.openlink.co/explore">https://www.openlink.co/explore</a></p>

      <p>หากทางร้านสนใจ เอมี่สามารถช่วยสร้างหน้าเพจตัวอย่างให้ฟรี เพื่อทดลองใช้งานก่อนตัดสินใจได้นะคะ 😊</p>

      {second_img_block}

      <p>Supakarn Thanyakriengkrai<br>
      Business Development Team, openlink<br>
      Bangkok, Thailand<br>
      (+66)88-501-5035</p>
    </body></html>
    """

def build_message(recipient_email: str, business_name: str) -> MIMEMultipart:
    include_img2 = bool(HARDCODED_IMAGE_PATH_2 and os.path.isfile(HARDCODED_IMAGE_PATH_2))
    html_body = create_email_body(business_name, include_img2)

    msg = MIMEMultipart("related")
    msg['From'] = formataddr(("openlink BD (Amy)", EMAIL_ADDRESS))
    msg['Subject'] = str(Header("เริ่มใช้งาน openlink Premium ฟรี 3 เดือน", "utf-8"))
    msg['To'] = recipient_email

    alt = MIMEMultipart("alternative")
    msg.attach(alt)
    alt.attach(MIMEText(html_body, "html", _charset="utf-8"))

    # Inline PNG #1
    if HARDCODED_IMAGE_PATH and os.path.isfile(HARDCODED_IMAGE_PATH):
        with open(HARDCODED_IMAGE_PATH, "rb") as f:
            mime_img = MIMEImage(f.read(), _subtype="png")
        mime_img.add_header("Content-ID", "<embedded_image>")
        mime_img.add_header("Content-Disposition", "inline", filename=os.path.basename(HARDCODED_IMAGE_PATH))
        msg.attach(mime_img)

    # Inline PNG #2
    if include_img2:
        with open(HARDCODED_IMAGE_PATH_2, "rb") as f:
            mime_img2 = MIMEImage(f.read(), _subtype="png")
        mime_img2.add_header("Content-ID", "<embedded_image_2>")
        mime_img2.add_header("Content-Disposition", "inline", filename=os.path.basename(HARDCODED_IMAGE_PATH_2))
        msg.attach(mime_img2)

    # Attach PDF (optional)
    if HARDCODED_PDF_PATH and os.path.isfile(HARDCODED_PDF_PATH):
        with open(HARDCODED_PDF_PATH, "rb") as f:
            part = MIMEBase("application", "pdf")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(HARDCODED_PDF_PATH)}"')
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
