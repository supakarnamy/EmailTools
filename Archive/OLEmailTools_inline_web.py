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
import mimetypes
from urllib.request import urlopen, Request

# ================== CONFIG (hardcoded as requested) ==================
EMAIL_ADDRESS = "supakarn.than@gmail.com"
EMAIL_PASSWORD = "mujczfdscvwwhvvd"  # Gmail App Password (16 chars)

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587          # STARTTLS
SMTP_SSL_PORT = 465      # SMTPS (SSL on connect)
SMTP_TIMEOUT = 60
PER_EMAIL_DELAY = 1.0
MAX_SEND_RETRIES = 3     # retry per recipient with fresh reconnects
REQUIRED_COLUMNS = {"email", "business_name"}

# === Use RAW GitHub URLs instead of local files ===
RAW_PDF_URL  = "https://raw.githubusercontent.com/supakarnamy/EmailTools/images/gtm_playbook.pdf"
RAW_IMG1_URL = "https://raw.githubusercontent.com/supakarnamy/EmailTools/images/oplmail1.png"
RAW_IMG2_URL = "https://raw.githubusercontent.com/supakarnamy/EmailTools/images/oplmail2.png"

LOG_BASENAME = "email_sender.log"

# ================== Helper functions ==================
def sanitize_business_name(name: str) -> str:
    return re.sub(r"\W+", "", str(name)).lower()

def create_email_body(business_name: str, include_second_image: bool) -> str:
    username = sanitize_business_name(business_name)
    signup_link = f"https://www.openlink.co/auth/signup?username={username}"
    second_img_block = (
        """
        <div style=\"text-align:center;margin:16px 0;\">
          <img src=\"cid:embedded_image_2\" width=\"400\" />
        </div>
        """ if include_second_image else ""
    )
    return f"""
    <html><body style=\"font-family:Arial, sans-serif; line-height:1.6;\">
      <p>สวัสดีค่ะ เอมี่จากทีม openlink นะคะ<br>
      ทางทีมเห็นว่าร้าน <strong>{business_name}</strong> มีเมนูและสไตล์ที่น่าสนใจ จึงอยากแนะนำ <strong>openlink</strong> แพลตฟอร์มที่ช่วยรวมข้อมูลสำคัญไว้ในลิงก์เดียว แชร์ได้ทุกโซเชียล ให้ลูกค้าหาเจอง่ายขึ้น สั่งซื้อง่ายขึ้น หลายร้านใช้แล้วมียอดขายเพิ่มขึ้น ช่วยสร้างแบรนด์ให้เป็นที่จดจำ</p>

      <p>ตอนนี้เรามีแคมเปญพิเศษสำหรับร้านค้าชั้นนำ<br>
      <strong>Invite Package: ใช้งาน openlink Premium ฟรี 3 เดือน</strong><br>
      พร้อมบริการ Onboarding มูลค่า 3,990 บาท สำหรับร้านค้าที่เริ่มใช้งานภายในเดือนนี้ค่ะ</p>

      <p><strong>แนะนำ openlink และวิธีใช้งาน</strong><br>
      ดูรายละเอียดฟีเจอร์ ตัวอย่างเว็บไซต์ และขั้นตอนการใช้งานได้ในเอกสารที่แนบไปได้เลยค่ะ</p>

      <p><strong>เริ่มต้นใช้งานทันที</strong><br>
      หลังจากกดยืนยันอีเมลแล้วสามารถเข้าใช้งานในชื่อร้านของคุณได้เลยค่ะ:<br>
      <a href=\"{signup_link}\">{signup_link}</a></p>

      <div style=\"text-align:center;margin:16px 0;\">
        <img src=\"cid:embedded_image\" width=\"400\" />
      </div>

      <p><strong>ตัวอย่างร้านค้าที่ใช้งานจริง:</strong><br>
      <a href=\"https://www.openlink.co/hintcoffee\">https://www.openlink.co/hintcoffee</a><br>
      <a href=\"https://www.openlink.co/maguro\">https://www.openlink.co/maguro</a><br>
      หรือสำรวจเพิ่มเติมที่ <a href=\"https://www.openlink.co/explore\">https://www.openlink.co/explore</a></p>

      <p>หากทางร้านสนใจ เอมี่สามารถช่วยสร้างหน้าเพจตัวอย่างให้ฟรี เพื่อทดลองใช้งานก่อนตัดสินใจได้นะคะ 😊</p>

      {second_img_block}

      <p>Supakarn Thanyakriengkrai<br>
      Business Development Team, openlink<br>
      Bangkok, Thailand<br>
      (+66)88-501-5035</p>
    </body></html>
    """

def _http_get_bytes(url: str, timeout: int = 45) -> bytes:
    req = Request(url, headers={"User-Agent": "EmailTools/1.0"})
    with urlopen(req, timeout=timeout) as r:
        return r.read()

def _image_subtype_from_url(url: str) -> str:
    # Guess MIME subtype from extension; fallback to 'png'
    ctype, _ = mimetypes.guess_type(url)
    if ctype and ctype.startswith("image/"):
        return ctype.split("/")[1]
    return "png"

# ---- message builder ----
def build_message(recipient_email: str, business_name: str) -> MIMEMultipart:
    # Decide whether to include second image (if fetchable)
    include_img2 = True  # we'll try; if fails, we just skip attaching

    html_body = create_email_body(business_name, include_img2)

    msg = MIMEMultipart("related")
    msg['From'] = formataddr(("openlink BD (Amy)", EMAIL_ADDRESS))
    msg['Subject'] = str(Header("เริ่มใช้งาน openlink Premium ฟรี 3 เดือน", "utf-8"))
    msg['To'] = recipient_email

    alt = MIMEMultipart("alternative")
    msg.attach(alt)
    alt.attach(MIMEText(html_body, "html", _charset="utf-8"))

    # Inline image #1 (from RAW URL)
    try:
        img1_bytes = _http_get_bytes(RAW_IMG1_URL)
        subtype1 = _image_subtype_from_url(RAW_IMG1_URL)
        mime_img = MIMEImage(img1_bytes, _subtype=subtype1)
        mime_img.add_header("Content-ID", "<embedded_image>")
        mime_img.add_header("Content-Disposition", "inline", filename=os.path.basename(RAW_IMG1_URL))
        msg.attach(mime_img)
    except Exception:
        pass  # skip silently if fetch fails

    # Inline image #2 (from RAW URL)
    try:
        img2_bytes = _http_get_bytes(RAW_IMG2_URL)
        subtype2 = _image_subtype_from_url(RAW_IMG2_URL)
        mime_img2 = MIMEImage(img2_bytes, _subtype=subtype2)
        mime_img2.add_header("Content-ID", "<embedded_image_2>")
        mime_img2.add_header("Content-Disposition", "inline", filename=os.path.basename(RAW_IMG2_URL))
        msg.attach(mime_img2)
    except Exception:
        # If we can't fetch image 2, that's fine — the email still renders
        pass

    # Attach PDF (from RAW URL)
    try:
        pdf_bytes = _http_get_bytes(RAW_PDF_URL)
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(RAW_PDF_URL)}"')
        msg.attach(part)
    except Exception:
        pass

    return msg

# ---- TLS / SMTP connection helpers ----
def get_tls_context():
    import ssl
    ctx = ssl.create_default_context()
    try:
        import certifi
        ca_file = certifi.where()
        os.environ.setdefault("SSL_CERT_FILE", ca_file)
        try:
            ctx.load_verify_locations(cafile=ca_file)
        except Exception:
            pass
    except Exception:
        pass
    return ctx

def smtp_connect():
    """Always establish a NEW connection and login. Prefer SMTPS:465, fallback to 587."""
    ctx = get_tls_context()
    last_exc = None
    try:
        srv = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_SSL_PORT, context=ctx, timeout=SMTP_TIMEOUT)
        srv.ehlo()
        srv.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        return srv
    except Exception as e:
        last_exc = e
    try:
        srv = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=SMTP_TIMEOUT)
        srv.ehlo()
        srv.starttls(context=ctx)
        srv.ehlo()
        srv.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        return srv
    except Exception as e:
        last_exc = e
        msg = str(last_exc)
        if "CERTIFICATE_VERIFY_FAILED" in msg:
            msg += ("\n\nFix suggestions (macOS):\n"
                    "• Python.org install: run 'Install Certificates.command' in Applications/Python 3.x\n"
                    "• Conda: `conda install -y certifi` then restart\n"
                    "• Network: try off VPN/Proxy or allow smtp.gmail.com")
        raise RuntimeError(msg)

# ---- per-recipient sender (with retries and reconnect) ----
def send_one_with_retries(recipient: str, business_name: str, log_path: str):
    import socket
    attempts = 0
    last_err = None
    while attempts < MAX_SEND_RETRIES:
        attempts += 1
        try:
            server = smtp_connect()
            try:
                msg = build_message(recipient, business_name)
                server.send_message(msg)
                return None  # success -> no error message
            finally:
                try:
                    server.quit()
                except Exception:
                    pass
        except (smtplib.SMTPServerDisconnected, smtplib.SMTPException, socket.error) as e:
            last_err = f"{e} (attempt {attempts}/{MAX_SEND_RETRIES})"
            time.sleep(min(2 ** attempts, 8))
        except Exception as e:
            last_err = f"{e} (attempt {attempts}/{MAX_SEND_RETRIES})"
            time.sleep(min(2 ** attempts, 8))
    # return last error string to be logged
    try:
        with open(log_path, "a", encoding="utf-8") as lf:
            lf.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {recipient}: {last_err}\n")
    except Exception:
        pass
    return last_err or "unknown error"

# ================== GUI logic ==================
def ui_set_status(text: str):
    status_label.config(text=text)

def start_email_thread():
    start_btn.config(state='disabled')
    ui_set_status("📤 Preparing to send…")
    threading.Thread(target=send_batch, daemon=True).start()

def read_recipients(path: str):
    import pandas as pd
    _, ext = os.path.splitext(path.lower())
    if ext == ".csv":
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)

    # Normalize column names to lowercase for checking
    lower_cols = [c.lower() for c in df.columns]
    missing = REQUIRED_COLUMNS - set(lower_cols)
    if missing:
        raise RuntimeError(
            "Excel/CSV is missing required columns: " + ", ".join(sorted(missing)) +
            f"\nExpected columns: {sorted(REQUIRED_COLUMNS)}"
        )
    # Map back to original case
    cols = {c.lower(): c for c in df.columns}
    return df[cols['email']].tolist(), df[cols['business_name']].tolist()

def send_batch():
    try:
        if not app.file_path:
            raise RuntimeError("Please select a recipient file first.")

        emails, names = read_recipients(app.file_path)
        total = len(emails)
        successes, failures = 0, 0
        failed_rows = []

        log_path = os.path.join(os.path.dirname(app.file_path) or os.getcwd(), LOG_BASENAME)

        for i, (email, biz) in enumerate(zip(emails, names), start=1):
            email = str(email).strip()
            biz = str(biz).strip()
            app.after(0, ui_set_status, f"📨 Sending to {email} ({i}/{total})")
            try:
                if not re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
                    raise ValueError("Invalid email format")
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
            report_path = os.path.join(os.path.dirname(app.file_path) or os.getcwd(), "failed_sends.csv")
            pd.DataFrame(failed_rows).to_csv(report_path, index=False)
            app.after(0, messagebox.showwarning, "Partial Success",
                      f"Sent: {successes} | Failed: {failures}\nSaved details to {report_path}\nLog: {log_path}")
        else:
            app.after(0, messagebox.showinfo, "✅ Done",
                      f"All {successes} emails sent successfully!\nLog: {log_path}")
    except Exception as e:
        app.after(0, messagebox.showerror, "Error", str(e))
    finally:
        app.after(0, ui_set_status, "✅ Ready")
        app.after(0, start_btn.config, {"state": "normal"})

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV Files", "*.xlsx *.csv")])
    if file_path:
        app.file_path = file_path
        file_label.config(text=f"✔️ Loaded: {os.path.basename(file_path)}")

# ================== GUI ==================
app = tk.Tk()
app.title("Openlink Email Sender")
app.geometry("600x360")
app.file_path = None

tk.Label(app, text="📧 Openlink Email Outreach Tool", font=("Arial", 16)).pack(pady=10)

tk.Button(app, text="📁 Upload Recipients (.xlsx or .csv)", command=select_file).pack(pady=5)
file_label = tk.Label(app, text="No file selected", fg="gray")
file_label.pack()

start_btn = tk.Button(app, text="🚀 Start Sending Emails", command=start_email_thread,
                      bg="#4CAF50", fg="black", font=("Arial", 12))
start_btn.pack(pady=10)

status_label = tk.Label(app, text="✅ Ready", fg="blue")
status_label.pack(pady=10)

tk.Label(app, text="by Supakarn @ Openlink", fg="gray").pack(side="bottom", pady=5)

app.mainloop()