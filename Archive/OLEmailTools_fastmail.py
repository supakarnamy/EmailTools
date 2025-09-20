# ==============================================
#      OLEmailTools_fastmail_jmap.py
#      (Fastmail JMAP bulk sender + .txt log)
# ==============================================

import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import requests, pandas as pd, threading, time, re, json, os

# -------- FASTMAIL JMAP CONFIG -------
FASTMAIL_API_TOKEN = "fmu1-9349404e-d6411b980a322cdc52db92284d8897c7-0-a542ed5f1043d42fe93de5dd12eff1c4"
FASTMAIL_ACCOUNT_ID = "u9349404e"
FASTMAIL_FROM = "partnership@openlink.co"
FASTMAIL_SEND_URL = "https://api.fastmail.com/jmap/"

# -------- EMAIL CONTENT CONFIG -------
HARDCODED_PDF_URL  = "https://raw.githubusercontent.com/supakarnamy/EmailTools/images/gtm_playbook.pdf"
IMG1 = "https://raw.githubusercontent.com/supakarnamy/EmailTools/images/oplmail1.png"
IMG2 = "https://raw.githubusercontent.com/supakarnamy/EmailTools/images/oplmail2.png"
PER_EMAIL_DELAY = 1.0
REQUIRED_COLUMNS = {"email", "business_name"}
LOG_FILE = "fastmail_sends.txt"

# ----------- HELPER ------------------
def sanitize(b): return re.sub(r"\W+", "", str(b)).lower()

def create_html(business_name: str) -> str:
    username = sanitize(business_name)
    signup_link = f"https://www.openlink.co/auth/signup?username={username}"
    return f"""
    <html><body style="font-family:Arial, sans-serif; line-height:1.6;">
      <p>สวัสดีค่ะ เอมี่จากทีม openlink นะคะ<br>
      ทางทีมเห็นว่าร้าน <strong>{business_name}</strong> มีเมนูและสไตล์ที่น่าสนใจที่ขายอยู่หลายช่องทาง หลายสาขา จึงอยากแนะนำ <strong>openlink</strong> แพลตฟอร์มที่ช่วยรวมทุกหน้าร้านออนไลน์ ออฟไลน์ และข้อมูลสำคัญไว้ในลิงก์เดียว แชร์ได้ทุกโซเชียล ให้ลูกค้าหาเจอง่ายขึ้น สั่งซื้อง่ายขึ้น หลายร้านใช้แล้วมียอดขายเพิ่มขึ้น ช่วยสร้างแบรนด์ให้เป็นที่จดจำ</p>

      <p>บริการพิเศษเฉพาะเดือนนี้ สำหรับคุณ<br>
      <strong>Invite Package: openlink ฟรีบริการสร้างและตกแต่งเพจ ใช้งานได้ทันที</strong><br>
      เพียงตอบกลับอีเมลนี้ภายใน 31 สิงหาคมนี้</p>

      <p><strong>สร้างเพจเองง่าย ๆ</strong> เพียงสมัครผ่าน <a href="{signup_link}">{signup_link}</a></p>

      <div style="text-align:center;margin:16px 0;">
        <img src="{IMG1}" style="max-width:100%; height:auto;" />
      </div>

      <p>ตัวอย่างแบรนด์บน openlink เช่น <a href="https://www.openlink.co/burgerkingthailand">Burger King</a>, <a href="https://www.openlink.co/maguro">Maguro</a>, และอีกหลายพันหรือสำรวจเพิ่มเติมที่ <a href="https://www.openlink.co/explore">แบรนด์ชั้นนำ</a></p>

      <div style="text-align:center;margin:16px 0;">
        <img src="{IMG2}" style="max-width:100%; height:auto;" />
      </div>

      <p>เอมี่ได้แนบเอกสารแนะนำและวิธีการใช้งานมาในอีเมลนี้ หากมีคำถามเพิ่มเติม สามารถสอบถามผ่านอีเมลนี้ได้เลยค่ะ</p>

      <p>Supakarn Thanyakriengkrai<br>
      Business Development Team, openlink<br>
      Bangkok, Thailand<br>
      (+66)88-501-5035</p>
    </body></html>
    """

def jmap_send(to_addr, subject, html):
    r = requests.get(HARDCODED_PDF_URL)
    attach = [{
        "blobId": None,
        "name": "gtm_playbook.pdf",
        "type": "application/pdf",
        "data": r.content.decode("latin1")
    }] if r.ok else []

    payload = {
      "using": ["urn:ietf:params:jmap:core","urn:ietf:params:jmap:mail"],
      "methodCalls":[
        ["Email/set",{
            "accountId": FASTMAIL_ACCOUNT_ID,
            "create": {
               "msg": {
                 "from":[{"email": FASTMAIL_FROM,"name":"openlink BD (Amy)"}],
                 "to":[{"email": to_addr}],
                 "subject": subject,
                 "htmlBody": [{"partId":"1","type":"text/html","value": html}],
                 "attachments": attach
               }
            }
         },"c1"],
        ["EmailSubmission/set",{
            "accountId": FASTMAIL_ACCOUNT_ID,
            "create": {
              "es1": {
                "emailId": "#msg",
                "identityId": FASTMAIL_ACCOUNT_ID
              }
            }
         },"c2"]
      ]
    }
    headers={
      "Authorization": f"Bearer {FASTMAIL_API_TOKEN}",
      "Content-Type":"application/json",
      "Accept":"application/json"
    }
    res = requests.post(FASTMAIL_SEND_URL, headers=headers, data=json.dumps(payload))
    if res.status_code != 200: raise RuntimeError(res.text)

# ------------ GUI BATCH --------------
def start_thread(): threading.Thread(target=do_send, daemon=True).start()

def do_send():
    try:
        df = pd.read_excel(app.file_path)
        missing = REQUIRED_COLUMNS - set(map(str.lower, df.columns))
        if missing: raise RuntimeError(f"Missing columns: {missing}")

        progress['maximum'] = len(df)
        ok = 0; fail = 0

        with open(LOG_FILE, "a", encoding="utf-8") as lf:
            for i,(em,biz) in enumerate(zip(df['email'],df['business_name'])):
                progress['value']=i+1; app.update_idletasks()
                if not re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", str(em)):
                    lf.write(f"[SKIP] row {i+1}: invalid {em}\n")
                    continue
                html=create_html(str(biz))
                try:
                    jmap_send(str(em),"เริ่มใช้งาน openlink Premium ฟรี 3 เดือน", html)
                    lf.write(f"[OK] {em} ({biz})\n"); ok+=1
                except Exception as ex:
                    lf.write(f"[FAIL] {em} ({biz}) => {ex}\n"); fail+=1
                time.sleep(PER_EMAIL_DELAY)

        messagebox.showinfo("Result",f"✅ Sent {ok} | ❌ Failed {fail}  (see {LOG_FILE})")
    except Exception as e:
        messagebox.showerror("Error",str(e))
    finally:
        start_btn.config(state='normal')

def pick():
    f = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])
    if f:
        app.file_path = f
        file_label.config(text=f"✔ {os.path.basename(f)}")

# GUI -------------------------------
app = tk.Tk(); app.title("Openlink Fastmail Sender"); app.geometry("500x300"); app.file_path=None
tk.Label(app,text="📧 Openlink Email Outreach Tool",font=("Arial",16)).pack(pady=10)
tk.Button(app,text="📁 Upload .xlsx",command=pick).pack()
file_label = tk.Label(app,text="No file selected",fg="gray"); file_label.pack()
start_btn = tk.Button(app,text="🚀 Start Sending",command=start_thread,bg="#4CAF50",fg="white"); start_btn.pack(pady=10)
progress = ttk.Progressbar(app,length=400); progress.pack(pady=10)
tk.Label(app,text="by Supakarn @ Openlink",fg="gray").pack(side="bottom",pady=5)
app.mainloop()
