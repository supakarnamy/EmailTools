import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openai
import threading

# --- CONFIG ---
EMAIL_ADDRESS = "supakarn.than@gmail.com"
EMAIL_PASSWORD = "saiodgthtdoepnvd"  # Gmail App Password
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
OPENAI_API_KEY = "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"  # Required for GPT
openai.api_key = OPENAI_API_KEY

# --- Functions ---
def create_email_body(business_name):
    return f"""สวัสดีค่า เอมี่นะคะ ติดต่อมาจากทีม openlink

ทางทีมเห็นว่าทางร้าน "{business_name}" มีหลายเมนูน่าสนใจ จึงอยากทำเว็บ openlink รวมเมนู และข้อมูลร้านให้เผื่อทางร้านลองนำไปใช้งาน 
ตอนนี้ openlink มี Invite Package Openlink ใช้ Premium ฟรี 3 เดือนพร้อม Onboarding มูลค่า 3,990 บาท สำหรับร้านค้าที่เข้าใช้งานภายในเดือนกรกฎาคม

ตัวอย่างร้านใน community ของเรา:
- https://www.openlink.co/hintcoffee 
- https://www.openlink.co/soyou_bake

openlink เป็น platform ที่ช่วยให้ร้านค้าสร้าง digital profile ของตัวเองได้ง่ายๆ พร้อมหน้า analytics ผู้ชมที่เข้าใช้งานที่ช่วยเพิ่มยอดขายทั้งทางออนไลน์ออฟไลน์ 
ด้วยการใส่ช่องทางจำหน่ายและตำแหน่งสาขา และช่วยโปรโมตร้านค้าบนช่องทาง Social กว่า 3,000+ ผู้ติดตาม

ดูตัวอย่างเพิ่มเติมได้ที่: https://www.openlink.co/explore

หากทางร้านสนใจ เอมี่จะลองทำหน้าเพจของทางร้านและส่งให้ลองใช้งานดูค่ะ 😊

Supakarn Thanyakriengkrai
Business Development Team, openlink
Bangkok, Thailand
"""

def generate_tailored_email(business_name):
    prompt = (
        f"เขียนอีเมลแนะนำแพลตฟอร์ม openlink ให้ร้าน {business_name} "
        "โดยใช้ภาษาไทย เป็นกันเอง มีการยกตัวอย่างร้านในชุมชน และแนะนำการใช้งาน Invite Package ฟรี 3 เดือน พร้อมรายละเอียดบริการ"
    )
    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content.strip()

def send_email(recipient_email, subject, body):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.send_message(msg)

# --- GUI Functions ---
def start_email_thread():
    threading.Thread(target=start_sending).start()

def start_sending():
    try:
        df = pd.read_excel(app.file_path)
        progress['maximum'] = len(df)
        for i, row in df.iterrows():
            email = row['email']
            business = row['business_name']

            if use_ai.get():
                body = generate_tailored_email(business)
            else:
                body = create_email_body(business)

            send_email(email, "openlink Free Premium 3 Months Subscription", body)
            progress['value'] = i + 1
            app.update_idletasks()

        messagebox.showinfo("Done", "Emails sent successfully! ✅")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        app.file_path = file_path
        file_label.config(text=f"✔️ Loaded: {file_path.split('/')[-1]}")

# --- GUI Layout ---
app = tk.Tk()
app.title("Openlink Email Sender")
app.geometry("500x300")
app.file_path = None

title = tk.Label(app, text="📧 Openlink Email Outreach Tool", font=("Arial", 16))
title.pack(pady=10)

file_btn = tk.Button(app, text="📁 Upload Excel File (.xlsx)", command=select_file)
file_btn.pack(pady=5)

file_label = tk.Label(app, text="No file selected", fg="gray")
file_label.pack()

use_ai = tk.BooleanVar()
ai_checkbox = tk.Checkbutton(app, text="✨ Enable AI Tailoring (OpenAI GPT-4)", variable=use_ai)
ai_checkbox.pack(pady=5)

start_btn = tk.Button(app, text="🚀 Start Sending Emails", command=start_email_thread, bg="#4CAF50", fg="white", font=("Arial", 12))
start_btn.pack(pady=10)

progress = ttk.Progressbar(app, length=400)
progress.pack(pady=10)

footer = tk.Label(app, text="by Supakarn @ Openlink", fg="gray")
footer.pack(side="bottom", pady=5)

app.mainloop()
