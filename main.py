import requests
import pandas as pd
import datetime
from io import StringIO
from openpyxl import Workbook
import holidays

# 设置日期
today = datetime.date.today()
cz_holidays = holidays.CZ()

# ✅ Step 1: 如果今天是节假日或周末，只打印一句话
if today.weekday() >= 5 or today in cz_holidays:
    reason = "weekend" if today.weekday() >= 5 else cz_holidays[today]
    print(f"Today ({today}) is a non-working day in Czech Republic due to {reason}. No FX rate retrieval.")
    exit()  # 直接退出程序

# 1. 设定计算时间为T-1
FX_date = datetime.date.today() - datetime.timedelta(days=1)

# 3. 如果不是工作日，就一直往前推
while FX_date.weekday() >= 5 or FX_date in cz_holidays:
    reason = "weekend" if FX_date.weekday() >= 5 else cz_holidays[FX_date]
    print(f"{FX_date} is a non-working day in Czech Republic due to {reason}. Pushing date back by 1 day.")
    FX_date -= datetime.timedelta(days=1)

print(f"Working day found: {FX_date}, start to load FX rate from CNB")

cnb_date = FX_date.strftime("%d.%m.%Y")  # CNB格式
filename_date = FX_date.strftime("%Y_%m_%d")
url = f"https://www.cnb.cz/cs/financni-trhy/devizovy-trh/kurzy-devizoveho-trhu/kurzy-devizoveho-trhu/denni_kurz.txt;jsessionid=ADF9BFAAE020EFB8483B43BB4DEAD0C8?date={cnb_date}"

response = requests.get(url)
lines = response.text.strip().split("\n")
info_line = lines[0].strip()
data_str = '\n'.join(lines[1:])

df = pd.read_csv(StringIO(data_str), sep='|')
df['množství'] = pd.to_numeric(df['množství'], errors='coerce').astype('Int64')
df['kurz'] = df['kurz'].str.replace(',', '.', regex=False).astype(float)

wb = Workbook()
ws = wb.active
ws.title = 'CNB FX Rates'

ws.append([info_line])
ws.append(df.columns.tolist())
for _, row in df.iterrows():
    ws.append(row.tolist())

output_file = "CNB_FX_rates.xls"
wb.save(output_file)
print(f"file saved as {output_file}")

# 4. 发送邮件

import smtplib
from email.message import EmailMessage
from pathlib import Path

EMAIL_ADDRESS = "binyi.zhangs@gmail.com"
EMAIL_PASSWORD = "bbpx vaoh ptbm gzjy"

msg = EmailMessage()
msg['Subject'] = f"CNB FX Rates: {filename_date}"
msg['From'] = EMAIL_ADDRESS
msg['To'] = "binyi.zhang@cz.icbc.com.cn, zhida.guo@cz.icbc.com.cn, " \
            "shan.he@cz.icbc.com.cn, tomas.houdek@cz.icbc.com.cn"
msg.set_content(f"""
Dear Colleagues,

Attached is the updated CNB exchange rates for {filename_date}.

Best regards,
Binyi Zhang.
""")

with open(output_file, 'rb') as f:
    file_data = f.read()
    file_name = Path(output_file).name
    msg.add_attachment(
        file_data,
        maintype='application',
        subtype='vnd.ms-excel',  # xls后缀就用这个
        filename=file_name
    )

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    smtp.send_message(msg)
    print("✅ Email sent successfully.")
