# =============================
import requests 
import pandas as pd
import datetime
from io import StringIO
from openpyxl import Workbook, load_workbook
import holidays


# Need to find out czech holiday
# set date function: t-1 
FX_date = datetime.date.today() - datetime.timedelta(days=1)


# Load Czech Holidays
cz_holidays = holidays.CZ()

# Check: is today a weekend (Saturday=5 or Sunday=6)?
is_weekend = FX_date.weekday() >= 5

if FX_date.weekday() >= 5 or FX_date in cz_holidays:
    reason = "weekend" if FX_date.weekday() >= 5  else cz_holidays[FX_date]
    print(f"{FX_date} is a non-working day in Czech Republic due to {reason} FX rating not available from CNB")

else:
    print('Working day in Czech Republic, start to load FX date from CNB')
    cnb_date = FX_date.strftime("%d.%m.%Y")  # 转成 CNB 格式：DD.MM.YYYY
    filename_date = FX_date.strftime("%Y_%m_%d")
    url = f"https://www.cnb.cz/cs/financni-trhy/devizovy-trh/kurzy-devizoveho-trhu/kurzy-devizoveho-trhu/denni_kurz.txt;jsessionid=ADF9BFAAE020EFB8483B43BB4DEAD0C8?date={cnb_date}"
    
    response = requests.get(url)
    lines = response.text.strip().split("\n")  # today's date
    info_line = lines[0].strip()     # 'dd mm yyyy #xxx'
    data_str = '\n'.join(lines[1:])  # 从第2行开始保留（含标题）

    df = pd.read_csv(StringIO(data_str), sep = '|')

    df['množství'] = pd.to_numeric(df['množství'], errors='coerce').astype('Int64')
    # Convert czech number to normal numer
    df['kurz'] = df['kurz'].str.replace(',', '.', regex=False).astype(float)

    wb = Workbook()
    ws = wb.active
    ws.title = 'CNB FX Rates'


    # load info line and header
    ws.append([info_line])
    ws.append(df.columns.tolist())

    # load fx rate
    for _, row in df.iterrows():
        ws.append(row.tolist())

    # Save as Excel
    output_file = f"CNB_FX_rates.xls"
    wb.save(output_file)
    print(f"file saved as {output_file}")


    import smtplib
    from email.message import EmailMessage
    from pathlib import Path
    
    # ========= Gmail 账户 =========
    EMAIL_ADDRESS = "binyi.zhangs@gmail.com"       
    EMAIL_PASSWORD = "bbpx vaoh ptbm gzjy"        

    msg = EmailMessage()
    msg['Subject'] = f"CNB FX Rates: {filename_date}"
    msg['From'] = EMAIL_ADDRESS
    # Send Email to ICBC RMD department
    msg['To'] = "binyi.zhang@cz.icbc.com.cn, zhijia.guo@cz.icbc.com.cn, " \
                "shan.he@cz.icbc.com.cn, tomas.houdek@cz.icbc.com.cn"
    msg.set_content(f"""
    Dear Colleagues,
    
    Attached is the updated CNB exchange rates for {filename_date}.
    
    Best regards,
    Binyi Zhang.
    
    Best regards,
    Binyi Zhang
    """)
    
    # ========= add Excel 附件 =========
    with open(output_file, 'rb') as f:
        file_data = f.read()
        file_name = Path(output_file).name
        msg.add_attachment(
            file_data,
            maintype='application',
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=file_name
        )
    
    # ========= 发送邮件 =========
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)
    
    print("✅ Email sent successfully.")