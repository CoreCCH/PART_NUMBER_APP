import smtplib
from email.mime.text import MIMEText
# import random

def send_mail(mail_addr, verify_code):
    # verify_code = ''.join(map(str, random.sample(range(0, 9), 6)))

    msg = MIMEText(f"驗證碼為:{verify_code}", 'plain', 'utf-8') # 郵件內文
    msg['Subject'] = '[認證信]Orient-SunTech ERP 輔助軟體'            # 郵件標題
    msg['From'] = 'RD-James'                  # 暱稱或是 email
    msg['To'] = mail_addr   # 收件人 email
    msg['Cc'] = ''   # 副本收件人 email ( 開頭的 C 大寫 )
    msg['Bcc'] = ''  # 密件副本收件人 email


    smtp = smtplib.SMTP('smtp.gmail.com', 587)
    smtp.ehlo()
    smtp.starttls()
    smtp.login('j1618james@gmail.com','hwog rigf iata hopj')
    status = smtp.send_message(msg)    # 改成 send_message

    smtp.quit()

    if status == {} or status[''][0] == 555:
        return True
    elif status[''][0] == 500:
        return "伺服器因語法錯誤而無法辨識指令"
    elif status[''][0] == 501:
        return "指令參數或引數中有語法錯誤"
    elif status[''][0] == 502:
        return "未執行指令"
    elif status[''][0] == 503:
        return "伺服器執行的指令順序有誤"
    elif status[''][0] == 541:
        return "收件者地址拒絕接收郵件"
    elif status[''][0] == 550:
        return "無法執行要求的指令，因為使用者的信箱無法使用，或收件伺服器將郵件視為垃圾郵件而拒絕接收"
    elif status[''][0] == 551:
        return "預定收件者的信箱無法在收件伺服器中使用"
    elif status[''][0] == 552:  
        return "收件者信箱的儲存空間不足，因此郵件無法傳送"
    elif status[''][0] == 553:
        return "信箱名稱不存在，因此指令已停止執行"
    elif status[''][0] == 554:  
        return "缺乏其他詳細資料，因此交易失敗"
    else:
        return "未知錯誤"

def send_mail_froget(mail_addr, account, password):
    msg = MIMEText(f"------------------------------------------\n|    您的帳號為:{account}    \n|   您的密碼為:{password}   \n-------------------------------------------", 'plain', 'utf-8') # 郵件內文
    msg['Subject'] = '[帳號密碼]Orient-SunTech ERP 輔助軟體'            # 郵件標題
    msg['From'] = 'RD-James'                  # 暱稱或是 email
    msg['To'] = mail_addr   # 收件人 email
    msg['Cc'] = ''   # 副本收件人 email ( 開頭的 C 大寫 )
    msg['Bcc'] = ''  # 密件副本收件人 email


    smtp = smtplib.SMTP('smtp.gmail.com', 587)
    smtp.ehlo()
    smtp.starttls()
    smtp.login('j1618james@gmail.com','hwog rigf iata hopj')
    status = smtp.send_message(msg)    # 改成 send_message

    smtp.quit()

    if status == {} or status[''][0] == 555:
        return True
    elif status[''][0] == 500:
        return "伺服器因語法錯誤而無法辨識指令"
    elif status[''][0] == 501:
        return "指令參數或引數中有語法錯誤"
    elif status[''][0] == 502:
        return "未執行指令"
    elif status[''][0] == 503:
        return "伺服器執行的指令順序有誤"
    elif status[''][0] == 541:
        return "收件者地址拒絕接收郵件"
    elif status[''][0] == 550:
        return "無法執行要求的指令，因為使用者的信箱無法使用，或收件伺服器將郵件視為垃圾郵件而拒絕接收"
    elif status[''][0] == 551:
        return "預定收件者的信箱無法在收件伺服器中使用"
    elif status[''][0] == 552:  
        return "收件者信箱的儲存空間不足，因此郵件無法傳送"
    elif status[''][0] == 553:
        return "信箱名稱不存在，因此指令已停止執行"
    elif status[''][0] == 554:  
        return "缺乏其他詳細資料，因此交易失敗"
    else:
        return "未知錯誤"

# send_mail('james.chiu@orient-suntech.com')