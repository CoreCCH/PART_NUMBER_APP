from component import grid, clear_widgets, show_alert, create_label, create_lineedit, create_button
from component import login_page_widgets as widgets
from PyQt5.QtWidgets import QLineEdit
from PyQt5.QtCore import Qt
from aliy_mail import send_mail, send_mail_froget
import SQL_handler
import random
import re

verify_code = ''
user = 1

def send_forget_mail(mail_addr):
    employees = SQL_handler.fetch_data_from_Employees_table({'Email':mail_addr})
    if len(employees) == 0:
        show_alert("不存在的電子郵件")
        return
    else:
        status = send_mail_froget(mail_addr, employees[0][1], employees[0][2])
        if (status == True):
            show_alert("郵件發送成功","通知")
            log_in_page()
        else:
            show_alert(status)
            return

def sign_in(account, pw1, pw2, mail, certify_code):
    if account == "":
        show_alert("請輸入帳號")
        return
    if pw1 == "":
        show_alert("請輸入密碼")
        return
    if pw1 != pw2:
        show_alert("兩次輸入密碼不同")
        return
    if (verify_code != certify_code):
        show_alert("驗證碼有誤")
        return
    if certify_code == "":
        show_alert("請輸入驗證碼")
        return
    
    if len(SQL_handler.fetch_data_from_Employees_table({'Email':mail})) != 0:
        show_alert("已存在信箱")
        return
    
    status = SQL_handler.add_data_to_Employees_table(user_info=[account, pw1, mail])
    if status == True:
        show_alert("帳號創建成功!","通知")
        log_in_page()
    else:
        show_alert(status)

def send_certify_mail(mail_addr, label_msg):
    regex = r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'
    if re.match(regex, mail_addr) is not None:
        pass
    elif len(SQL_handler.fetch_data_from_Employees_table({'Email':mail_addr})) != 0:
        label_msg.setVisible(True)
        label_msg.setText("已存在信箱")
        return
    else:
        label_msg.setVisible(True)
        label_msg.setText("輸入信箱格式錯誤")
        return

    global verify_code
    verify_code = ''.join(map(str, random.sample(range(0, 9), 6)))
    status = send_mail(mail_addr, verify_code)
    if status == True:
        label_msg.setText("郵件已寄送")
        label_msg.setVisible(True)
    else:
        label_msg.setText(status)

def mail_repeat(mail_addr, label_msg):
    if len(SQL_handler.fetch_data_from_Employees_table({'Email':mail_addr})) != 0:
        label_msg.setVisible(True)
        label_msg.setText("已存在信箱")
    else:
        label_msg.setVisible(False)

def string_compare(pw1, pw2, label_msg2):
    if (pw1 == pw2):
        label_msg2.setVisible(False)
    else:
        label_msg2.setVisible(True)

def eye_click(line_password, Hide_Show):
    line_password.setEchoMode(Hide_Show)

def log_in(line_account, line_password):
    employees = SQL_handler.get_table_data("Employees")

    if isinstance(employees, str):
        show_alert((f"Error: {employees}"))
        log_in_page()
        return
    
    for employee in employees:
        employee_id, username, password, mail = employee
        if username == line_account.text() and password == line_password.text():
            global user
            user = employee_id
            from main_page import main_page
            main_page()
            return
        
    show_alert("帳號或密碼有誤")
    log_in_page()

def sign_in_page():
    clear_widgets(widgets)

    label_account = create_label('帳號:', 0 ,0 , 'center')
    widgets['label_account'].append(label_account)

    grid.addWidget(widgets['label_account'][-1],2,2,1,1)

    line_account = create_lineedit(0,0,width=230)
    widgets['line_account'].append(line_account)

    grid.addWidget(widgets['line_account'][-1],2,3,1,2)

    label_password = create_label('密碼:', 0 ,0 , 'center')
    widgets['label_password'].append(label_password)

    grid.addWidget(widgets['label_password'][-1],3,2,1,1)

    line_password = create_lineedit(0,0,width=230)
    widgets['line_password'].append(line_password)

    grid.addWidget(widgets['line_password'][-1],3,3,1,2)
    line_password.setEchoMode(QLineEdit.Password)

    label_password2 = create_label('確認密碼:', 0 ,0 , 'center')
    widgets['label_password2'].append(label_password2)

    grid.addWidget(widgets['label_password2'][-1],4,2,1,1)
    
    line_password2 = create_lineedit(0,0,width=230)
    widgets['line_password2'].append(line_password2)

    grid.addWidget(widgets['line_password2'][-1],4,3,1,2)
    line_password2.setEchoMode(QLineEdit.Password)

    label_message2 = create_label("兩次輸入密碼不同!",0,0,'left',font_size=10)
    widgets['label_message2'].append(label_message2)

    grid.addWidget(widgets['label_message2'][-1],4,3,2,1,alignment=Qt.AlignVCenter)
    label_message2.setVisible(False)

    label_mail = create_label('信箱:', 0 ,0 , 'center')
    widgets['label_mail'].append(label_mail)

    grid.addWidget(widgets['label_mail'][-1],5,2,1,1)

    line_mail = create_lineedit(0,0,width=230)
    widgets['line_mail'].append(line_mail)

    grid.addWidget(widgets['line_mail'][-1],5,3,1,2)

    button_send_mail = create_button("寄送認證信","#DDDDDD",0,0)
    widgets['button_send_mail'].append(button_send_mail)

    grid.addWidget(widgets['button_send_mail'][-1],5,4,1,1,alignment=Qt.AlignRight)

    label_message = create_label("郵件已寄出",0,0,'left',font_size=10)
    widgets['label_message'].append(label_message)

    grid.addWidget(widgets['label_message'][-1],5,3,2,1,alignment=Qt.AlignVCenter)
    label_message.setVisible(False)


    button_signin = create_button("註冊","#FF8888",0,0)
    widgets['button_signin'].append(button_signin)

    grid.addWidget(widgets['button_signin'][-1],6,4,1,1,alignment=Qt.AlignLeft)

    label_certify = create_label("驗證碼",0,0,'right')
    widgets['label_certify'].append(label_certify)

    grid.addWidget(widgets['label_certify'][-1],6,2,1,1)


    line_certify = create_lineedit(0,0,width=150)
    widgets['line_certify'].append(line_certify)

    grid.addWidget(widgets['line_certify'][-1],6,3,1,1)

    line_password2.editingFinished.connect(lambda: string_compare(line_password.text(), line_password2.text(), label_message2))
    line_mail.editingFinished.connect(lambda: mail_repeat(line_mail.text(), label_message))
    button_send_mail.clicked.connect(lambda: send_certify_mail(line_mail.text(), label_message))
    button_signin.clicked.connect(lambda: sign_in(line_account.text(), line_password.text(), line_password2.text(), line_mail.text(), line_certify.text()))

def log_in_page():
    clear_widgets(widgets)

    label_account = create_label('帳號:', 0 ,0 , 'center')
    widgets['label_account'].append(label_account)

    grid.addWidget(widgets['label_account'][-1],2,2,1,1)

    line_account = create_lineedit(0,0,width=230)
    widgets['line_account'].append(line_account)

    grid.addWidget(widgets['line_account'][-1],2,3,1,2)

    label_password = create_label('密碼:', 0 ,0 , 'center')
    widgets['label_password'].append(label_password)

    grid.addWidget(widgets['label_password'][-1],3,2,1,1)

    line_password = create_lineedit(0,0,width=230)
    widgets['line_password'].append(line_password)

    line_password.setEchoMode(QLineEdit.Password)

    grid.addWidget(widgets['line_password'][-1],3,3,1,2)

    button_forget_password = create_button("忘記密碼","#faffff",0,0, qcolor=[250, 255, 255, 255], font=15)
    widgets['button_forget_password'].append(button_forget_password)

    grid.addWidget(widgets['button_forget_password'][-1],3, 4, 1, 1,alignment=Qt.AlignRight)
    
    button_eye = create_button("","#6c6c6c",0,0,width=40,qcolor=[130,130,130,255])
    widgets['button_eye'].append(button_eye)

    grid.addWidget(widgets['button_eye'][-1],3,4,1,2)

    button_login = create_button("登入","#008E8E",0,0)
    widgets['button_login'].append(button_login)

    grid.addWidget(widgets['button_login'][-1],4,3,1,2,alignment=Qt.AlignCenter|Qt.AlignTop)

    button_signin = create_button("註冊","#FF8888",0,0)
    widgets['button_signin'].append(button_signin)

    grid.addWidget(widgets['button_signin'][-1],4,3,1,1,alignment=Qt.AlignLeft|Qt.AlignTop)

    button_eye.pressed.connect(lambda: eye_click(line_password, QLineEdit.Normal))
    button_eye.released.connect(lambda: eye_click(line_password, QLineEdit.Password))
    button_forget_password.clicked.connect(forget_page)
    button_login.clicked.connect(lambda:log_in(line_account, line_password))
    button_signin.clicked.connect(sign_in_page)

def forget_page():
    clear_widgets(widgets)

    label_mail = create_label('請輸入信箱:', 0 ,0 , 'center')
    widgets['label_mail'].append(label_mail)

    grid.addWidget(widgets['label_mail'][-1],2,2,1,1)

    line_mail = create_lineedit(0,0,width=230)
    widgets['line_mail'].append(line_mail)

    grid.addWidget(widgets['line_mail'][-1],2,3,1,2)

    button_send_mail = create_button("寄送","#DDDDDD",0,0)
    widgets['button_send_mail'].append(button_send_mail)

    grid.addWidget(widgets['button_send_mail'][-1],3,3,1,2,alignment=Qt.AlignHCenter)
    button_send_mail.clicked.connect(lambda: send_forget_mail(line_mail.text()))
