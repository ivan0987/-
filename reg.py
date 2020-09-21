import PySimpleGUI as sg
import win32clipboard
import re
import random
from os import path


def gen_pass():  # генератор пароля
    my_pass = []
    my_parts = ['uppers', 'lowers', 'numbers', 'parts']
    parts = {
         'uppers': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'],
         'lowers': ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z'],
         'numbers': ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0'],
         'parts': ['!', '@', '#', '%', '&', '*', ')', '(']
      }
    my_pass += parts['uppers'][random.randint(0, len(parts['uppers']) - 1)]
    my_pass += parts['lowers'][random.randint(0, len(parts['lowers']) - 1)]
    my_pass += parts['numbers'][random.randint(0, len(parts['numbers']) - 1)]
    my_pass += parts['parts'][random.randint(0, len(parts['parts']) - 1)]
    for i in my_parts:
        rand = my_parts[random.randint(0, len(my_parts) - 1)]
        my_pass += parts[rand][random.randint(0, len(parts[rand]) - 1)]
    random.shuffle(my_pass)
    return ''.join(my_pass)


def writeexcel(last, first, email, squad, dom):  #запись в файл
    file_corp = './corp.csv'
    file_region = './region.csv'
    if dom == 'CORP':
        file = file_corp
    else:
        file = file_region
    pattern = '(\w*)(@)'
    find = re.findall(pattern, email)
    if find:  # проверяем почту
        match = re.search(pattern, email)
        user = match[1]
        user_pass = gen_pass()
        with open(file, 'a', encoding='utf8') as f:  # пишем
            append = user + ';' + user_pass + ';' + first + ';' + last.title() + ';' + email + ';' + squad + '\n'
            f.write(append)
        summary = [user, user_pass, first, last, email, squad]
        return summary
    else:
        sg.Popup('Почта некорректна!')
        return 'No'


if not path.exists('./corp.csv'):  # проверяем наличие файлов CSV, если нет, создаем и пишем шапку
    f = open('./corp.csv', 'w', encoding='utf8')
    f.write('username;password;firstname;lastname;email;cohort1\n')
    f.close()
    sg.Popup('Проверяем файлы...', auto_close=True, auto_close_duration=3, no_titlebar=True, modal=True)

if not path.exists('./region.csv'):
    f = open('./region.csv', 'w', encoding='utf8')
    f.write('username;password;firstname;lastname;email;cohort1\n')
    f.close()
    sg.Popup('Проверяем файлы...', auto_close=True, auto_close_duration=3, no_titlebar=True, modal=True)

# шаблон
layout = [
    [sg.Text('Фамилия', size=(7, 1)), sg.InputText(key='-Text_Last_Name-'), sg.Button(button_text='Из буфера', k='-Buff_Last_Name-', enable_events=True)],
    [sg.Text('Имя', size=(7, 1)), sg.InputText(key='-Text_First_Name-'), sg.Button(button_text='Из буфера', k='-Buff_First_Name-', enable_events=True)],
    [sg.Text('Почта', size=(7, 1)), sg.InputText(key='-Text_Email-'), sg.Button(button_text='Из буфера', k='-Buff_Email-', enable_events=True)],
    [sg.Combo(['Разработка', 'Управление качеством', 'Служба Эксплуатации'], default_value='Разработка', key='-cohort-')],
    [sg.Combo(['CORP', 'REGION'], default_value='CORP', key='-domain-')],
    [sg.Submit('Записать'), sg.Cancel('Отмена')],
    [sg.Output(size=(65, 7), key='-Output-')]
]

window = sg.Window('Регистрация для Лены;-)', layout)  # создаем окно
while True:                             # Ждем события
    event, values = window.read()
    if event in (None, 'Exit', 'Отмена'):
        break
    elif event == '-Buff_Last_Name-':  # тут вставки из буфера
        win32clipboard.OpenClipboard()
        buff = win32clipboard.GetClipboardData()
        values['-Text_Last_Name-'] = buff
        window['-Text_Last_Name-'].update(buff.title())
        win32clipboard.CloseClipboard()
    elif event == '-Buff_First_Name-':
        win32clipboard.OpenClipboard()
        buff = win32clipboard.GetClipboardData()
        values['-Text_First_Name-'] = buff
        window['-Text_First_Name-'].update(buff)
        win32clipboard.CloseClipboard()
    elif event == '-Buff_Email-':
        win32clipboard.OpenClipboard()
        buff = win32clipboard.GetClipboardData()
        values['-Text_Email-'] = buff
        window['-Text_Email-'].update(buff)
        win32clipboard.CloseClipboard()
    elif event == 'Записать':
        if values['-Text_Last_Name-'] == '' or values['-Text_First_Name-'] == '' or values['-Text_Email-'] == '':
            window['-Output-'].update('Чего-то не хватает')
        else:
            cohort = values['-cohort-']
            domain = values['-domain-']
            a = writeexcel(values['-Text_Last_Name-'], values['-Text_First_Name-'], values['-Text_Email-'], cohort, domain)
            if a != 'No':
                b = 'Записано:' + '\n' + 'Логин: ' + a[0] + '\n' + 'Пароль: ' + a[1] + '\n' + a[2] + '\n' + a[3] + '\n' + a[4] + '\n' + a[5] + '\n'
                log = b
                sg.Popup('ждем закрытия документа...', auto_close=True, auto_close_duration=3, no_titlebar=True, modal=True)
                window['-Output-'].update(log)
                window['-Text_Last_Name-'].update('')
                window['-Text_First_Name-'].update('')
                window['-Text_Email-'].update('')
            else:
                log = 'Попробуйте снова'
                window['-Output-'].update(log)
