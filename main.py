# -*- coding: utf8 -*-
import telebot
import time
from telebot import types
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials  # ипортируем ServiceAccountCredentials
id_sms=None
number=None

def create_keyfile_dict():
    variables_keys = {
    "type": "<json>",
    "project_id": "<json>",
    "private_key_id": "<json>",
    "private_key": "<json>",
    "client_email": "<json>",
    "client_id": "<json>",
    "auth_uri": "<json>",
    "token_uri": "<json>",
    "auth_provider_x509_cert_url": "<json>",
    "client_x509_cert_url": "<json>"
    }
    return variables_keys
scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(create_keyfile_dict(), scope)
gs = gspread.authorize(creds)
bot = telebot.TeleBot('<YOUT TOKEN>')

def RASPISANIE(number, day, Name_day):
    sheet = gs.open("timesheet").get_worksheet(day)
    search_text = number.upper()
    find_number = sheet.find(search_text)

    column = find_number.col
    row = find_number.row + 1


    para = ["1 Пара",
            "2 Пара",
            "3 Пара",
            "4 Пара",
            "5 Пара",
            "6 Пара",
            "7 Пара"]
    calls = ["8.00-9.20",
             "9.30-10.50",
             "11.10-12.30",
             "12.50-14.10",
             "14.30-15.50",
             "16.10-17.30",
             "17.40-19.00"]

    raspisanie = []
    raspisanie.append(f'Группа - {search_text}\n')
    raspisanie.append(f'✅{Name_day}✅\n\n')
    for i in range(0, 7):
        if sheet.cell(row + i, column).value == '':
            raspisanie.append('')
        else:
            raspisanie.append(f'{para[i]} - ⏰{calls[i]}\n📕{sheet.cell(row + i, column).value}\n\n')
    return raspisanie

def WEEKEND_today(number, Name_day):
    search_text = number.upper()
    raspisanie = []
    raspisanie.append(f'Группа - {search_text}')
    raspisanie.append(f'✅{Name_day}✅')
    raspisanie.append(f'Сегодня можно отдохнуть 😴 🍻')
    return raspisanie
def WEEKEND_tommorow(number, Name_day):
    search_text = number.upper()
    raspisanie = []
    raspisanie.append(f'Группа - {search_text}')
    raspisanie.append(f'✅{Name_day}✅')
    raspisanie.append(f'Завтра можно отдохнуть 😴 🍻')
    return raspisanie
# -----------------------------------------------------------------------------------------------------------------------
number = ''
day = ''
group_list = ['11ГД', '11УМД', '11СХ', '21ГД', '21ГСк', '22ГД', '31ГД', '31ГСк',
              '32ГД', '41ГД', '42ГД', '21М', '31М', '32М', '42М', '21СХ', '31СХ', '21КЛ',
              '21П', '11КЛ', '11П', '11ИС', '11ММР', '11РМ', '12РМ', '22 ИС', '21ИС', '21РМ',
              '22РМ', '21КСК', '31ПИ', '32ПИ', '31РМ', '32РМ', '31КСК', '41КСК', '41ПИ', '42ПИ'
              ]

def keys(message):
    keyboard = telebot.types.ReplyKeyboardMarkup(True)
    keyboard.row('Расписание', 'Выбор группы')
    keyboard.row('Помощь')
    keyboard.row('/start')
    bot.send_message(message.chat.id, 'Привет  {}! \n'
                                      'Если ты учишься в КИГМ №23, то этот бот сможет показывать тебе твое расписание.\n'
                                      'Выбрать группу - нажмите /group.\n'
                                      'Чтобы получить расписание /timetable.\n'
                                      'Если возникили проблемы - нажмите /help'.format(message.chat.first_name),
                                      reply_markup=keyboard
                     )



@bot.message_handler(commands=['start'])
def start_command(message):
    keys(message)

@bot.message_handler(commands=['help'])
def help_command(message):
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.add(
        telebot.types.InlineKeyboardButton(
            'Написать разработчику', url='telegram.me/<YOUR TG>'
  )
    )
    bot.send_message(
        message.chat.id,
        '1) Напишите свою группу (не важно большими или маленькими буквами).\n' +
        '2) Выберете нужный вам день недели.\n' +
        '3) Вы получите расписание для выбранной вами группы и выбранного дня недели.\n' +
        '4) Чтобы получить расписание на другой день, просто нажмите на нужный вам день выше.',
        reply_markup=keyboard
    )

def error(message):
    bot.send_message(message.chat.id, 'Вы не выбрали группу /group\n'
                                      '\n'
                                      '\n'
                                      '/help\n')


@bot.message_handler(commands=['group'])
def group_command(message):
    group_com = bot.send_message(message.from_user.id, "{}, из какой вы группы?".format(message.chat.first_name))
    bot.register_next_step_handler(group_com, number_group)
def number_group(message):
    global number
    number = message.text
    if message.text == '/start':
        start_command(message)
    elif message.text == 'Помощь':
        help_command(message)
    elif message.text == 'Расписание':
        bot.send_message(message.from_user.id, 'Вы еще не выбрали группу')
        group_command(message)
    elif message.text.upper() in group_list:
        msg = (message.text).upper()
        group = (f'Выбрана группа: {msg}')
        day_number(message)
    else:
        bot.send_message(message.from_user.id, 'Увы, но такой группы нет, введите группу правильно')
        group_command(message)

@bot.message_handler(commands=['timetable'])
def day_number(message):
    if number == '':
        error(message)
    else:
        keyboard = types.InlineKeyboardMarkup()
        key_monday = types.InlineKeyboardButton(text='Понедельник', callback_data='monday')
        keyboard.add(key_monday)

        key_tuesday = types.InlineKeyboardButton(text='Вторник', callback_data='tuesday')
        keyboard.add(key_tuesday)

        key_wednesday = types.InlineKeyboardButton(text='Среда', callback_data='wednesday')
        keyboard.add(key_wednesday)

        key_thursday = types.InlineKeyboardButton(text='Четверг', callback_data='thursday')
        keyboard.add(key_thursday)

        key_friday = types.InlineKeyboardButton(text='Пятница', callback_data='friday')
        keyboard.add(key_friday)

        key_today = types.InlineKeyboardButton(text='Сегодня', callback_data='today')
        key_tomorrow = types.InlineKeyboardButton(text='Завтра', callback_data='tomorrow')
        keyboard.add(key_today, key_tomorrow)

        bot.send_message(message.chat.id, 'Выбери день недели', reply_markup=keyboard)

        rasp_sms = bot.send_message(message.chat.id, 'Выберете день и расписнаие поятвится тут')
        global id_sms
        id_sms = rasp_sms.id

global id_error_quota_msg
id_error_quota_msg = []
def error_quota(message):
    error_quota_msg = bot.send_message(message.chat.id, 'Вы так много раз запросили расписание, что сервера Google просят вас подождать(\n'
                                                        'Обычно это около 3-10 мин\n'
                                                        'После чего попробуйте просто выбрать нужный вам день снова, а эти сообщения удалятся автоматически.\n'
                                                        '/start\n'
                                                        '/group\n'
                                                        '/timetable\n'
                                                        '/help\n')

    id_error_quota_msg.append(error_quota_msg.id)




def msg_out(message,mes):

    try:
        if id_error_quota_msg != []:
            for i in range(len(id_error_quota_msg)):
                bot.delete_message(message.chat.id, id_error_quota_msg[i])
                id_error_quota_msg.remove(id_error_quota_msg[i])
            message_output = bot.edit_message_text(chat_id=message.chat.id, message_id=id_sms, text=''.join(mes))
            time.sleep(1)
        else:
            message_output = bot.edit_message_text(chat_id=message.chat.id, message_id=id_sms, text=''.join(mes))
            time.sleep(1)
    except IndexError:
        time.sleep(15)
        if id_error_quota_msg != []:
            for i in range(len(id_error_quota_msg)):
                bot.delete_message(message.chat.id, id_error_quota_msg[i])
                id_error_quota_msg.remove(id_error_quota_msg[i])
        return

@bot.callback_query_handler(func=lambda call: True)
def callback_worker(call):
    try:
        weekDays = ("Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскреенье")
        now = datetime.date.today()
        tomorrow = datetime.date.today() + datetime.timedelta(days=1)
        if call.data == "today":
            day = now.weekday()
            Name_day = weekDays[day]
            if day == 5:
                mes = WEEKEND_today(number, Name_day)
                msg_out(call.message, mes)
            elif day == 6:
                mes = WEEKEND_today(number, Name_day)
                msg_out(call.message, mes)
            else:
                mes = RASPISANIE(number, day, Name_day)
                msg_out(call.message, mes)
        elif call.data == "tomorrow":
            day = int(tomorrow.weekday())
            Name_day = weekDays[day]
            if day == 5:
                mes = WEEKEND_tommorow(number, Name_day)
                msg_out(call.message, mes)
            elif day == 6:
                mes = WEEKEND_tommorow(number, Name_day)
                msg_out(call.message, mes)
            else:
                mes = RASPISANIE(number, day, Name_day)
                msg_out(call.message, mes)
        elif call.data == "monday":
            day = 0
            Name_day = 'Понедельник'
            mes = RASPISANIE(number, day, Name_day)
            msg_out(call.message, mes)
        elif call.data == "tuesday":
            day = 1
            Name_day = 'Вторник'
            mes = RASPISANIE(number, day, Name_day)
            msg_out(call.message, mes)
        elif call.data == "wednesday":
            day = 2
            Name_day = 'Среда'
            mes = RASPISANIE(number, day, Name_day)
            msg_out(call.message, mes)
        elif call.data == "thursday":
            day = 3
            Name_day = 'Четверг'
            mes = RASPISANIE(number, day, Name_day)
            msg_out(call.message, mes)
        elif call.data == "friday":
            day = 4
            Name_day = 'Пятница'
            mes = RASPISANIE(number, day, Name_day)
            msg_out(call.message, mes)
    except gspread.exceptions.APIError:
        error_quota(call.message)
@bot.message_handler(content_types=['text'])
def callback_worker(message):
    if message.text == "Расписание":
        if number.upper() in group_list:
            day_number(message)
        else:
            bot.send_message(message.from_user.id, 'Вы еще не выбрали группу')
            group_command(message)
    elif message.text == "Выбор группы":
        group_command(message)
    elif message.text == "Помощь":
        help_command(message)


bot.polling(none_stop=True)

if __name__ == '__main__':
    bot.polling(none_stop=True)