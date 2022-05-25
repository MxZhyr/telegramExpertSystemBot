import json

import telebot
# from telebot import types   ???
import openpyxl

notebooks = openpyxl.open("data.xlsx", read_only=True)
sheet = notebooks.active

bot = telebot.TeleBot('5159502195:AAGWyOky_QohEG9jOorz_r0zwZmMn7lFk8Q')  # +++


@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, '<b>Вас приветствует экспертная система по выбору одежды!</b>', parse_mode='html')
    bot.send_message(message.chat.id, '<b>Пожалуйста, введите тип занятия на этот день: </b>', parse_mode='html')


list = ['activity_type',
        'weather', 'size',
        'gender', 'style']
specIterator = iter(list)
selected = {}
current = next(specIterator)


def is_int(value):
    try:
        return int(value) == value
    except ValueError:
        return False


@bot.message_handler()
def get_text_messages(message):
    global current
    global selected
    global sheet
    global specIterator

    selected[current] = message.text
    current = next(specIterator, None)
    if current is None:
        bot.send_message(message.chat.id, json.dumps(selected), parse_mode='html')
        bot.send_message(message.chat.id, 'Вам подходит: ', parse_mode='html')

        for row in range(2, 9):
            ok = True
            for i, (k, v) in enumerate(selected.items()):
                if v == '-':
                    continue
                if k in ['size']:
                    b = is_int(v)
                    if b:
                        if int(v) != int(sheet[row][i].value):
                            ok = False
                        # print(sheet[row][1].value + ' price>' + k + '=' + v)
            if not ok:
                continue
            noteBookInfo = ''
            for i, it in enumerate(list):
                noteBookInfo += it + ': ' + str(sheet[row][i].value) + '; '
            bot.send_message(message.chat.id, noteBookInfo, parse_mode='html')
        selected = {}
        specIterator = iter(list)
        current = next(specIterator)
        return
    bot.send_message(message.chat.id, 'Выбирайте ' + current, parse_mode='html')


bot.polling(none_stop=True)
