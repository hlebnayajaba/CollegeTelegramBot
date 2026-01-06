# -*- coding: utf-8 -*-
import telebot

bot = telebot.TeleBot('8497231406:AAE7bNFUgkzgTS4t5tzrEBB26CrKZ8dG96o')

@bot.message_handler(content_types=['text'])
def get_text_messages(message):
    if message.text == "/info":
        bot.send_message(message.from_user.id, "Это бот для учебной части колледжа IT Top. На данный момент функционал еще не разработан, но скоро все будет готово.")
    elif message.text == "/help":
        bot.send_message(message.from_user.id, "Список доступных команд: /info - информация о боте. /help - список доступных команд.")
    else:
        bot.send_message(message.from_user.id, "Напишите /help, чтобы получить список команд.")

bot.polling(none_stop=True, interval=0)