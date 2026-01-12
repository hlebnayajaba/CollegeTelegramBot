import telebot
import os
import pandas as pd
import io
bot = telebot.TeleBot('8497231406:AAE7bNFUgkzgTS4t5tzrEBB26CrKZ8dG96o')


EXCEL_EXTENSIONS = ['.xls', '.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm', '.xlam', '.xla', '.xlw']
#тестовая функция для работы с таблицами, пока логика подробно не прописана т.к. мне нужны примеры всех 6 таблиц чтобы прописать взаимодействие с данными
def get_info(filename):
    try:
        file_path = os.path.join("Sheets", filename)
        df = pd.read_excel(file_path, engine='openpyxl')
        if df.dropna(how='all').empty:
            return "Файл открывается, но все ячейки пустые"
        else:
            info = {df.head().to_string()}
            return info
    except Exception as e:
        return None, f"Ошибка при чтении файла: {str(e)}"


@bot.message_handler(content_types=['text', 'document'])
def get_text_messages(message):
    if message.document:
        file_info = bot.get_file(message.document.file_id)
        file_name = message.document.file_name
        
        file_ext = os.path.splitext(file_name)[1].lower()
        
        if file_ext in EXCEL_EXTENSIONS:
            bot.send_message(message.from_user.id, "Файл принят")
        else:
            bot.send_message(message.from_user.id, "Формат файла не подходит. Отправьте файл Excel (.xls, .xlsx и т.д.)")
    elif message.text == "/info":
        bot.send_message(message.from_user.id, "Это бот для учебной части колледжа IT Top. На данный момент функционал еще не разработан, но скоро все будет готово.")
    elif message.text == "/help":
        bot.send_message(message.from_user.id, "Список доступных команд: /info - информация о боте. /help - список доступных команд.")
    else:
        bot.send_message(message.from_user.id, "Напишите /help, чтобы получить список команд.")



bot.polling(none_stop=True, interval=0)