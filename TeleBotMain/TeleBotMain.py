import telebot
import os
import pandas as pd
import io
bot = telebot.TeleBot('8497231406:AAE7bNFUgkzgTS4t5tzrEBB26CrKZ8dG96o')


EXCEL_EXTENSIONS = ['.xls', '.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm', '.xlam', '.xla', '.xlw']
#функция для извлечения данных из таблиц
def get_info(filename):
    try:
        file_ext = os.path.splitext(filename)[1].lower()
        file_path = os.path.join("Sheets", filename)
        if file_ext == '.xls':
            df = pd.read_excel(file_path, engine='xlrd')
        else:
            df = pd.read_excel(file_path, engine='openpyxl')

        if df.dropna(how='all').empty:
            return "Файл открывается, но все ячейки пустые"
        else:
            return df
    except Exception as e:
        return None, f"Ошибка при чтении файла: {str(e)}"

#домашки, проверенные учителями
def checked_homework():
    df = get_info("Отчет по домашним заданиям.xlsx")
    if type(df) != str:
        result_lines = []        
        for index, row in df.iloc[2:].iterrows():
            if pd.notna(row.iloc[1]) and row.iloc[1].strip() != "":
                teacher_name = row.iloc[1]
                month_checked = row.iloc[5]  
                month_plan = row.iloc[6]    
                week_checked = row.iloc[10]  
                week_plan = row.iloc[11]     
                if pd.notna(month_checked) and pd.notna(month_plan) and month_plan != 0:
                    month_percent = (month_checked / month_plan) * 100
                else:
                    month_percent = 0
                
                if pd.notna(week_checked) and pd.notna(week_plan) and week_plan != 0:
                    week_percent = (week_checked / week_plan) * 100
                else:
                    week_percent = 0
                if month_percent < 70 or week_percent < 70:
                    result_lines.append(f"{teacher_name}: месяц {month_percent:.2f}%, неделя {week_percent:.2f}%")          
        return "\n".join(result_lines)
    else:
        return df



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
        bot.send_message(message.from_user.id, "Это бот для учебной части колледжа IT Top. На данный момент функционал в процессе разработки.")
    elif message.text == "/help":
        bot.send_message(message.from_user.id, "Список доступных команд: \n/info - информация о боте. \n/help - список доступных команд.\n/checked_homework - получить отчет по проверяемым домашним заданиям.")
    elif message.text == "/checked_homework":
        bot.send_message(message.from_user.id, "Получаем данные...")
        bot.send_message(message.from_user.id, "Преподаватели с низкой проверяемостью домашних заданий: "+checked_homework());
    else:
        bot.send_message(message.from_user.id, "Напишите /help, чтобы получить список команд.")



bot.polling(none_stop=True, interval=0)