import telebot
import os
import pandas as pd
import io
bot = telebot.TeleBot('8497231406:AAE7bNFUgkzgTS4t5tzrEBB26CrKZ8dG96o')


EXCEL_EXTENSIONS = ['.xls', '.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm', '.xlam', '.xla', '.xlw']
#функция для извлечения данных из таблиц
def get_info(filename):
    try:
        file_path = os.path.join("Sheets", filename)
        df = pd.read_excel(file_path)

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

#посещаемость по преподавателям
def attendance_by_teachers():
    df = get_info("Посещаемость по преподавателям.xlsx")
    if type(df)!=str:
        result_lines = []
        for index, row in df.iloc[2:].iterrows():
            if pd.notna(row.iloc[0]) and row.iloc[0].strip() !="":
                teacher_name = row.iloc[0]
                attendance_str = row.iloc[10]
                if isinstance(attendance_str, str):
                    attendance_str = attendance_str.replace('%', '').strip()
                    try:
                        average_attendance = float(attendance_str)
                    except ValueError:
                        continue
                else:
                    average_attendance = float(attendance_str)
                if average_attendance < 40:
                    result_lines.append(f"{teacher_name}: {average_attendance}%")
        return "\n".join(result_lines)
    else:
        return df

#оценки студентов
def student_review():
    df = get_info("Отчет по студентам.xls")
    if type(df)!=str:
        result_lines = []
        for index, row in df.iloc[1:].iterrows():
            if pd.notna(row.iloc[0]) and row.iloc[0].strip() !="":
                student_name = row.iloc[0]
                homework_score = row.iloc[15]
                class_score = row.iloc[16]
                if homework_score == 1 or class_score <= 3:
                    result_lines.append(f"{student_name}: Средняя оценка за домашнюю работу: {homework_score}, за классную работу: {class_score}")
        return "\n".join(result_lines)
    else:
        return df

#выполненные домашки
def completed_homeworks():
    df = get_info("Отчет по студентам.xls")
    if type(df)!=str:
        result_lines = []
        for index, row in df.iloc[1:].iterrows():
            if pd.notna(row.iloc[0]) and row.iloc[0].strip() !="":
                student_name = row.iloc[0]
                homework_percent = row.iloc[19]
                if homework_percent <= 70:
                    result_lines.append(f"{student_name}: {homework_percent}%")
        return "\n".join(result_lines)
    else:
        return df

# Функция для разбивки длинных сообщений на части
def split_message(message, max_length=4000):
    parts = []
    while len(message) > max_length:
        split_index = message.rfind('\n', 0, max_length)
        if split_index == -1:
            split_index = max_length
        parts.append(message[:split_index])
        message = message[split_index:].lstrip()
    parts.append(message)
    return parts

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
        bot.send_message(message.from_user.id, "Список доступных команд: \n/info - информация о боте. \n/help - список доступных команд.\n/checked_homework - получить отчет по проверяемым домашним заданиям. \n /attendance_by_teachers - получить отчет по посещаемости среди преподавателей \n/student_review - получить отчет по успеваемости студентов\n/completed_homeworks - получить отчет по выполненным домашним заданиям")
    elif message.text == "/checked_homework":
        bot.send_message(message.from_user.id, "Получаем данные...")
        result = checked_homework()
        parts = split_message(f"Преподаватели с низкой проверяемостью домашних заданий:\n{result}")
        for part in parts:
            bot.send_message(message.from_user.id, part)
    elif message.text == "/attendance_by_teachers":
        bot.send_message(message.from_user.id, "Получаем данные...")
        result = attendance_by_teachers()
        parts = split_message(f"Преподаватели с низкой посещаемостью:\n{result}")
        for part in parts:
            bot.send_message(message.from_user.id, part)
    elif message.text == "/student_review":
        bot.send_message(message.from_user.id, "Получаем данные...")
        result = student_review()
        parts = split_message(f"Студенты с низкой успеваемостью:\n{result}")
        for part in parts:
            bot.send_message(message.from_user.id, part)
    elif message.text == "/completed_homeworks":
        bot.send_message(message.from_user.id, "Получаем данные...")
        result = completed_homeworks()
        parts = split_message(f"Студенты с низким количеством выполненных домашних заданий:\n{result}")
        for part in parts:
            bot.send_message(message.from_user.id, part)
    else:
        bot.send_message(message.from_user.id, "Напишите /help, чтобы получить список команд.")

bot.polling(none_stop=True, interval=0)