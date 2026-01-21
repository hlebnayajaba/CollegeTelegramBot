import datetime
from pickle import NONE
import telebot
import os
import pandas as pd
import io

bot = telebot.TeleBot('8497231406:AAE7bNFUgkzgTS4t5tzrEBB26CrKZ8dG96o')

user_states = {}
user_data = {} 
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
        return f"Ошибка при чтении файла: {str(e)}"

#функция для обновления информации
def update_info(file_bytes, file_path):
    try:
        directory = os.path.dirname(file_path)
        if directory and not os.path.exists(directory):
            os.makedirs(directory, exist_ok=True)
        with open(file_path, 'wb') as f:
            f.write(file_bytes)        
        return f"Файл успешно сохранен по пути: {file_path}"
    except Exception as e:
        return f"Ошибка при сохранении файла: {str(e)}"

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

#темы уроков
def lesson_themes():
    df = get_info("Темы уроков.xls")
    if type(df)!=str:
        result_lines = []
        for index, row in df.iloc[1:].iterrows():
            if pd.notna(row.iloc[0]) and row.iloc[0].strip() !="":
                lesson_date = row.iloc[0]
                group_num = row.iloc[3]
                teacher_name = row.iloc[4]
                lesson_topic = row.iloc[5]
                if lesson_topic.startswith('Урок №') and 'Тема:' in lesson_topic:
                    continue               
                result_lines.append(f"Дата: {lesson_date}. Группа: {group_num}. Имя преподавателя: {teacher_name}. Тема занятия: {lesson_topic}.")
        return "\n".join(result_lines)
    else:
        return df

def check_schedule(group_name):
    file_group_name = group_name.replace("/", "-")
    schedule_file = None
    
    for ext in EXCEL_EXTENSIONS:
        file_path = os.path.join("Sheets", "Расписание групп", f"{file_group_name}{ext}")
        if os.path.exists(file_path):
            schedule_file = f"Расписание групп/{file_group_name}{ext}"
            break

    if not schedule_file:
        return "Расписание для указанной группы не найдено."
    else:
        df = get_info(schedule_file)
        if type(df) != str:
            subject_list = {}
            weekdays = [3, 5, 7, 9, 11, 13]
            
            for index, row in df.iterrows():
                if pd.isna(row.iloc[1]) or row.iloc[1] == 0:
                    continue
                
                for day in weekdays:
                    if day < len(row):
                        lesson_cell = row.iloc[day]
                        if pd.isna(lesson_cell) or str(lesson_cell).strip() == "":
                            continue
                        
                        lesson_text = str(lesson_cell)
                        if "Предмет:" in lesson_text:
                            start_idx = lesson_text.find("Предмет:") + len("Предмет:")
                            if start_idx < len(lesson_text) and lesson_text[start_idx] == " ":
                                start_idx += 1
                            
                            end_idx = lesson_text.find("<br>", start_idx)
                            if end_idx == -1:
                                end_idx = len(lesson_text)
                            
                            subject_name = lesson_text[start_idx:end_idx].strip()
                            
                            if subject_name in subject_list:
                                subject_list[subject_name] += 1
                            else:
                                subject_list[subject_name] = 1
            
            result_lines = []
            for subject, count in sorted(subject_list.items()):
                result_lines.append(f"{subject}: {count} пар")
            
            return "\n".join(result_lines)
        else:
            return df


#функция для разбивки длинных сообщений на части
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
    if message.from_user.id not in user_states:
        user_states[message.from_user.id] = None
    if message.from_user.id not in user_data:
        user_data[message.from_user.id] = {}

    if message == "/start":
        bot.send_message(message.from_user.id, "Добро пожаловать!\nЭто бот для учебной части филиала колледжа IT Top города Нижний Новгород.\nНапишите /help, чтобы получить список команд.")

    elif message.document:
        file_ext = os.path.splitext(message.document.file_name)[1].lower()        
        if file_ext not in EXCEL_EXTENSIONS:
            bot.send_message(message.from_user.id, "Формат файла не подходит. Отправьте файл Excel (.xls, .xlsx и т.д.)")
            return        
        if user_states[message.from_user.id] == 'waiting_for_file':
            try:
                file_info = bot.get_file(message.document.file_id)
                downloaded_file = bot.download_file(file_info.file_path)
                category = user_data[message.from_user.id].get('category', 'unknown')
                group_name = user_data[message.from_user.id].get('group_name', '')
                if category == 'расписание групп' and group_name:
                    file_group_name = group_name.replace("/", "-")
                    file_name = f"{file_group_name}.xlsx"
                    save_path = os.path.join("Sheets", "Расписание групп", file_name)
                else:
                    category_files = {
                        'отчет по студентам': 'Отчет по студентам.xls',
                        'отчет по домашним заданиям у преподавателей': 'Отчет по домашним заданиям.xlsx',
                        'темы уроков': 'Темы уроков.xls',
                        'посещаемость по преподавателям': 'Посещаемость по преподавателям.xlsx'
                    }
                    file_name = category_files.get(category, 'uploaded_file.xlsx')
                    save_path = os.path.join("Sheets", file_name)
                result = update_info(downloaded_file, save_path)
                bot.send_message(message.from_user.id, f"Файл '{message.document.file_name}' принят.\n{result}")
                user_states[message.from_user.id] = None
                user_data[message.from_user.id] = {}
                
            except Exception as e:
                bot.send_message(message.from_user.id, f"Ошибка при обработке файла: {str(e)}")
        else:
            bot.send_message(message.from_user.id, "Сначала используйте команду /send_data для начала загрузки файла.")

    elif message.text == "/info":
        bot.send_message(message.from_user.id, "Это бот для учебной части колледжа IT Top. На данный момент функционал в процессе разработки.")

    elif message.text == "/help":
        bot.send_message(message.from_user.id, "Список доступных команд: \n/info - информация о боте. \n/help - список доступных команд.\n/checked_homework - получить отчет по проверяемым домашним заданиям. \n /attendance_by_teachers - получить отчет по посещаемости среди преподавателей \n/student_review - получить отчет по успеваемости студентов\n/completed_homeworks - получить отчет по выполненным домашним заданиям\n/lesson_themes - получить отчет по неправильно написанным темам.\n/check_schedule - проверить расписание для группы\n/send_data - отправить файл с данными\n/cancel - отменить текущую операцию")

    elif message.text == "/check_schedule":
        user_states[message.from_user.id] = 'waiting_for_group_name'
        bot.send_message(message.from_user.id, "Введите название группы (например: 9/3-РПО-23/2) или /cancel, чтобы отменить операцию:")
    
    elif user_states[message.from_user.id] == 'waiting_for_group_name' and message.text:
        group_name = message.text.strip()
        bot.send_message(message.from_user.id, f"Получаем расписание для группы {group_name}...")
        
        result = check_schedule(group_name)
        if result:
            parts = split_message(f"Количество пар по дисциплинам для группы {group_name}:\n{result}")
            for part in parts:
                bot.send_message(message.from_user.id, part)
        
        user_states[message.from_user.id] = None
        user_data[message.from_user.id] = {}

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

    elif message.text == "/lesson_themes":
        bot.send_message(message.from_user.id, "Получаем данные...")
        result = lesson_themes()
        if result and result.strip():
            parts = split_message(f"Темы уроков, не подходящие под формат:{result}")
            for part in parts:
                bot.send_message(message.from_user.id, part)

    elif message.text == "/send_data":
        user_states[message.from_user.id] = 'waiting_for_category'
        bot.send_message(message.from_user.id, "Выберите категорию файла:\n1. Отчет по студентам\n2. Отчет по домашним заданиям у преподавателей\n3. Темы уроков\n4. Посещаемость по преподавателям\n5. Расписание групп\nВведите номер категории или ее название, либо /cancel, чтобы отменить операцию.")
    
    elif user_states[message.from_user.id] == 'waiting_for_category' and message.text:
        category_input = message.text.strip().lower()        
        category_map = {
            '1': 'отчет по студентам',
            '2': 'отчет по домашним заданиям у преподавателей', 
            '3': 'темы уроков',
            '4': 'посещаемость по преподавателям',
            '5': 'расписание групп',
            'отчет по студентам': 'отчет по студентам',
            'отчет по домашним заданиям': 'отчет по домашним заданиям у преподавателей',
            'отчет по домашним заданиям у преподавателей': 'отчет по домашним заданиям у преподавателей',
            'темы уроков': 'темы уроков',
            'посещаемость': 'посещаемость по преподавателям',
            'посещаемость по преподавателям': 'посещаемость по преподавателям',
            'расписание': 'расписание групп',
            'расписание групп': 'расписание групп'
        }
        
        category = category_map.get(category_input)
        if category:
            user_data[message.from_user.id]['category'] = category
            
            if category == 'расписание групп':
                user_states[message.from_user.id] = 'waiting_for_group_for_schedule'
                bot.send_message(message.from_user.id, "Введите название группы для расписания (например: 9/3-РПО-23/2):")
            else:
                user_states[message.from_user.id] = 'waiting_for_file'
                bot.send_message(message.from_user.id, f"Вы выбрали: {category}. Теперь отправьте Excel файл.")
        else:
            bot.send_message(message.from_user.id, "Неверный выбор категории. Пожалуйста, введите номер от 1 до 5 или название категории.")
    
    elif user_states[message.from_user.id] == 'waiting_for_group_for_schedule' and message.text:
        group_name = message.text.strip()
        user_data[message.from_user.id]['group_name'] = group_name
        user_states[message.from_user.id] = 'waiting_for_file'
        bot.send_message(message.from_user.id, f"Группа для расписания: {group_name}. Теперь отправьте Excel файл с расписанием.")

    elif message.text == "/cancel":
        user_states[message.from_user.id] = None
        user_data[message.from_user.id] = {}
        bot.send_message(message.from_user.id, "Операция отменена.")
        return

    else:
        bot.send_message(message.from_user.id, "Напишите /help, чтобы получить список команд.")

bot.polling(none_stop=True, interval=0)