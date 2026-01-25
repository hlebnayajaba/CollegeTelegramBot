import datetime
from pickle import NONE
import telebot
import os
import io
from Config import BOT_TOKEN, EXCEL_EXTENSIONS, CATEGORY_MAP, CATEGORY_FILES
from ExcelUtils import get_info, update_info
from BotFunctions import (
    checked_homework, attendance_by_teachers, student_review,
    completed_homeworks, lesson_themes, check_schedule, split_message
)

bot = telebot.TeleBot(BOT_TOKEN)

user_states = {}
user_data = {} 


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
                    file_name = CATEGORY_FILES.get(category, 'uploaded_file.xlsx')
                    save_path = os.path.join("Sheets", file_name)
                result = update_info(downloaded_file, save_path)
                bot.send_message(message.from_user.id, f"Файл '{message.document.file_name}' принят.\n{result}")
                user_states[message.from_user.id] = None
                user_data[message.from_user.id] = {}
                
            except Exception as e:
                bot.send_message(message.from_user.id, f"Ошибка при обработке файла: {str(e)}")
        else:
            bot.send_message(message.from_user.id, "Сначала используйте команду /send_data для начала загрузки файла.")

    elif message.text == "/help":
        bot.send_message(message.from_user.id, "Список доступных команд: \n/help - список доступных команд.\n/checked_homework - получить отчет по проверяемым домашним заданиям. \n /attendance_by_teachers - получить отчет по посещаемости среди преподавателей \n/student_review - получить отчет по успеваемости студентов\n/completed_homeworks - получить отчет по выполненным домашним заданиям\n/lesson_themes - получить отчет по неправильно написанным темам.\n/check_schedule - проверить расписание для группы\n/send_data - отправить файл с данными\n/cancel - отменить текущую операцию")

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
            parts = split_message(f"Темы уроков, не подходящие под формат:\n{result}")
            for part in parts:
                bot.send_message(message.from_user.id, part)

    elif message.text == "/send_data":
        user_states[message.from_user.id] = 'waiting_for_category'
        bot.send_message(message.from_user.id, "Выберите категорию файла:\n1. Отчет по студентам\n2. Отчет по домашним заданиям у преподавателей\n3. Темы уроков\n4. Посещаемость по преподавателям\n5. Расписание групп\nВведите номер категории или ее название, либо /cancel, чтобы отменить операцию.")
    
    elif user_states[message.from_user.id] == 'waiting_for_category' and message.text:
        category_input = message.text.strip().lower()        
        category = CATEGORY_MAP.get(category_input)
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