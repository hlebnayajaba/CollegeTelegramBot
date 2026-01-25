import os
import pandas as pd
import telebot
from Config import EXCEL_EXTENSIONS, CATEGORY_MAP, CATEGORY_FILES
from ExcelUtils import get_info

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
                subject_name = row.iloc[2]
                lesson_topic = row.iloc[5]
                if lesson_topic.startswith('Урок №') and 'Тема:' in lesson_topic:
                    continue               
                result_lines.append(f"Дата: {lesson_date}. Предмет: {subject_name}. Тема занятия: {lesson_topic}.")
        return "\n".join(result_lines)
    else:
        return df

#расписание 
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
                                end_idx = lesson_text.find("\n", start_idx)
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
