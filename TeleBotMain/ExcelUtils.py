import os
import pandas as pd
#получение данных из файла
def get_info(filename, sheets_dir="Sheets"):
    try:
        file_path = os.path.join(sheets_dir, filename)
        df = pd.read_excel(file_path)

        if df.dropna(how='all').empty:
            return "Файл открывается, но все ячейки пустые"
        else:
            return df
    except Exception as e:
        return f"Ошибка при чтении файла: {str(e)}"
#сохранение данных в файл
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
