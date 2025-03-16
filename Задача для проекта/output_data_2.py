import pandas as pd # библиотека pandas для работы с (таблицами - DataFrame, массивами данных)
import ast # модуль ast для преобразования строки в список
import re # модуль re для работы с регулярными выражениями

# загрузка Excel-файла
file_path = "./output_data.xlsx" # путь к файлу 
df = pd.read_excel(file_path, sheet_name="Sheet1") # читаем таблицу и загружаем данные из листа Sheet1 в новую таблицу DataFrame

# очистка текста 
def clean_text(text):
    text = re.sub(r"\[.*?\]", "", text)  # убираем содержимое в квадратных скобках
    text = re.sub(r"[-]{2,}.*?[-]{2,}", "", text)  # убираем строки с символами ----
    text = re.sub(r"[\n\r\b\t]", " ", text)  # убираем спецсимволы
    text = re.sub(r"[^a-zA-Zа-яА-ЯёЁ0-9\s.,?!\-()\"';:]", "", text, flags=re.UNICODE)  # оставляем только буквы, цифры, пробелы и знаки препинания
    return text.strip() # убираем лишние пробелы по краям строки

# выбираем сообщения из столбца history и фильтруем по ролям Оператор, Клиент
def extract_messages(history, role):
    messages = [] # пустой список
    for entry in history: # проходим по каждому элементу столбца history
        try:
            logs = ast.literal_eval(entry) if isinstance(entry, str) else []  # преобразуем строку в список
            extracted = [clean_text(log.split('] ', 2)[-1]) for log in logs if f"[{role}]" in log] #разбиваем строку по (']') и берем само сообщение, если оно принадлежит нужной роли
            messages.extend(extracted)  # каждое сообщение в отдельную строку
        except (ValueError, SyntaxError): # если ошибка в синтаксисе или преобразовании строки, пропускаем элемент
            continue  # Пропускаем ошибочные записи
    return messages

# обновляем только столбец history, заполняя пропуски пустыми строками
operator_messages = extract_messages(df["history"].fillna(""), "Оператор")  # получаем сообщения оператора
client_messages = extract_messages(df["history"].fillna(""), "Клиент")  # получаем сообщения клиента

# Убираем пустые строки из сообщений
operator_messages = [msg for msg in operator_messages if msg.strip()]
client_messages = [msg for msg in client_messages if msg.strip()]

# Создаем новые таблицы DataFrame с одной колонкой history
operator_df = pd.DataFrame({"history": operator_messages})
client_df = pd.DataFrame({"history": client_messages})

# Сохраняем в новые файлы
operator_df.to_excel("./operator_messages.xlsx", index=False)
client_df.to_excel("./client_messages.xlsx", index=False)
