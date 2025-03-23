import pandas as pd 
import ast 
import re 

# загрузка Excel-файла
file_path = "./output_data.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1") 

# очистка текста 
def clean_text(text):
    text = re.sub(r"\[.*?\]", "", text)  
    text = re.sub(r"[-]{2,}.*?[-]{2,}", "", text)  
    text = re.sub(r"[\n\r\b\t]", " ", text)  
    text = re.sub(r"[^a-zA-Zа-яА-ЯёЁ0-9\s.,?!\-()\"';:]", "", text, flags=re.UNICODE)  
    return text.strip() 

# выбираем сообщения из столбца history и фильтруем по ролям Оператор, Клиент
def extract_messages(history, role):
    messages = [] 
    for entry in history:
        try:
            logs = ast.literal_eval(entry) if isinstance(entry, str) else []  
            extracted = [clean_text(log.split('] ', 2)[-1]) for log in logs if f"[{role}]" in log]
            messages.extend(extracted)  
        except (ValueError, SyntaxError): 
            continue  
    return messages

# обновляем только столбец history, заполняя пропуски пустыми строками
operator_messages = extract_messages(df["history"].fillna(""), "Оператор")  
client_messages = extract_messages(df["history"].fillna(""), "Клиент")  

# Убираем пустые строки из сообщений
operator_messages = [msg for msg in operator_messages if msg.strip()]
client_messages = [msg for msg in client_messages if msg.strip()]

# Создаем новые таблицы DataFrame с одной колонкой history
operator_df = pd.DataFrame({"history": operator_messages})
client_df = pd.DataFrame({"history": client_messages})

# Сохраняем в новые файлы
operator_df.to_excel("./operator_messages.xlsx", index=False)
client_df.to_excel("./client_messages.xlsx", index=False)
