import pandas as pd
import ast
import re

# CSV-версия Google-таблицы
url = "https://docs.google.com/spreadsheets/d/1M-CdODvpVuRlVHqheTJMZFcjAs95ef_uI7fOZX3az_E/gviz/tq?tqx=out:csv"

# загрузка данных
try:
    df = pd.read_csv(url)
except Exception as e:
    print(f"Ошибка загрузки данных: {e}")
    exit()

# очистка текста
def clean_text(text):
    text = re.sub(r"\[.*?\]", "", text) 
    text = re.sub(r"[-]{2,}.*?[-]{2,}", "", text) 
    text = re.sub(r"[\n\r\b\t]", " ", text) 
    text = re.sub(r"[^a-zA-Zа-яА-ЯёЁ0-9\s.,?!\-()\"';:]", "", text, flags=re.UNICODE) 
    return text.strip() 

# функция обработки 
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

if "Вопрос" in df.columns:
    df["Вопрос"] = df["Вопрос"].fillna("")  # Заполняем пропуски пустыми строками

    # извлекаем сообщения
    operator_messages = extract_messages(df["Вопрос"], "Оператор")
    client_messages = extract_messages(df["Вопрос"], "Клиент")

    # создаем новые таблицы DataFrame с одной колонкой history разделенной на сообщения, убираем пустые строки
    operator_df = pd.DataFrame({"history": [msg for msg in operator_messages if msg.strip()]})
    client_df = pd.DataFrame({"history": [msg for msg in client_messages if msg.strip()]})

    # сохраняем в файлы
    operator_df.to_excel("./operator_messages.xlsx", index=False)
    client_df.to_excel("./client_messages.xlsx", index=False)
else:
    print("Столбец 'Вопрос' не найден в таблице.")
