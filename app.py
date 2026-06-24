# ============================================
# БЛОК 0. ИМПОРТЫ И НАСТРОЙКА
# ============================================
import streamlit as st
import os
import re
import json
import time
from typing import List, Dict, Optional
from docx import Document
from openpyxl import load_workbook
from openai import OpenAI
import pandas as pd

# ============================================
# БЛОК 1. ЗАГРУЗКА ФАЙЛОВ ИЗ ПАПКИ
# ============================================
def detect_file_type(file_path: str) -> Optional[str]:
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.docx': return 'docx'
    elif ext == '.xlsx': return 'xlsx'
    elif ext == '.txt': return 'txt'
    return None

def read_docx(file_path: str, column_name: str = 'PropertyValue') -> List[Dict]:
    doc = Document(file_path)
    records = []
    row_counter = 0
    for table in doc.tables:
        if not table.rows: continue
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        col_idx = next((i for i, h in enumerate(headers) if column_name.lower() in h.lower()), None)
        if col_idx is None: continue
        for row in table.rows[1:]:
            row_counter += 1
            cells = row.cells
            if col_idx < len(cells):
                text = cells[col_idx].text.strip()
                if text:
                    records.append({
                        'row_num': row_counter,
                        'text': text,
                        'source': f"{os.path.basename(file_path)}, строка {row_counter}",
                        'patient_id': None
                    })
    return records

def read_xlsx(file_path: str, column_name: str = 'PropertyValue') -> List[Dict]:
    wb = load_workbook(file_path, data_only=True)
    records = []
    row_counter = 0
    for sheet in wb.worksheets:
        first_row = list(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
        if not first_row: continue
        headers = [str(cell) if cell else "" for cell in first_row[0]]
        id_col_idx = next((i for i, h in enumerate(headers) if 'propertyid' in h.lower()), None)
        val_col_idx = next((i for i, h in enumerate(headers) if column_name.lower() in h.lower()), None)
        if val_col_idx is None: continue
        for row in sheet.iter_rows(min_row=2, values_only=True):
            row_counter += 1
            if val_col_idx < len(row) and row[val_col_idx]:
                text = str(row[val_col_idx]).strip()
                if text:
                    patient_id = str(row[id_col_idx]) if id_col_idx is not None and id_col_idx < len(row) and row[id_col_idx] else None
                    records.append({
                        'row_num': row_counter,
                        'text': text,
                        'source': f"{os.path.basename(file_path)}, лист {sheet.title}, строка {row_counter}",
                        'patient_id': patient_id
                    })
    return records

def read_txt(file_path: str) -> List[Dict]:
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    records = []
    for i, line in enumerate(lines, 1):
        text = line.strip()
        if text:
            records.append({
                'row_num': i,
                'text': text,
                'source': f"{os.path.basename(file_path)}, строка {i}",
                'patient_id': None
            })
    return records

def load_folder(folder_path: str = 'input', column_name: str = 'PropertyValue') -> List[Dict]:
    os.makedirs(folder_path, exist_ok=True)
    all_records = []
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        ft = detect_file_type(file_path)
        if ft == 'docx':
            all_records.extend(read_docx(file_path, column_name))
        elif ft == 'xlsx':
            all_records.extend(read_xlsx(file_path, column_name))
        elif ft == 'txt':
            all_records.extend(read_txt(file_path))
    return all_records

# ============================================
# БЛОК 2. АДАПТЕР ДЛЯ OPENROUTER
# ============================================
class OpenRouterAdapter:
    def __init__(self, model: str, api_key: str = None, timeout: int = 300):
        self.model = model
        self.api_key = api_key
        if not self.api_key:
            raise ValueError("API-ключ OpenRouter не указан")
        self.client = OpenAI(
            base_url="https://openrouter.ai/api/v1",
            api_key=self.api_key,
            timeout=timeout
        )

    def generate(self, system_prompt: str, user_prompt: str, **kwargs) -> str:
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": user_prompt})
        response = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            temperature=kwargs.get("temperature", 0),
        )
        return response.choices[0].message.content

# ============================================
# БЛОК 3. ИЗВЛЕЧЕНИЕ ПАР
# ============================================
def extract_pairs(text: str, llm) -> list:
    if len(text) > 1500:
        text = text[:1500] + "..."
    prompt = f"""
Извлеки из текста все медицинские термины (симптомы, диагнозы, состояния) и их наличие/значение.
Верни JSON-массив массивов, каждый подмассив из двух строк: ["термин", "значение"].
Возможные значения: присутствует, отсутствует, усиливается, ослабевает, постоянный, периодический и т.п.
Только JSON, без пояснений.

Текст:
{text}
"""
    content = llm.generate(
        system_prompt="Ты медицинский эксперт. Отвечай только JSON.",
        user_prompt=prompt,
        temperature=0
    )
    content = content.strip().removeprefix("```json").removesuffix("```").strip()
    data = json.loads(content)
    if isinstance(data, dict) and "terms" in data:
        data = data["terms"]
    return data

# ============================================
# БЛОК 4. ДОБАВЛЕНИЕ ХАРАКТЕРИСТИК
# ============================================
def enrich_with_characteristics(pairs: list, llm) -> list:
    unique_values = set(val for _, val in pairs)
    value_to_char = {}

    for value in unique_values:
        prompt = f"""
Определи категорию (характеристику) для медицинского значения "{value}".
Категория должна быть существительным или короткой фразой, обобщающей тип значения.
Примеры:
- "присутствует" → "присутствие"
- "отсутствует" → "отсутствие"
- "усиливается" → "динамика"
- "постоянный" → "характер"
- "ноющая" → "интенсивность" (или "характер боли")
- "при нагрузке" → "условие возникновения"
- "5 мг" → "дозировка"

Верни ТОЛЬКО категорию, без пояснений. Одно слово или короткая фраза.
"""
        try:
            response = llm.generate(
                system_prompt="Ты эксперт по медицинской терминологии.",
                user_prompt=prompt,
                temperature=0
            )
            char = response.strip().strip('"').strip("'")
            if not char or len(char) > 100:
                char = "характеристика"
            value_to_char[value] = char
        except Exception as e:
            value_to_char[value] = "характеристика"

    triples = []
    for term, value in pairs:
        char = value_to_char.get(value, "характеристика")
        triples.append((term, char, value))
    return triples

# ============================================
# БЛОК 5. ОСНОВНОЙ ПАЙПЛАЙН
# ============================================
def run_pipeline(folder: str = "input", column: str = "PropertyValue",
                 model: str = "nvidia/nemotron-3-ultra-550b-a55b:free",
                 api_key: str = None, timeout: int = 300):
    llm = OpenRouterAdapter(api_key=api_key, model=model, timeout=timeout)
    os.makedirs(folder, exist_ok=True)
    data = load_folder(folder_path=folder, column_name=column)
    if not data:
        return None, None, None
    all_pairs = []
    for rec in data:
        pairs = extract_pairs(rec['text'], llm)
        all_pairs.extend([{
            'term': t,
            'value': v,
            'source': rec['source'],
            'source_text': rec['text'],
            'patient_id': rec.get('patient_id')
        } for t, v in pairs])
    if not all_pairs:
        return None, None, None
    pair_list = [(p['term'], p['value']) for p in all_pairs]
    triples = enrich_with_characteristics(pair_list, llm)
    enriched = [{
        'term': t,
        'characteristic': c,
        'value': v,
        'source': p['source'],
        'source_text': p['source_text'],
        'patient_id': p['patient_id']
    } for (t, c, v), p in zip(triples, all_pairs)]
    return data, all_pairs, enriched

# ============================================
# БЛОК 6. STREAMLIT ИНТЕРФЕЙС
# ============================================
st.set_page_config(page_title="Медицинский парсер", layout="centered")
st.title("🏥 Преобразование медицинских записей в датасет")
st.markdown("Загрузите файлы (Excel, Word, TXT) с колонкой **PropertyValue** и получите структурированный датасет.")

# Поле для API-ключа
api_key = st.text_input(
    "Введите ваш API-ключ OpenRouter",
    type="password",
    help="Получите ключ на openrouter.ai (бесплатно)"
)

# Загрузка файлов
uploaded_files = st.file_uploader(
    "Выберите файлы",
    accept_multiple_files=True,
    type=['xlsx', 'docx', 'txt']
)

if st.button("Обработать"):
    if not api_key:
        st.error("Пожалуйста, введите API-ключ.")
    elif not uploaded_files:
        st.error("Пожалуйста, загрузите хотя бы один файл.")
    else:
        # Создаём папку input
        input_dir = "input"
        os.makedirs(input_dir, exist_ok=True)
        
        # Очищаем от старых файлов
        for f in os.listdir(input_dir):
            os.remove(os.path.join(input_dir, f))
        
        # Сохраняем загруженные файлы
        for uploaded_file in uploaded_files:
            with open(os.path.join(input_dir, uploaded_file.name), "wb") as f:
                f.write(uploaded_file.getbuffer())
        
        with st.spinner("Идёт обработка... Это может занять несколько минут."):
            try:
                data, all_pairs, enriched = run_pipeline(
                    api_key=api_key,
                    model="nvidia/nemotron-3-ultra-550b-a55b:free"
                )
                
                if enriched is None:
                    st.error("Не удалось обработать данные. Проверьте файлы и ключ.")
                else:
                    # Сохраняем Excel в память
                    from io import BytesIO
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        pd.DataFrame(data).to_excel(writer, sheet_name='Исходные записи', index=False)
                        pd.DataFrame(all_pairs).to_excel(writer, sheet_name='Пары', index=False)
                        pd.DataFrame(enriched).to_excel(writer, sheet_name='Итог', index=False)
                        char_dict = {r['value']: r['characteristic'] for r in enriched}
                        pd.DataFrame(list(char_dict.items()), columns=['Значение', 'Характеристика']).to_excel(writer, sheet_name='Словарь характеристик', index=False)
                    output.seek(0)
                    
                    st.success("Обработка завершена!")
                    st.download_button(
                        label="📥 Скачать result.xlsx",
                        data=output,
                        file_name="result.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Ошибка: {e}")