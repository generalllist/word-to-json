import os
import json
from docx import Document

print("✅ Готово: python-docx доступен")

# Получаем все .docx-файлы в текущей папке
current_dir = os.getcwd()
docx_files = [f for f in os.listdir(current_dir) if f.endswith('.docx')]

if not docx_files:
    print("⚠️ В папке не найдено ни одного .docx-файла.")
    exit()

print(f"📄 Найдено {len(docx_files)} .docx-файлов: {docx_files}")

# Жёстко заданный порядок заголовков — как в вашей таблице
EXPECTED_HEADERS = [
    "event",
    "status",
    "psychophysiological_state",
    "circumstances",
    "pattern_of_behavior"
]

for filename in docx_files:
    print(f"\n🔄 Обработка: {filename}")
    try:
        doc = Document(filename)
        all_data = []

        for table in doc.tables:
            # Пропускаем таблицы без строк
            if len(table.rows) < 1:
                continue

            # Мы НЕ используем первую строку как заголовки!
            # Вместо этого — жёстко задаём порядок колонок (по вашему требованию)
            # Нумерация строк начинается со второй строки таблицы (после заголовков)

            for row_idx, row in enumerate(table.rows[1:], start=2):  # start=2 → первая строка данных = строка 2
                # Пропускаем полностью пустые строки (опционально)
                if all(cell.text.strip() == "" for cell in row.cells):
                    continue

                record = {}
                # Обрабатываем ровно 5 колонок — даже если в строке меньше ячеек
                for i, header in enumerate(EXPECTED_HEADERS):
                    if i < len(row.cells):
                        value = row.cells[i].text.strip()
                    else:
                        value = ""  # если ячеек меньше 5 — ставим пустую строку
                    record[header] = value

                record["_source_file"] = filename
                record["_row_number"] = row_idx
                all_data.append(record)

        # Генерируем имя JSON-файла
        base_name = os.path.splitext(filename)[0]
        safe_name = "".join(c if c.isalnum() or c in "._-" else "_" for c in base_name)
        json_filename = f"{safe_name}.json"

        # Сохраняем
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(all_data, f, indent=2, ensure_ascii=False)

        print(f"  ✅ Сохранено: {json_filename} ({len(all_data)} записей)")

    except Exception as e:
        print(f"❌ Ошибка при обработке {filename}: {e}")

print(f"\n🎉 Готово! Обработано {len(docx_files)} файлов.")
