import pandas as pd
import os
import time

def compare_excel_sheets(file_path, sheet1_name, sheet2_name):
    print("Начало выполнения...")

    # Этап 1: Проверка существования файла
    print("Этап 1/5: Проверка файла...")
    if not os.path.exists(file_path):
        print(f"Ошибка: Файл {file_path} не найден")
        return
    time.sleep(0.5)

    # Этап 2: Проверка листов
    print("Этап 2/5: Проверка листов...")
    try:
        excel_file = pd.ExcelFile(file_path)
        available_sheets = excel_file.sheet_names
        print(f"Доступные листы в файле: {available_sheets}")
        if sheet1_name not in available_sheets:
            print(f"Ошибка: Лист '{sheet1_name}' не найден")
            return
        if sheet2_name not in available_sheets:
            print(f"Ошибка: Лист '{sheet2_name}' не найден")
            return
    except Exception as e:
        print(f"Ошибка при чтении файла: {e}")
        return
    time.sleep(0.5)

    # Этап 3: Чтение данных
    print("Этап 3/5: Чтение данных из листов...")
    try:
        df1 = pd.read_excel(file_path, sheet_name=sheet1_name, usecols=[0])
        df2 = pd.read_excel(file_path, sheet_name=sheet2_name, usecols=[0])
        print(f"Прочитано {len(df1)} строк из {sheet1_name}, {len(df2)} строк из {sheet2_name}")
    except Exception as e:
        print(f"Ошибка при чтении листов: {e}")
        return
    time.sleep(0.5)

    # Извлечение значений из первых столбцов
    col1_sheet1 = df1.iloc[:, 0].dropna().tolist()
    col1_sheet2 = df2.iloc[:, 0].dropna().tolist()

    # Этап 4: Сравнение данных
    print("Этап 4/5: Сравнение данных...")
    # Различия: в List1, но не в List2
    diff_list1_not_in_list2 = [x for x in col1_sheet1 if x not in col1_sheet2]
    # Различия: в List2, но не в List1
    diff_list2_not_in_list1 = [x for x in col1_sheet2 if x not in col1_sheet1]
    # Общие значения
    common_values = [x for x in col1_sheet1 if x in col1_sheet2]

    print(f"Найдено {len(diff_list1_not_in_list2)} значений в {sheet1_name}, отсутствующих в {sheet2_name}")
    print(f"Найдено {len(diff_list2_not_in_list1)} значений в {sheet2_name}, отсутствующих в {sheet1_name}")
    print(f"Найдено {len(common_values)} общих значений")
    time.sleep(0.5)

    # Этап 5: Сохранение результатов
    print("Этап 5/5: Сохранение результатов...")
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
            # Сохранение различий List1 → List2
            pd.DataFrame(diff_list1_not_in_list2, columns=['В List1, но не в List2']).to_excel(
                writer, sheet_name='Diff_List1_not_in_List2', index=False
            )
            # Сохранение различий List2 → List1
            pd.DataFrame(diff_list2_not_in_list1, columns=['В List2, но не в List1']).to_excel(
                writer, sheet_name='Diff_List2_not_in_List1', index=False
            )
            # Сохранение общих значений
            pd.DataFrame(common_values, columns=['Общие значения']).to_excel(
                writer, sheet_name='Common', index=False
            )
        print(f"Результаты сохранены в файле {file_path} на листах 'Diff_List1_not_in_List2', 'Diff_List2_not_in_List1' и 'Common'")
    except ValueError as e:
        print(f"Ошибка: Один из листов ('Diff_List1_not_in_List2', 'Diff_List2_not_in_List1', 'Common') уже существует. Удалите их или измените имена листов в коде.")
    except Exception as e:
        print(f"Ошибка при сохранении: {e}")
    print("Выполнение завершено!")

# Параметры
file_path = 'C:/python/exel_compare/data.xlsx'  # Укажите полный путь к вашему файлу
sheet1_name = 'List1'    # Имя первого листа
sheet2_name = 'List2'    # Имя второго листа

# Запуск
compare_excel_sheets(file_path, sheet1_name, sheet2_name)
