import os
import random
from openpyxl import load_workbook
from datetime import datetime, timedelta
from openpyxl.styles import Font, PatternFill, Alignment

def get_working_days_september_2025():
    """Возвращает список рабочих дней (пн-пт) сентября 2025 года"""
    working_days = []
    start_date = datetime(2025, 9, 1)  # 1 сентября 2025
    end_date = datetime(2025, 9, 30)   # 30 сентября 2025
    
    current_date = start_date
    while current_date <= end_date:
        # Проверяем, что это рабочий день (понедельник = 0, воскресенье = 6)
        if current_date.weekday() < 5:  # 0-4 = пн-пт
            working_days.append(current_date.strftime("%d.%m.%Y"))
        current_date += timedelta(days=1)
    
    return working_days

def add_dates_and_grades_to_excel_files_in_folders(base_folder):
    """
    В каждой папке внутри base_folder ищет файлы Excel (xlsx), где есть ФИО студентов,
    и добавляет столбцы с датами сентября 2025 (только рабочие дни) и случайными оценками.
    """
    # Получаем рабочие дни сентября 2025
    working_days = get_working_days_september_2025()
    
    # Стили для заголовков дат
    date_header_font = Font(bold=True, color="FFFFFF")
    date_header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    date_header_alignment = Alignment(horizontal="center", vertical="center")
    
    # Вероятность пропуска и стили для отметки "Н"
    absence_probability = 0.15  # 15% вероятность пропуска
    absence_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    absence_font = Font(bold=True, color="9C0006")
    
    # Распределение профилей студентов по типу оценок
    # only_5: только 5
    # only_4_5: только 4 и 5
    # no_4_5: только 2 и 3 (без 4 и 5)
    # mixed: любые 2-5
    prob_only_5 = 0.15
    prob_only_4_5 = 0.35
    prob_no_4_5 = 0.20
    # Остальное идёт на смешанный профиль
    
    for group_folder in os.listdir(base_folder):
        group_path = os.path.join(base_folder, group_folder)
        if os.path.isdir(group_path):
            for file in os.listdir(group_path):
                if file.endswith('.xlsx') and file != 'студенты.xlsx':  # Пропускаем файл со списком студентов
                    file_path = os.path.join(group_path, file)
                    try:
                        wb = load_workbook(file_path)
                        ws = wb.active
                        
                        # Проверяем, есть ли столбец "ФИО"
                        headers = [cell.value for cell in ws[1]]
                        if "ФИО" in headers:
                            # Находим последний столбец
                            last_col = ws.max_column
                            
                            # Добавляем заголовки дат
                            for i, date in enumerate(working_days):
                                col = last_col + 1 + i
                                cell = ws.cell(row=1, column=col, value=date)
                                cell.font = date_header_font
                                cell.fill = date_header_fill
                                cell.alignment = date_header_alignment
                            
                            # Добавляем оценки/пропуски с учётом профилей студентов
                            for row in range(2, ws.max_row + 1):
                                r = random.random()
                                if r < prob_only_5:
                                    profile = "only_5"
                                elif r < prob_only_5 + prob_only_4_5:
                                    profile = "only_4_5"
                                elif r < prob_only_5 + prob_only_4_5 + prob_no_4_5:
                                    profile = "no_4_5"
                                else:
                                    profile = "mixed"

                                for i, date in enumerate(working_days):
                                    col = last_col + 1 + i
                                    # Случайный пропуск "Н" с заданной вероятностью, иначе оценка по профилю
                                    if random.random() < absence_probability:
                                        cell = ws.cell(row=row, column=col, value="Н")
                                        cell.font = absence_font
                                        cell.fill = absence_fill
                                        cell.alignment = Alignment(horizontal="center", vertical="center")
                                    else:
                                        if profile == "only_5":
                                            grade = 5
                                        elif profile == "only_4_5":
                                            grade = 5 if random.random() < 0.5 else 4
                                        elif profile == "no_4_5":
                                            grade = 3 if random.random() < 0.5 else 2
                                        else:
                                            grade = random.randint(2, 5)
                                        ws.cell(row=row, column=col, value=grade)
                            
                            # Автоматически подгоняем ширину столбцов
                            for column in ws.columns:
                                max_length = 0
                                column_letter = column[0].column_letter
                                for cell in column:
                                    try:
                                        if len(str(cell.value)) > max_length:
                                            max_length = len(str(cell.value))
                                    except:
                                        pass
                                adjusted_width = min(max_length + 2, 15)
                                ws.column_dimensions[column_letter].width = adjusted_width
                            
                            wb.save(file_path)
                            print(f"В файл '{file}' в папке '{group_folder}' добавлены даты и оценки.")
                            
                    except Exception as e:
                        print(f"Ошибка при обработке файла {file_path}: {e}")

if __name__ == "__main__":
    # Используем папку с журналами, где находятся папки с группами
    base_folder = "Журналы/1 Курс"
    
    print("Добавляем даты сентября 2025 и оценки в файлы...")
    print("Рабочие дни сентября 2025:")
    
    # Показываем рабочие дни
    working_days = get_working_days_september_2025()
    for i, day in enumerate(working_days, 1):
        print(f"{i:2d}. {day}")
    
    print(f"\nВсего рабочих дней: {len(working_days)}")
    print("-" * 50)
    
    add_dates_and_grades_to_excel_files_in_folders(base_folder)
    print("\nГотово! Даты и оценки добавлены во все файлы по предметам.")
