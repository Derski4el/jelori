import os
import random
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
from typing import List, Dict, Tuple, Optional
import logging
from difflib import SequenceMatcher
import re

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class MonthlyAssessmentGenerator:
    """Класс для генерации месячной аттестации с оптимизациями"""
    
    # Константы
    SUBJECTS = [
        "Математика", "Русский язык", "Литература", "Иностранный язык",
        "Информатика", "История", "Обществознание", "Физика", "Химия",
        "Биология", "География", "Физическая культура",
        "Основы безопасности жизнедеятельности", "Основы профессиональной деятельности"
    ]
    
    # Стили (создаются один раз)
    HEADER_FONT = Font(bold=True, color="FFFFFF")
    HEADER_FILL = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
    
    SUBJECT_HEADER_FONT = Font(bold=True, color="FFFFFF")
    SUBJECT_HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    SUBJECT_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")

    def __init__(self, journals_path: str = "Журналы/1 Курс", result_folder: str = "Итог"):
        self.journals_path = journals_path
        self.result_folder = result_folder
        self.workbook_cache = {}  # Кэш для открытых файлов
        self.student_data_cache = {}  # Кэш данных студентов
        
        # Создаем папку результатов
        os.makedirs(result_folder, exist_ok=True)

    def get_working_days_for_month(self, year: int, month: int) -> List[str]:
        """Универсальная функция для получения рабочих дней любого месяца"""
        from datetime import datetime, timedelta
        working_days = []
        
        # Определяем последний день месяца
        if month == 12:
            next_month = datetime(year + 1, 1, 1)
        else:
            next_month = datetime(year, month + 1, 1)
        
        last_day = (next_month - timedelta(days=1)).day
        
        start_date = datetime(year, month, 1)
        end_date = datetime(year, month, last_day)
        
        current_date = start_date
        while current_date <= end_date:
            if current_date.weekday() < 5:  # 0-4 = пн-пт
                working_days.append(current_date.strftime("%d.%m.%Y"))
            current_date += timedelta(days=1)
        
        return working_days
    
    def get_working_days_september_2025(self) -> List[str]:
        """Возвращает список рабочих дней (пн-пт) сентября 2025 года"""
        return self.get_working_days_for_month(2025, 9)
    
    def get_working_days_october_2025(self) -> List[str]:
        """Возвращает список рабочих дней (пн-пт) октября 2025 года"""
        return self.get_working_days_for_month(2025, 10)
    
    def get_working_days_november_2025(self) -> List[str]:
        """Возвращает список рабочих дней (пн-пт) ноября 2025 года"""
        return self.get_working_days_for_month(2025, 11)
    
    def get_working_days_december_2025(self) -> List[str]:
        """Возвращает список рабочих дней (пн-пт) декабря 2025 года"""
        return self.get_working_days_for_month(2025, 12)
    
    def calculate_average_grade(self, grades: List[float]) -> float:
        """Вычисляет средний балл из списка оценок"""
        if not grades:
            return 0
        return round(sum(grades) / len(grades), 0)
    
    def get_groups(self) -> List[str]:
        """Получает список всех групп"""
        if not os.path.exists(self.journals_path):
            logger.error(f"Папка {self.journals_path} не найдена!")
            return []
        
        groups = []
        for item in os.listdir(self.journals_path):
            if os.path.isdir(os.path.join(self.journals_path, item)):
                groups.append(item)
        
        logger.info(f"Найдено групп: {len(groups)}")
        return groups
    
    def load_workbook_cached(self, file_path: str) -> Optional[Workbook]:
        """Загружает рабочую книгу с кэшированием"""
        if file_path in self.workbook_cache:
            return self.workbook_cache[file_path]
        
        try:
            if os.path.exists(file_path):
                wb = load_workbook(file_path)
                self.workbook_cache[file_path] = wb
                return wb
        except Exception as e:
            logger.error(f"Ошибка при загрузке файла {file_path}: {e}")
        
        return None
    
    def get_students_from_group(self, group_name: str) -> List[str]:
        """Получает список студентов группы с кэшированием"""
        cache_key = f"students_{group_name}"
        if cache_key in self.student_data_cache:
            return self.student_data_cache[cache_key]
        
        students_file = os.path.join(self.journals_path, group_name, "студенты.xlsx")
        students = []
        
        wb = self.load_workbook_cached(students_file)
        if wb:
            try:
                ws = wb.active
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row[1] and row[2] and row[3]:  # Фамилия, Имя, Отчество
                        fio = f"{row[1]} {row[2]} {row[3]}"
                        students.append(fio)
            except Exception as e:
                logger.error(f"Ошибка при чтении студентов группы {group_name}: {e}")
        
        self.student_data_cache[cache_key] = students
        return students
    
    def get_student_grades_from_subject(self, group_name: str, subject: str, student_fio: str, target_month: int = None) -> Tuple[List[float], int, int]:
        """Получает оценки, пропуски и количество занятий студента по предмету

        Args:
            group_name: название группы
            subject: название предмета
            student_fio: ФИО студента
            target_month: номер месяца для фильтрации (9-12). Если None, обрабатывает все данные.

        Возвращает кортеж: (оценки, пропуски_в_занятиях, всего_занятий)
        где "занятия" считаются как количество отметок (оценка или "Н").
        """
        subject_file = os.path.join(self.journals_path, group_name, f"{subject}.xlsx")
        grades = []
        absences = 0
        lessons_count = 0
        
        wb = self.load_workbook_cached(subject_file)
        if not wb:
            return grades, absences, lessons_count
        
        try:
            ws = wb.active
            
            # Находим строку студента (оптимизированный поиск)
            student_row = None
            for r in range(2, ws.max_row + 1):
                if ws.cell(row=r, column=1).value == student_fio:
                    student_row = r
                    break
            
            if student_row:
                # Определяем диапазон столбцов для обработки
                start_col = 3  # Столбец C (первые даты)
                
                if target_month:
                    # Получаем даты для целевого месяца
                    target_dates = self.get_working_days_for_month(2025, target_month)
                    # Находим столбцы с датами целевого месяца
                    month_columns = self._get_month_columns(ws, target_dates)
                    if not month_columns:
                        # Если столбцы с датами месяца не найдены, возвращаем пустые данные
                        return grades, absences, lessons_count
                else:
                    # Обрабатываем все столбцы с датами
                    month_columns = list(range(3, ws.max_column + 1))
                
                # Собираем оценки и пропуски только для целевого месяца
                for c in month_columns:
                    if c > ws.max_column:
                        continue
                        
                    grade = ws.cell(row=student_row, column=c).value
                    if grade is None or str(grade).strip() == "":
                        continue
                    # Любая отметка (оценка или "Н") считается занятием
                    lessons_count += 1
                    if isinstance(grade, (int, float)) and 2 <= grade <= 5:
                        grades.append(grade)
                    elif str(grade).strip() == "Н":
                        absences += 1
                        
        except Exception as e:
            logger.error(f"Ошибка при обработке предмета {subject} для студента {student_fio}: {e}")
        
        return grades, absences, lessons_count
    
    def _get_month_columns(self, ws, target_dates: List[str]) -> List[int]:
        """Находит номера столбцов с датами целевого месяца
        
        Args:
            ws: рабочая таблица Excel
            target_dates: список дат в формате DD.MM.YYYY
            
        Returns:
            List[int]: список номеров столбцов с датами целевого месяца
        """
        month_columns = []
        target_dates_set = set(target_dates)
        
        # Проверяем заголовки (строка 1) начиная с столбца C (3)
        for col in range(3, ws.max_column + 1):
            cell_value = ws.cell(row=1, column=col).value
            if cell_value and str(cell_value).strip() in target_dates_set:
                month_columns.append(col)
        
        return month_columns
    
    def _get_month_name(self, month: int) -> str:
        """Возвращает название месяца по номеру"""
        month_names = {9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь"}
        return month_names.get(month, "неизвестный")
    
    def apply_header_styles(self, ws, headers: List[str]):
        """Применяет стили к заголовкам"""
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            if col == 1:  # ФИО
                cell.font = self.HEADER_FONT
                cell.fill = self.HEADER_FILL
                cell.alignment = self.HEADER_ALIGNMENT
            else:  # Предметы
                cell.font = self.SUBJECT_HEADER_FONT
                cell.fill = self.SUBJECT_HEADER_FILL
                cell.alignment = self.SUBJECT_HEADER_ALIGNMENT
    
    def auto_adjust_column_width(self, ws):
        """Автоматически подгоняет ширину столбцов"""
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 30)
            ws.column_dimensions[column_letter].width = adjusted_width
    
    def process_group(self, wb: Workbook, group_name: str, target_month: int = None) -> int:
        """Обрабатывает одну группу и возвращает количество студентов
        
        Args:
            wb: рабочая книга Excel
            group_name: название группы
            target_month: номер месяца для фильтрации (9-12). Если None, обрабатывает все данные.
        """
        month_info = f" за {self._get_month_name(target_month)}" if target_month else ""
        logger.info(f"Обрабатываем группу: {group_name}{month_info}")
        
        # Создаем новый лист для группы
        ws = wb.create_sheet(title=group_name)
        
        # Заголовки
        headers = ["ФИО"] + self.SUBJECTS + ["Пропуски"]
        self.apply_header_styles(ws, headers)
        
        # Получаем студентов
        students = self.get_students_from_group(group_name)
        
        # Заполняем данные студентов
        # Агрегаторы метрик по группе
        failing_students_count = 0  # есть хотя бы одна "2" (в аттестации)
        students_with_one_2 = 0
        students_with_one_3 = 0
        students_with_one_4 = 0
        students_with_one_5 = 0
        students_4_and_5_only = 0  # только 4 и 5 в аттестации (нет 2 и 3)
        total_absences_lessons = 0  # всего пропусков (в занятиях)
        total_lessons = 0  # всего занятий (оценка или "Н")

        for row_idx, student_fio in enumerate(students, 2):
            ws.cell(row=row_idx, column=1, value=student_fio)
            
            total_absences = 0  # в занятиях
            lessons_for_student = 0
            subject_avgs: List[int] = []
            
            # Обрабатываем каждый предмет
            for col_idx, subject in enumerate(self.SUBJECTS, 2):
                grades, subject_absences, subject_lessons = self.get_student_grades_from_subject(
                    group_name, subject, student_fio, target_month
                )
                
                total_absences += subject_absences
                lessons_for_student += subject_lessons
                avg_grade = self.calculate_average_grade(grades)
                # Записываем средний балл по предмету (целое значение)
                ws.cell(row=row_idx, column=col_idx, value=int(avg_grade) if avg_grade else 0)
                # Копим аттестационные оценки для последующего расчёта метрик
                if avg_grade:
                    subject_avgs.append(int(avg_grade))
            
            # Записываем общее количество пропусков в часах
            absences_hours_col = 2 + len(self.SUBJECTS)
            ws.cell(row=row_idx, column=absences_hours_col, value=total_absences * 2)

            # Агрегируем по группе
            total_absences_lessons += total_absences
            total_lessons += lessons_for_student

            # Метрики по студенту
            if subject_avgs:
                count_2 = sum(1 for g in subject_avgs if g == 2)
                count_3 = sum(1 for g in subject_avgs if g == 3)
                count_4 = sum(1 for g in subject_avgs if g == 4)
                count_5 = sum(1 for g in subject_avgs if g == 5)

                if count_2 > 0:
                    failing_students_count += 1
                if count_2 == 1:
                    students_with_one_2 += 1
                if count_3 == 1:
                    students_with_one_3 += 1
                if count_4 == 1:
                    students_with_one_4 += 1
                if count_5 == 1:
                    students_with_one_5 += 1
                if count_2 == 0 and count_3 == 0 and (count_4 > 0 or count_5 > 0):
                    students_4_and_5_only += 1
        
        # После заполнения таблицы — вывод метрик под аттестацией
        current_last_row = 1 + len(students)
        metrics_start_row = current_last_row + 2

        # Расчёт итоговых метрик по группе
        students_count = len(students)
        avg_absences_per_student_hours = 0.0
        if students_count > 0:
            avg_absences_per_student_hours = (total_absences_lessons * 2) / students_count
        attendance_percent = 0.0
        if total_lessons > 0:
            attendance_percent = round(100.0 * (1 - (total_absences_lessons / total_lessons)), 1)
        success_percent = 0.0
        if students_count > 0:
            success_percent = round(100.0 * ((students_count - failing_students_count) / students_count), 1)

        metrics = [
            ("Неуспевающих, чел.", failing_students_count),
            ("Студентов с одной '2', чел.", students_with_one_2),
            ("Студентов с одной '3', чел.", students_with_one_3),
            ("Студентов с одной '4', чел.", students_with_one_4),
            ("Студентов с одной '5', чел.", students_with_one_5),
            ("Кол-во пропусков на 1 студента, часов", round(avg_absences_per_student_hours, 1)),
            ("Посещаемость, %", attendance_percent),
            ("Учатся на 4 и 5, чел.", students_4_and_5_only),
            ("Число студентов, чел.", students_count),
            ("Успеваемость, %", success_percent),
        ]

        for idx, (label, value) in enumerate(metrics, start=0):
            ws.cell(row=metrics_start_row + idx, column=1, value=label)
            ws.cell(row=metrics_start_row + idx, column=2, value=value)

        # Автоматически подгоняем ширину столбцов
        self.auto_adjust_column_width(ws)
        
        logger.info(f"Лист для группы {group_name} создан ({len(students)} студентов)")
        return len(students)
    
    def cleanup_cache(self):
        """Очищает кэш открытых файлов"""
        for wb in self.workbook_cache.values():
            try:
                wb.close()
            except:
                pass
        self.workbook_cache.clear()
        self.student_data_cache.clear()
    
    def calculate_similarity(self, text1: str, text2: str) -> float:
        """Вычисляет схожесть между двумя строками (0-1)"""
        return SequenceMatcher(None, text1.lower(), text2.lower()).ratio()
    
    def _normalize_name(self, name: str) -> str:
        """Нормализует строку ФИО: нижний регистр, один пробел, без лишних символов"""
        if not name:
            return ""
        # Заменяем небуквенно-цифровые разделители на пробел
        cleaned = re.sub(r"[\s\-._]+", " ", name.strip().lower())
        # Схлопываем множественные пробелы
        cleaned = re.sub(r"\s+", " ", cleaned)
        return cleaned

    def _tokenize_name(self, name: str) -> List[str]:
        """Разбивает ФИО на токены (фамилия, имя, отчество)"""
        normalized = self._normalize_name(name)
        return [t for t in normalized.split(" ") if t]

    def calculate_name_match_score(self, query: str, candidate_fio: str) -> float:
        """Считает оценку совпадения запроса с ФИО по токенам и целой строке (0-1).

        Логика:
        - Если любой токен запроса является префиксом любого токена ФИО → высокий балл.
        - Учитываем средний максимум похожести каждого токена запроса к токенам ФИО.
        - Также учитываем похожесть полной строки (как раньше); берем максимум из метрик.
        - Поддержка инициалов: одиночная буква = совпадение по первой букве токена ФИО.
        """
        if not query or not candidate_fio:
            return 0.0

        query_tokens = self._tokenize_name(query)
        fio_tokens = self._tokenize_name(candidate_fio)

        if not query_tokens or not fio_tokens:
            return 0.0

        # 1) Метрика по целой строке
        full_score = self.calculate_similarity(" ".join(query_tokens), " ".join(fio_tokens))

        # 2) Метрика по токенам: для каждого токена запроса берем лучший матч среди токенов ФИО
        per_token_scores: List[float] = []
        for q in query_tokens:
            best = 0.0
            # Инициал: одна буква
            if len(q) == 1:
                for t in fio_tokens:
                    if t and t[0] == q:
                        best = max(best, 0.85)
            for t in fio_tokens:
                if t.startswith(q) and len(q) >= 2:
                    # Префиксное совпадение — сильный сигнал
                    best = max(best, 0.92)
                # Fuzzy схожесть токенов
                best = max(best, self.calculate_similarity(q, t))
            per_token_scores.append(best)

        tokens_avg = sum(per_token_scores) / len(per_token_scores) if per_token_scores else 0.0

        # 3) Итоговый счет — максимум из токенного и полного
        return max(tokens_avg, full_score)
    
    def find_students_by_name(self, search_name: str, threshold: float = 0.6) -> List[Tuple[str, str, float]]:
        """
        Ищет студентов по ФИО с учетом неточности ввода
        Возвращает список кортежей: (группа, ФИО, коэффициент схожести)
        """
        matching_students = []
        groups = self.get_groups()
        
        # Адаптация порога для коротких запросов
        tokens = self._tokenize_name(search_name)
        min_token_len = min((len(t) for t in tokens), default=0)
        adaptive_threshold = threshold
        if len(tokens) == 1 and min_token_len <= 3:
            adaptive_threshold = max(0.5, threshold - 0.1)

        for group_name in groups:
            students = self.get_students_from_group(group_name)
            for student_fio in students:
                similarity = self.calculate_name_match_score(search_name, student_fio)
                if similarity >= adaptive_threshold:
                    matching_students.append((group_name, student_fio, similarity))
        
        # Сортируем по коэффициенту схожести (убывание)
        matching_students.sort(key=lambda x: x[2], reverse=True)
        return matching_students
    
    def get_all_student_grades(self, group_name: str, student_fio: str, target_month: int = None) -> Dict[str, Dict]:
        """
        Получает все оценки студента по всем предметам
        
        Args:
            group_name: название группы
            student_fio: ФИО студента
            target_month: номер месяца для фильтрации (9-12). Если None, обрабатывает все данные.
        
        Returns:
            Dict[str, Dict]: {предмет: {'grades': [оценки], 'absences': количество_пропусков, 'average': средний_балл}}
        """
        student_grades = {}
        
        for subject in self.SUBJECTS:
            grades, absences, lessons = self.get_student_grades_from_subject(group_name, subject, student_fio, target_month)
            average = self.calculate_average_grade(grades)
            
            student_grades[subject] = {
                'grades': grades,
                'absences': absences,
                'average': average
            }
        
        return student_grades
    
    def display_student_grades(self, group_name: str, student_fio: str, student_grades: Dict[str, Dict]):
        """Выводит на экран все оценки студента в удобном формате"""
        print(f"\n{'='*80}")
        print(f"ОЦЕНКИ СТУДЕНТА: {student_fio}")
        print(f"ГРУППА: {group_name}")
        print(f"{'='*80}")
        
        total_grades = 0
        total_absences = 0
        subjects_with_grades = 0
        
        for subject, data in student_grades.items():
            grades = data['grades']
            absences = data['absences']
            average = data['average']
            
            if grades or absences > 0:
                print(f"\n[ПРЕДМЕТ] {subject}:")
                if grades:
                    print(f"   Оценки: {', '.join(map(str, grades))}")
                    print(f"   Средний балл: {average}")
                    total_grades += len(grades)
                    subjects_with_grades += 1
                if absences > 0:
                    print(f"   Пропуски: {absences} (часов: {absences * 2})")
                    total_absences += absences
        
        print(f"\n{'='*80}")
        print(f"ОБЩАЯ СТАТИСТИКА:")
        print(f"   Всего оценок: {total_grades}")
        print(f"   Всего пропусков: {total_absences} (часов: {total_absences * 2})")
        print(f"   Предметов с оценками: {subjects_with_grades}")
        print(f"{'='*80}")
    
    def search_and_display_student(self, search_name: str):
        """Основная функция для поиска студента и вывода его оценок"""
        print(f"\n[ПОИСК] Студента: '{search_name}'")
        
        # Ищем студентов
        matching_students = self.find_students_by_name(search_name)
        
        if not matching_students:
            print("[ОШИБКА] Студенты не найдены. Попробуйте изменить поисковый запрос.")
            return
        
        print(f"\n[НАЙДЕНО] Студентов: {len(matching_students)}")
        
        # Показываем найденных студентов
        for i, (group_name, student_fio, similarity) in enumerate(matching_students, 1):
            print(f"{i}. {student_fio} (группа: {group_name}, схожесть: {similarity:.2f})")
        
        # Если найден только один студент, сразу показываем его оценки
        if len(matching_students) == 1:
            group_name, student_fio, _ = matching_students[0]
            student_grades = self.get_all_student_grades(group_name, student_fio)
            self.display_student_grades(group_name, student_fio, student_grades)
        else:
            # Если несколько студентов, просим выбрать
            try:
                choice = input(f"\nВыберите номер студента (1-{len(matching_students)}) или 0 для выхода: ")
                choice_num = int(choice)
                
                if choice_num == 0:
                    print("Поиск отменен.")
                    return
                elif 1 <= choice_num <= len(matching_students):
                    group_name, student_fio, _ = matching_students[choice_num - 1]
                    student_grades = self.get_all_student_grades(group_name, student_fio)
                    self.display_student_grades(group_name, student_fio, student_grades)
                else:
                    print("[ОШИБКА] Неверный номер.")
            except ValueError:
                print("[ОШИБКА] Введите корректный номер.", exc_info=True)

    def create_monthly_assessment(self, month: int = None) -> str:
        """Создает итоговую таблицу 'Месячная аттестация' с средними баллами по предметам
        
        Args:
            month: номер месяца (9-12 для сентября-декабря 2025). Если None, создается общая аттестация.
        """
        try:
            # Получаем список групп
            groups = self.get_groups()
            if not groups:
                logger.error("Группы не найдены!")
                return ""
            
            # Создаем новую рабочую книгу
            wb = Workbook()
            wb.remove(wb.active)  # Удаляем стандартный лист
            
            month_names = {9: "сентября", 10: "октября", 11: "ноября", 12: "декабря"}
            
            if month:
                logger.info(f"Создаем месячную аттестацию за {month_names[month]} 2025...")
            else:
                logger.info("Создаем общую месячную аттестацию...")
            
            total_students = 0
            for group_name in groups:
                students_count = self.process_group(wb, group_name, month)
                total_students += students_count
            
            # Сохраняем файл
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if month:
                filename = os.path.join(self.result_folder, f"Месячная аттестация_{month_names[month]}_2025_{timestamp}.xlsx")
            else:
                filename = os.path.join(self.result_folder, f"Месячная аттестация_{timestamp}.xlsx")
            wb.save(filename)
            
            logger.info(f"Файл сохранен: {filename}")
            logger.info(f"Создано листов: {len(groups)}")
            logger.info(f"Обработано студентов: {total_students}")
            logger.info("Готово!")
            
            return filename
            
        except Exception as e:
            logger.error(f"Критическая ошибка при создании аттестации: {e}")
            return ""
        finally:
            # Очищаем кэш
            self.cleanup_cache()

def create_monthly_assessment():
    """Основная функция для создания месячной аттестации"""
    generator = MonthlyAssessmentGenerator()
    return generator.create_monthly_assessment()

def search_student_grades():
    """Основная функция для поиска студента и вывода его оценок"""
    generator = MonthlyAssessmentGenerator()
    
    print("СИСТЕМА ПОИСКА СТУДЕНТОВ И ОЦЕНОК")
    print("="*50)
    
    try:
        while True:
            print("\nВыберите действие:")
            print("1. Найти студента и показать оценки")
            print("2. Создать месячную аттестацию (общую)")
            print("3. Создать месячную аттестацию за конкретный месяц")
            print("4. Создать аттестации за все месяцы (сентябрь-декабрь)")
            print("0. Выход")
            
            choice = input("\nВведите номер действия: ").strip()
            
            if choice == "1":
                search_name = input("\nВведите ФИО студента (можно не точно): ").strip()
                if search_name:
                    generator.search_and_display_student(search_name)
                else:
                    print("[ОШИБКА] Введите ФИО студента.")
            
            elif choice == "2":
                print("\n[СОЗДАНИЕ] Общей месячной аттестации...")
                result = generator.create_monthly_assessment()
                if result:
                    print(f"[УСПЕХ] Аттестация создана: {result}")
                else:
                    print("[ОШИБКА] При создании аттестации.")
            
            elif choice == "3":
                print("\nВыберите месяц для создания аттестации:")
                print("9. Сентябрь 2025")
                print("10. Октябрь 2025")
                print("11. Ноябрь 2025")
                print("12. Декабрь 2025")
                
                month_choice = input("Введите номер месяца (9-12): ").strip()
                try:
                    month = int(month_choice)
                    if month in [9, 10, 11, 12]:
                        month_names = {9: "сентября", 10: "октября", 11: "ноября", 12: "декабря"}
                        print(f"\n[СОЗДАНИЕ] Аттестации за {month_names[month]} 2025...")
                        result = generator.create_monthly_assessment(month)
                        if result:
                            print(f"[УСПЕХ] Аттестация создана: {result}")
                        else:
                            print("[ОШИБКА] При создании аттестации.")
                    else:
                        print("[ОШИБКА] Неверный номер месяца. Введите число от 9 до 12.")
                except ValueError:
                    print("[ОШИБКА] Введите корректный номер месяца.")
            
            elif choice == "4":
                print("\n[СОЗДАНИЕ] Аттестаций за все месяцы (сентябрь-декабрь 2025)...")
                months = [9, 10, 11, 12]
                month_names = {9: "сентября", 10: "октября", 11: "ноября", 12: "декабря"}
                
                success_count = 0
                for month in months:
                    print(f"Создаем аттестацию за {month_names[month]}...")
                    result = generator.create_monthly_assessment(month)
                    if result:
                        print(f"✓ Аттестация за {month_names[month]} создана")
                        success_count += 1
                    else:
                        print(f"✗ Ошибка при создании аттестации за {month_names[month]}")
                
                print(f"\n[РЕЗУЛЬТАТ] Создано аттестаций: {success_count} из {len(months)}")
            
            elif choice == "0":
                print("\nДо свидания!")
                break
            
            else:
                print("[ОШИБКА] Неверный выбор. Попробуйте снова.")
    
    except KeyboardInterrupt:
        print("\n\nПрограмма завершена пользователем.")
    except Exception as e:
        print(f"\n[ОШИБКА] Произошла ошибка: {e}")
    finally:
        generator.cleanup_cache()

if __name__ == "__main__":
    search_student_grades()
