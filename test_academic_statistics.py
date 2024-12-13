import pytest
from unittest.mock import patch
from academic_statistics import (
    read_xlsx,
    validate_attendance,
    get_avg_grade,
    check_auto_pass,
    count_absences,
    calculate_statistics,
    get_top_students,
    get_bad_students,
    count_students_with_auto_pass
)

# Тестирование функции read_xlsx (Проверка моком на чтение данных с файла)
def test_read_xlsx():
    with patch("openpyxl.load_workbook") as mock_load_workbook:
        mock_workbook = mock_load_workbook.return_value
        mock_sheet = mock_workbook.active
        mock_sheet.iter_rows.return_value = [
            ("Иванов И.И.", "+", "3", "-"),
            ("Сидоров С.С.", "3", "4", "5"),
        ]
        students = read_xlsx("test.xlsx")
        assert len(students) == 2
        assert students[0]['name'] == "Иванов И.И."
        assert students[1]['attendance'] == ["3", "4", "5"]

# Тестирование функции validate_attendance (Проверка валидации значений)
def test_validate_attendance():
    cleaned = validate_attendance(["+", "1", "-"])
    assert cleaned == ["+", "1", "-"]
    cleaned = validate_attendance([1, 2, 3])
    assert cleaned == ["1", "2", "3"]

# Тестирование функции get_avg_grade (Получение средней оценки за курс)
def test_get_avg_grade():
    assert get_avg_grade(['2', '3', '4']) == 3
    assert get_avg_grade(['+', '4', '3']) == 4
    assert get_avg_grade(['+', '-', '3']) == 3
    assert get_avg_grade(['+', 'a', '5']) == 5
    assert get_avg_grade(['+']) == 0
    assert get_avg_grade(['-']) == 0

# Тестирование функции check_auto_pass (Проверка на получения автомата от посещений)
def test_check_auto_pass():
    assert check_auto_pass(['+', '3', '4'], 50) == 'Да'
    assert check_auto_pass(['+', '-', '5'], 100) == 'Нет'
    assert check_auto_pass(['+', '3', '-'], 25) == 'Да'
    assert check_auto_pass(['-'], 75) == 'Нет'

# Тестирование функции count_absences (Получаем кол-во пропусков студента)
def test_count_absences():
    assert count_absences(['+', '4', '-']) == 1
    assert count_absences(['+', '-', '-', '+']) == 2
    assert count_absences(['3', '4', '4']) == 0
    assert count_absences(['-', '-', '-']) == 3

# Тестирование функции calculate_statistics (Получаем статистику по курсу (средняя оценка и посещаемость))
def test_calculate_statistics():
    students = [
        {'name': 'Student 1', 'attendance': ['+', '4', '5', '-']},
        {'name': 'Student 2', 'attendance': ['+', '3', '2', '4']}
    ]
    stats = calculate_statistics(students)
    assert stats['avg_course_grade'] == 4 # 4 - средняя оценка по курсу
    assert stats['attendance_percentage'] == 87.5  # 87.5 - процент посещений по курсу

# Тест для get_bad_students (Получаем студентов с хорошей посещаемостью)
def test_get_top_students():
    students = [
        {'name': 'Иванов И.И.', 'attendance': ['+', '4', '3', '-']},
        {'name': 'Сидоров С.С.', 'attendance': ['+', '4', '5', '5']}
    ]
    top_students = get_top_students(students, 1)
    assert top_students == [('Сидоров С.С.', 5)]  # "Сидоров С.С." имеет больше посещений

# Тест для test_get_bad_students (Получаем студентов, у кого пропусков больше допустимного)
def test_get_bad_students():
    students = [
        {'name': 'Иванов И.И.', 'attendance': ['+', '3', '-', '-']},
        {'name': 'Сидоров С.С.', 'attendance': ['+', '4', '4', '2']}
    ]
    bad_students = get_bad_students(students, 1)
    assert bad_students == ['Иванов И.И.']  # "Иванов И.И." имеет 1+ пропуск, что превышает допустимое количество


# Тестирование функции count_students_with_auto_pass (получение автомата)
def test_count_students_with_auto_pass():
    students = [
        {'name': 'Student 1', 'attendance': ['+', '4', '4', '-']},
        {'name': 'Student 2', 'attendance': ['+', '5', '2', '3']}
    ]
    passed, failed = count_students_with_auto_pass(students, 75)
    assert passed == 2  # Оба студента получат автомат
    assert failed == 0

