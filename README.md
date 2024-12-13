## Проект: Анализ посещаемости и оценок студентов

### Описание
Данный проект предназначен для анализа данных о посещаемости и оценках студентов. 
На основе предоставленных данных программа рассчитывает:
- Среднюю оценку и процент посещаемости для курса.
- Статистику по каждому студенту.
- Топ студентов по успеваемости.
- Студентов с плохой посещаемостью.
- Количество студентов, получивших автомат, и тех, кто не получил.

Программа считывает данные из Excel-файла, анализирует их и сохраняет результаты в новый Excel-файл с подробной статистикой.

---

### Функциональные возможности
- Чтение данных из файла формата `.xlsx`.
- Проверка данных на валидность и их очистка.
- Вычисление средней оценки и количества пропусков.
- Определение, получил ли студент автомат.
- Генерация общей статистики по курсу.
- Формирование списка топ студентов и студентов с плохой посещаемостью.
- Сохранение результатов анализа в выходной Excel-файл.

---

### Установка и запуск

1. **Клонирование репозитория**
   ```bash
   git clone <URL данного репозитория>
   cd <название репозитория>
   ```

2. **Установка зависимостей**
   Выполните:
   ```bash
   pip install -r requirements.txt
   ```

3. **Запуск программы**
   Для запуска программы используется следующая команда:
   ```bash
   python academic_statistics.py <input_file> <output_file> --threshold <процент>
   ```
   - `<input_file>`: Путь к входному Excel-файлу с данными.
   - `<output_file>`: Путь к выходному файлу, где будут сохранены результаты.
   - `--threshold`: (опционально) Минимальный процент посещаемости для получения автомата. По умолчанию: `75`.

---

### Формат входного файла
Excel-файл должен иметь следующую структуру:

| ФИО          | День 1 | День 2 | День 3 | ... |
|--------------|--------|--------|--------|-----|
| Иванов И.И.  | +      | 3      | -      | ... |
| Сидоров С.С. | 4      | 5      | 2      | ... |

- `+` — посещение.
- `-` — отсутствие.
- Число — оценка.

---

### Выходной файл
Результаты анализа сохраняются в новом Excel-файле и включают следующие листы:
1. **Результаты**:
   - ФИО
   - Автомат (Да/Нет)
   - Средняя оценка
   - Количество пропусков
2. **Статистика**:
   - Средняя оценка по курсу.
   - Общий процент посещаемости.
3. **Топ-студенты**:
   - ФИО.
   - Средняя оценка.
4. **Проблемные студенты**:
   - Список студентов с количеством пропусков выше допустимого.

---

### Тестирование
Для запуска тестов выполните:
```bash
pytest
```
Пояснение тестов:
* Тесты для чтения XLSX: Тестируется функция read_xlsx с моком openpyxl.load_workbook, чтобы избежать необходимости работы с реальными файлами. Мы проверяем, что данные правильно считываются.
* Тесты для очистки посещаемости: В validate_attendance проверяем правильность очистки данных о посещаемости.
* Тесты для вычисления средней оценки: В get_avg_grade проверяется, как функция работает с различными входными данными, включая ошибки (например, текстовые значения).
* Тесты для проверки автомата: Тестируем логику проверки, был ли у студента автомат, на разных данных посещаемости.
* Тесты для подсчета пропусков: Проверяется, правильно ли функция подсчитывает количество пропусков.
* Тесты для расчета статистики по курсу: Проверяется вычисление средней оценки по всем студентам и процент посещаемости.
* Тесты для определения топ-студентов: Тестируются функции сортировки студентов по средней оценке.
* Тесты для определения проблемных студентов: Проверяется корректность определения студентов с большим количеством пропусков.
* Тесты для подсчета студентов с автоматом: Проверяется подсчет студентов, которые получили автомат, и тех, кто не прошел.

---

### Пример использования
```bash
python analyze_students.py students.xlsx results.xlsx --threshold 80
```
- Анализирует файл `students.xlsx` и сохраняет результаты в `results.xlsx`.
- Минимальный процент посещаемости для получения автомата: `80%`.

---

### Логирование
Логи программы сохраняются в консоли с уровнем `INFO`, включая:
- Успешное чтение данных.
- Количество обработанных студентов.
- Успешное сохранение результатов.

---

### Требования
- Python 3.8+
- Зависимости:
  - `openpyxl`
  - `pytest`

---