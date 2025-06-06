# Формат выходного файла

## Файл target.xlsx

Выходной файл `target.xlsx` содержит обработанные данные о гарантиях в следующем формате:

| Наименование | SN OY | Уровень поддержки | Срок | Дата начала | Дата окончания |
|--------------|-------|-------------------|------|-------------|----------------|

## Описание столбцов

1. **Наименование**
   - Название оборудования
   - Берется из столбца "Номенклатура" входного файла

2. **SN OY**
   - Серийный номер оборудования
   - Форматируется как строка
   - Удаляются лишние символы (например, .0 в конце)

3. **Уровень поддержки**
   - Цветовая кодировка:
     - Базовый: светло-синий
     - Базовый+невозврат: синий
     - Гарантия: зелёный
     - Премиум: жёлтый
     - Премиум+невозврат: оранжевый
     - Расширенный: фиолетовый
     - Расширенный+невозврат: красный
     - Не найдено: серый

4. **Срок**
   - Количество лет гарантии
   - Цветовая кодировка:
     - 1 год: светло-зелёный
     - 3 года: светло-жёлтый
     - 5 лет: светло-оранжевый

5. **Дата начала**
   - Дата начала гарантии
   - Формат: YYYY-MM-DD

6. **Дата окончания**
   - Дата окончания гарантии
   - Формат: YYYY-MM-DD
   - Рассчитывается автоматически как "Дата начала" + "Срок" лет

## Особенности
- Дубликаты по SN OY автоматически удаляются (оставляется первая запись)
- Все даты форматируются в единый формат YYYY-MM-DD
- Применяется цветовое форматирование для удобства визуального анализа
- Данные из всех входных файлов объединяются в один файл 