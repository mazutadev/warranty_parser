# Формат входных файлов

## Требования к файлам
- Файлы должны быть в формате Excel (.xlsx или .xls)
- На каждом входном Excel-файле должен быть лист с именем **Гарантия**

## Структура данных
На листе "Гарантия" должны быть следующие столбцы (названия должны совпадать):

| Номенклатура | SN | Warranty | Начало гарантии | Окончание гарантии |
|--------------|----|----------|-----------------|--------------------|

**Пример данных:**

| Номенклатура                | SN              | Warranty             | Начало гарантии | Окончание гарантии |
|-----------------------------|-----------------|----------------------|-----------------|--------------------|
| Описание  | 24100100000059  | Premium5Y НеВозврат  | 2024-10-29      | 2029-10-28         |
| Описание  | 24120100000043  | Extended5Y           | 2024-12-16      | 2029-12-15         |

## Важные замечания
1. Столбцы могут идти в любом порядке, но должны присутствовать все перечисленные выше
2. Даты должны быть в формате YYYY-MM-DD или распознаваться Excel как даты
3. Серийные номера (SN) должны быть строкой или числом (скрипт сам приведёт к нужному виду)
4. **Внимание:** скрипт сам рассчитывает значение "Окончание гарантии" на основе "Начало гарантии" и срока из столбца Warranty. Значение из исходного файла будет проигнорировано

## Преобразование значений столбца Warranty

Возможные значения во входном файле и их преобразование в итоговый столбец "Уровень поддержки":

| Значение в Warranty (вход)      | Итоговое значение "Уровень поддержки" |
|---------------------------------|---------------------------------------|
| Base1Y, Base3Y, Base5Y          | Базовый                              |
| Base3Y НеВозврат                | Базовый+невозврат                     |
| Extended3Y, Extended5Y          | Расширенный                           |
| Extended3Y НеВозврат            | Расширенный+невозврат                  |
| Premium1Y, Premium3Y, Premium5Y | Премиум                               |
| Premium1Y НеВозврат, Premium3Y НеВозврат, Premium5Y НеВозврат | Премиум+невозврат |
| Гарантия1Y                      | Гарантия                              |
| NotFound                        | Не найдено                            |
| (любое другое значение)         | Не найдено                            |

### Дополнительные правила преобразования
- Если в значении присутствует "НеВозврат", оно будет преобразовано в вариант с "+невозврат"
- Регистр и пробелы не имеют значения (например, "premium5y невозврат" тоже будет распознано) 