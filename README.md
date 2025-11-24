# YEAR_SPOD_Active_Rost_Ost

## 1. Формулировка задачи и ТЗ
- Необходимо сравнить два XLSX-файла с остатками задолженности (T-0 и T-1), рассчитать приросты по клиентам и корректно сопоставить их с табельными номерами клиентских менеджеров.
- Полное техническое задание, полученное от заказчика, находится в `Docs/Задача.txt`. Файл перенесён без изменений и является первоисточником требований.

## 2. Описание решения
- Вся бизнес-логика реализована в одном файле `src/main.py` (как требовалось в задаче).
- Скрипт использует только стандартную библиотеку Python и пакет `pandas`, входящий в базовую поставку Anaconda.
- Входные файлы читаются из каталога `IN`, очищаются от запрещённых значений, нормализуются и агрегируются по четырём вариантам ключей (`ID`, `ID+ТБ`, `ID+ТН`, `ID+ТБ+ТН`).
- Для каждого варианта рассчитываются суммы фактов T-0 и T-1, приросты, менеджеры по данным обоих периодов и «актуальный» менеджер (приоритет у T-0, затем T-1, иначе заглушка).
- Дополнительно формируются листы с уникальными сочетаниями `ТН+ВКО` и `ТН+ВКО+ТБ`, а также выгрузка СПОД (лист Excel и CSV) с форматированием и заданными кодами конкурса.
- Логи ведутся раздельно для уровней INFO/DEBUG в каталоге `log`, в консоль выводится только INFO.

## 3. Структура каталогов
```
YEAR_SPOD_Active_Rost_Ost/
├── Docs/               # Дополнительные материалы (исходное ТЗ)
├── IN/                 # Входные XLSX-файлы (T-0 и T-1)
├── OUT/                # Итоговые Excel и CSV файлы
├── log/                # Логи INFO/DEBUG
├── src/
│   ├── main.py         # Основной скрипт
│   └── Tests/README.md # Состояние автотестов
├── .env.example        # Шаблон для будущих расширений (значения теперь в коде)
├── README.md           # Текущая документация
└── Docs/Задача.txt     # Исходное описание задачи
```

## 4. Настройка окружения
1. Создать и активировать виртуальное окружение:
   ```bash
   cd ~/Desktop/MyProject/YEAR_SPOD_Active_Rost_Ost
   python3 -m venv .venv
   source .venv/bin/activate
   ```
2. Установки через `pip` не требуются (используется базовый набор Anaconda / стандартная библиотека).
3. (Опционально) Скопировать `.env.example` в `.env`, если планируется возвращение к внешней конфигурации. По умолчанию все параметры уже прошиты в `src/main.py` и сведены в дерево настроек.

## 5. Конфигурация
Все значения собраны в иерархической структуре (списки секций с вложенными словарями), которую формирует функция `build_settings_tree()` в `src/main.py`. Доступные блоки:

1. **spod** — имена файлов, тема логов, план и приоритет выгрузки.
2. **contest** — коды конкурса/турнира и дата отчёта.
3. **defaults** — заглушки для ФИО и табельного номера КМ.
4. **identifiers** — правила нормализации табельных номеров и ИНН.

Чтобы изменить параметры, достаточно поправить соответствующие значения в словаре `values` нужной секции; `process_project` автоматически применит обновления при построении `OutputConfig` и `ProcessingConfig`.

## 6. Использование
1. Поместите исходные файлы в каталог `IN` под именами:
   - `АКТИВЫ 31-10-2025 (ОСТАТОК-V2).xlsx` (T-0)
   - `АКТИВЫ 31-12-2024 (ОСТАТОК-V2).xlsx` (T-1)
2. Запустите скрипт:
   ```bash
   python src/main.py
   ```
3. Результаты появятся в каталоге `OUT` в виде Excel и CSV файлов с суффиксом `_YYYYMMDD_HH_MM`.
4. Логи формирования находятся в `log/INFO_*` и `log/DEBUG_*`.

## 7. Логирование
- INFO: ключевые этапы обработки, пишутся в файл `INFO_<topic>_<timestamp>.log` и дублируются в консоль.
- DEBUG: подробные сообщения для каждой функции с указанием класса и имени функции, записываются в `DEBUG_<topic>_<timestamp>.log`.
- Формат DEBUG строки: `YYYY-MM-DD HH:MM:SS - [DEBUG] - Сообщение [class: <...> | def: <...>]`.

## 8. Список функций и примеры использования
| Функция / структура | Назначение | Пример вызова |
|--------------------|-----------|---------------|
| `ColumnConfig` | Хранит оригинальные имена колонок и маппинг | `column_config = ColumnConfig()` |
| `ProcessingConfig` | Параметры очистки и форматирования | `processing = ProcessingConfig()` |
| `OutputConfig` | Настройки выходных файлов и СПОД | `output_config = OutputConfig(... )` |
| `build_settings_tree()` | Возвращает дерево секций настроек | `settings = build_settings_tree()` |
| `get_settings_section(tree, name)` | Извлекает словарь выбранной секции | `contest = get_settings_section(settings, 'contest')` |
| `ensure_directories(paths)` | Создаёт недостающие каталоги | `ensure_directories([Path('IN')])` |
| `timestamp_suffix()` | Возвращает строку `_YYYYMMDD_HH_MM` | `suffix = timestamp_suffix()` |
| `format_identifier(value, length, char)` | Форматирует идентификаторы с лидирующими символами | `format_identifier('85461', 8, '0') -> '00085461'` |
| `safe_to_float(value)` | Безопасно приводит строку к `float` | `safe_to_float('43,51') -> 43.51` |
| `normalize_string(value)` | Очищает текстовое поле | `normalize_string('  ABC ') -> 'ABC'` |
| `StructuredLogger` | Класс для INFO/DEBUG логов | `logger = build_logger(Path('log'), 'spod')` |
| `build_logger(log_dir, topic)` | Создаёт экземпляр `StructuredLogger` | `logger = build_logger(Path('log'), 'spod')` |
| `read_source_file(path, columns, processing, logger)` | Загружает Excel, нормализует данные | `df = read_source_file(file_path, column_config, processing, logger)` |
| `drop_forbidden_rows(df, rules, logger)` | Удаляет строки с запрещёнными значениями | `cleaned = drop_forbidden_rows(df, processing.drop_rules, logger)` |
| `aggregate_facts(df, keys, suffix, logger, variant)` | Суммирует факт по ключу | `agg = aggregate_facts(df, ['client_id'], 'T0', logger, 'ID')` |
| `select_best_manager(df, keys, logger, variant)` | Определяет менеджера с максимальным фактом | `best = select_best_manager(df, ['client_id'], logger, 'ID')` |
| `build_latest_manager(curr, prev, keys, defaults, logger, variant)` | Комбинирует актуального менеджера | `latest = build_latest_manager(curr, prev, ['client_id'], 'Не найден', '90000009', logger, 'ID')` |
| `assemble_variant_dataset(variant, keys, df_t0, df_t1, processing, logger)` | Строит итоговый набор данных для листа | `variant_tables['ID'] = assemble_variant_dataset('ID', ['client_id'], df_t0, df_t1, processing, logger)` |
| `build_manager_summary(df, include_tb, logger, name)` | Готовит свод по ТН/ВКО | `summary = build_manager_summary(variant_tables['ID_TN'], False, logger, 'TN_VKO')` |
| `format_excel_sheet(writer, sheet, df)` | Применяет оформление листа | `format_excel_sheet(writer, 'ID', df)` |
| `format_decimal_string(value)` | Приводит число к строке вида `0.00000` | `format_decimal_string(12.3) -> '12.30000'` |
| `build_spod_dataset(summary, output_cfg, processing, logger)` | Создаёт таблицу для СПОД и CSV | `spod = build_spod_dataset(summary, output_cfg, processing, logger)` |
| `rename_output_columns(df, column_config)` | Возвращает русские подписи колонок | `rename_output_columns(table, column_config)` |
| `process_project(project_root)` | Композиция всех шагов пайплайна | `process_project(Path.cwd())` |
| `main()` | Точка входа CLI | `python src/main.py` |

## 9. История версий
| Версия | Дата | Изменения |
|--------|------|-----------|
| 1.1.0 | 2025-11-24 | Значения из `.env` перенесены напрямую в `src/main.py`, убрано чтение внешнего файла конфигурации. |
| 1.0.0 | 2025-11-24 | Создан репозиторий, реализован основной сценарий расчёта приростов, добавлены логирование, шаблон `.env`, структура каталогов и документация. |

## 10. Дополнительные материалы
- `Docs/Задача.txt` — исходное ТЗ.
- `src/Tests/README.md` — инструкция по тестированию до появления автотестов.

