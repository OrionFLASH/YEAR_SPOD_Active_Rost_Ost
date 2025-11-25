"""Основной модуль расчёта приростов по клиентам СПОД.

Вся бизнес-логика собрана в одном файле main.py согласно требованиям.
"""

from __future__ import annotations

import csv
import datetime as dt
import operator
import traceback
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional

import pandas as pd


SettingsTree = Dict[str, Any]
SELECTED_MANAGER_ID_COL = "Таб. номер ВКО (выбранный)"
SELECTED_MANAGER_NAME_COL = "ВКО (выбранный)"
DIRECT_MANAGER_ID_COL = "Таб. номер ВКО (по файлу)"
DIRECT_MANAGER_NAME_COL = "ВКО (по файлу)"


def build_settings_tree() -> SettingsTree:
    """Возвращает вложенную структуру настроек проекта."""

    # Структура intentionally verbose:
    # - "files" описывает, какие книги и листы участвуют, и даёт словарь alias↔source.
    # - "filters" держит правила отбраковки и переиспользуется в drop_forbidden_rows.
    # - "defaults"/"identifiers" управляют «заглушками» и форматированием ID.
    # - "spod"/"contest" отвечают за выгрузку и коды турнира.
    # - "variants" задаёт листы Excel (кластеры клиентских ключей).

    return {
        "files": {
            # sheet — лист по умолчанию. Если у конкретного файла другой лист, задайте его в items.
            "sheet": "Sheet1",
            # items: для каждого XLSX указываем ключ (не менять без причины), ярлык периода и имя файла в каталоге IN.
            #  - Хотите обрабатывать другие файлы? Меняйте file_name или добавляйте новые записи по аналогии.
            #  - Нужно читать несколько листов? Задайте "sheet": "Имя листа" у соответствующего item.
            "items": [
                {
                    "key": "current",          # фиксированный ключ T-0; лучше не переименовывать.
                    "label": "T-0",            # подпись, которая пойдёт в логи.
                    "file_name": "M-10_DIF_20251125_1100.xlsx",
                    "sheet": "Sheet1",
                },
                {
                    "key": "previous",         # фиксированный ключ T-1.
                    "label": "T-1",
                    "file_name": "M-9_DIF_20251125_1059.xlsx",
                    "sheet": "Sheet1",
                },
            ],
            # columns: связываем внутренние alias и исходные названия колонок.
            #  - Чтобы подставить другое поле из Excel, достаточно изменить "source".
            #  - Чтобы добавить ещё колонку, расширьте список и пропишите alias (английское имя) + source (русский заголовок).
            "columns": [
                {"alias": "tb", "source": "Короткое ТБ"},
                {"alias": "gosb", "source": "Полное ГОСБ"},
                {"alias": "manager_name", "source": "ФИО"},
                {"alias": "manager_id", "source": "Табельный номер"},
                {"alias": "client_id", "source": "ИНН"},
                {"alias": "fact_value", "source": "Факт"},
            ],
        },
        "filters": {
            # drop_rules: список словарей {alias, values}, где values — любые запрещённые строки.
            #  - Можно дополнять массив values новыми маркерами (например, ["-", "N/A", "Удалить"]).
            #  - alias должен совпадать с alias из блока columns.
            "drop_rules": [
                {"alias": "manager_name", "values": ["-", "Серая зона"]},
                {"alias": "manager_id", "values": ["-", "Green_Zone", "Tech_Sib"]},
                {"alias": "client_id", "values": ["Report_id не определен"]},
            ],
        },
        "defaults": {
            # Заглушки менеджера, которые попадут в итог, если T-0/T-1 не дали значений.
            #  - Можно указывать реальные ФИО/табельные номера из справочника (строки).
            "manager_name": "Не найден КМ",
            "manager_id": "90000009",
        },
        "identifiers": {
            # Форматирование ID: задаём символ заполнения (fill_char) и итоговую длину.
            #  - Например, замените на {"fill_char": " ", "total_length": 6}, если нужен пробел и длина 6.
            "manager_id": {"fill_char": "0", "total_length": 8},
            "client_id": {"fill_char": "0", "total_length": 12},
        },
        "spod": {
            # Параметры выгрузки СПОД.
            #  file_prefix — префикс имен Excel/CSV; допустимы любые латинские буквы+цифры.
            #  log_topic — метка в названиях логов (можно использовать разные темы для разных запусков).
            #  plan_value — число, которое пойдёт в PLAN_VALUE (float/decimal).
            #  priority — строковый приоритет записи ("1", "2" и т.д., как требует СПОД).
            #  fact_value_filter — условие отбора строк на лист SPOD (all, >=0, >1000, <=-10, ==0, !=0 ...).
            #  csv_variant — выбранная комбинация для CSV (номер от 1 до 8 согласно матрице вариантов):
            #    1: ВКО, без ТБ, КМ по каждому файлу
            #    2: ВКО, с ТБ, КМ по каждому файлу
            #    3: ВКО, без ТБ, последний КМ
            #    4: ВКО, с ТБ, последний КМ
            #    5: ИНН, без ТБ, КМ по каждому файлу
            #    6: ИНН, с ТБ, КМ по каждому файлу
            #    7: ИНН, без ТБ, последний КМ
            #    8: ИНН, с ТБ, последний КМ
            "file_prefix": "YEAR_SPOD_Active_Rost_ost",
            "log_topic": "spod",
            "plan_value": 0.0,
            "priority": "1",
            "fact_value_filter": ">0",
            "csv_variant": 1,
        },
        "contest": {
            # Официальные коды и дата турнира (строка в формате DD/MM/YYYY).
            #  - При изменении даты здесь автоматически обновится поле CONTEST_DATE в CSV/Excel.
            "contest_code": "01_2025-2_14-1_2",
            "tournament_code": "t_01_2025-2_14-1_2_1001",
            "contest_date": "31/10/2025",
        },
        "manager_views": [
            {
                "name": "TN_VKO",
                "source_variant": "ID_TN",
                "include_tb": False,
                "manager_mode": "latest",
                "use_for_spod": True,
            },
            {
                "name": "TN_VKO_TB",
                "source_variant": "ID_TB_TN",
                "include_tb": True,
                "manager_mode": "latest",
                "use_for_spod": False,
            },
        ],
        "direct_manager_views": [
            {"name": "MANAGER_DIRECT", "include_tb": False},
            {"name": "MANAGER_DIRECT_TB", "include_tb": True},
        ],
        "growth_combinations": [
            {
                "name": "COMBO_VKO_NO_TB",
                "type": "direct",
                "source": "MANAGER_DIRECT",
                "description": "Прирост по ВКО без учёта ТБ (сумма по каждому КМ в T0 минус сумма по нему же в T1).",
            },
        ],
        "variants": [
            # Наборы ключей для листов Excel.
            #  - name превращается в имя вкладки (ID, ID_TB и т.д.).
            #  - columns — список alias, по которым группируются данные.
            #  - Можно добавить собственный лист, указав аналогичный словарь {"name": "...", "columns": [...]}.
            {"name": "ID", "columns": ["client_id"]},
            {"name": "ID_TB", "columns": ["client_id", "tb"]},
            {"name": "ID_TN", "columns": ["client_id", "manager_id"]},
            {"name": "ID_TB_TN", "columns": ["client_id", "tb", "manager_id"]},
        ],
    }


def build_column_profiles(columns: List[Dict[str, str]]) -> Dict[str, Dict[str, str]]:
    """Формирует маппинги alias↔source для переименования колонок."""

    # rename_map: перевод оригинальных колонок Excel в машинные имена;
    # alias_to_source: обратное отображение для вывода человекочитаемых заголовков.
    rename_map = {column["source"]: column["alias"] for column in columns}
    alias_to_source = {column["alias"]: column["source"] for column in columns}
    return {"rename_map": rename_map, "alias_to_source": alias_to_source}


def build_drop_rules(rule_items: List[Dict[str, Any]]) -> Dict[str, Iterable[str]]:
    """Возвращает словарь запретных значений по колонкам."""

    return {rule["alias"]: tuple(rule["values"]) for rule in rule_items}


def get_file_meta(file_section: Dict[str, Any], key: str) -> Dict[str, Any]:
    """Ищет метаданные файла по ключу."""

    for item in file_section["items"]:
        if item["key"] == key:
            return item
    raise KeyError(f"Не найдена конфигурация файла '{key}'")


def resolve_sheet_name(file_section: Dict[str, Any], file_key: str) -> str:
    """Определяет имя листа для конкретного файла (или общее значение)."""

    meta = get_file_meta(file_section, file_key)
    return meta.get("sheet") or file_section.get("sheet", "Sheet1")


def parse_contest_date(contest_settings: Mapping[str, Any]) -> str:
    """Возвращает дату турнира в формате ISO."""

    parsed = dt.datetime.strptime(contest_settings["contest_date"], "%d/%m/%Y")
    return parsed.strftime("%Y-%m-%d")


def get_manager_columns(mode: str) -> Mapping[str, str]:
    """Возвращает имена колонок для выбранного режима назначения менеджера."""

    mapping = {
        "latest": {
            "id": "Таб. номер ВКО_Актуальный",
            "name": "ВКО_Актуальный",
        },
        "current_period": {
            "id": "Таб. номер ВКО_T0",
            "name": "ВКО_T0",
        },
        "previous_period": {
            "id": "Таб. номер ВКО_T1",
            "name": "ВКО_T1",
        },
    }
    if mode not in mapping:
        raise ValueError(
            "Недопустимое значение manager_mode. Используйте latest, current_period или previous_period."
        )
    return mapping[mode]


def ensure_directories(directories: Iterable[Path]) -> None:
    """Создаёт недостающие каталоги."""

    for directory in directories:
        directory.mkdir(parents=True, exist_ok=True)


def timestamp_suffix() -> str:
    """Формирует суффикс вида _YYYYMMDD_HH_MM."""

    return dt.datetime.now().strftime("_%Y%m%d_%H_%M")


def format_identifier(value: Any, total_length: int, fill_char: str) -> str:
    """Преобразует числовой идентификатор с лидирующими символами."""

    text = "" if value is None else str(value).strip()
    if not text:
        return text
    digits = "".join(ch for ch in text if ch.isdigit())
    if digits:
        return digits.rjust(total_length, fill_char)
    return text


def safe_to_float(value: Any) -> Optional[float]:
    """Безопасно приводит значение к float."""

    try:
        return float(str(value).replace(" ", "").replace(",", "."))
    except (TypeError, ValueError):
        return None


def build_filter_mask(series: pd.Series, condition: str) -> pd.Series:
    """Возвращает булев маск для фильтрации значений по условию."""

    normalized = condition.strip().lower().replace(" ", "")
    if not normalized or normalized in ("all", "все"):
        return pd.Series(True, index=series.index)

    # Поддерживаются операторы сравнения и записи вида ">0", "<=1000", "==0", "!=5".
    # Часть после оператора парсится как число (точка или запятая).
    available_ops = {
        ">=": operator.ge,
        "<=": operator.le,
        "==": operator.eq,
        "!=": operator.ne,
        ">": operator.gt,
        "<": operator.lt,
        "=": operator.eq,
    }

    for token in ("<=", ">=", "==", "!=", ">", "<", "="):
        if normalized.startswith(token):
            threshold_text = normalized[len(token) :]
            try:
                threshold = float(threshold_text.replace(",", "."))
            except ValueError as error:
                raise ValueError(
                    f"Не удалось распознать значение фильтра '{condition}'."
                ) from error
            comparator = available_ops[token]
            return comparator(series, threshold)

    raise ValueError(
        "Фильтр FACT_VALUE должен начинаться с одного из операторов "
        "(>=, <=, >, <, ==, !=, = ) или быть 'all/все'."
    )


def normalize_string(value: Any) -> str:
    """Возвращает очищенную строку без None."""

    if value is None:
        return ""
    return str(value).strip()


def build_logger(log_dir: Path, topic: str) -> Dict[str, Any]:
    """Инициализирует файловый логгер и возвращает набор функций."""

    ensure_directories([log_dir])
    suffix = timestamp_suffix()
    info_path = log_dir / f"INFO_{topic}{suffix}.log"
    debug_path = log_dir / f"DEBUG_{topic}{suffix}.log"

    # INFO всегда дублируется в консоль, DEBUG пишется только в файл (согласно ТЗ).
    def info(message: str) -> None:
        line = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} - [INFO] - {message}"
        print(line)
        with info_path.open("a", encoding="utf-8") as info_file:
            info_file.write(f"{line}\n")

    def debug(message: str, class_name: str, func_name: str) -> None:
        line = (
            f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} - [DEBUG] - "
            f"{message} [class: {class_name} | def: {func_name}]"
        )
        with debug_path.open("a", encoding="utf-8") as debug_file:
            debug_file.write(f"{line}\n")

    return {"info": info, "debug": debug}


def log_info(logger: Mapping[str, Any], message: str) -> None:
    """Упрощает вызов INFO-сообщений."""

    logger["info"](message)


def log_debug(
    logger: Mapping[str, Any], message: str, class_name: str, func_name: str
) -> None:
    """Упрощает вызов DEBUG-сообщений."""

    logger["debug"](message, class_name, func_name)


def resolve_excel_engine(logger: Mapping[str, Any]) -> str:
    """Подбирает доступный движок для записи Excel."""

    try:
        import xlsxwriter  # noqa: F401

        log_info(
            logger,
            "Обнаружен модуль xlsxwriter: Excel будет сохранён с расширенным форматированием.",
        )
        return "xlsxwriter"
    except ModuleNotFoundError:
        log_info(
            logger,
            "Модуль xlsxwriter недоступен. Переключаюсь на openpyxl без расширенного форматирования.",
        )
        return "openpyxl"


# -------------------------- Работа с исходными файлами ----------------------


def read_source_file(
    file_path: Path,
    sheet_name: str,
    column_maps: Mapping[str, str],
    drop_rules: Mapping[str, Iterable[str]],
    identifiers: Mapping[str, Mapping[str, Any]],
    logger: Mapping[str, Any],
) -> pd.DataFrame:
    """Загружает исходный Excel и подготавливает данные."""

    if not file_path.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")

    log_info(logger, f"Загружаю данные из файла {file_path.name}")
    # Читаем один лист Excel и сразу переименовываем колонки в единый формат.
    raw_df = pd.read_excel(file_path, sheet_name=sheet_name)
    renamed = raw_df.rename(columns=column_maps)

    required_columns = list(column_maps.values())
    missing = [col for col in required_columns if col not in renamed.columns]
    if missing:
        raise ValueError(
            f"Отсутствуют обязательные колонки {missing} в файле {file_path}"
        )

    prepared = renamed[required_columns].copy()

    # Строковые столбцы очищаем от пробелов и None.
    for column in ("tb", "gosb", "manager_name"):
        prepared[column] = prepared[column].apply(normalize_string)

    manager_identifier = identifiers["manager_id"]
    client_identifier = identifiers["client_id"]

    # Форматируем табельные номера и ИНН в заранее заданную длину.
    prepared["manager_id"] = prepared["manager_id"].apply(
        lambda value: format_identifier(
            value=value,
            total_length=manager_identifier["total_length"],
            fill_char=manager_identifier["fill_char"],
        )
    )
    prepared["client_id"] = prepared["client_id"].apply(
        lambda value: format_identifier(
            value=value,
            total_length=client_identifier["total_length"],
            fill_char=client_identifier["fill_char"],
        )
    )

    prepared["fact_value_clean"] = prepared["fact_value"].apply(safe_to_float)

    cleaned = drop_forbidden_rows(prepared, drop_rules, logger)
    log_debug(
        logger,
        f"После очистки в {file_path.name} осталось строк: {len(cleaned)}",
        class_name="DataLoader",
        func_name="read_source_file",
    )
    return cleaned


def drop_forbidden_rows(
    df: pd.DataFrame,
    drop_rules: Mapping[str, Iterable[str]],
    logger: Mapping[str, Any],
) -> pd.DataFrame:
    """Удаляет строки с запрещёнными значениями."""

    cleaned = df.copy()
    for column, values in drop_rules.items():
        forbidden = {value.lower() for value in values}

        def is_forbidden(value: Any) -> bool:
            if value is None:
                return False
            return str(value).strip().lower() in forbidden

        before = len(cleaned)
        cleaned = cleaned[~cleaned[column].apply(is_forbidden)]
        log_debug(
            logger,
            f"Колонка {column}: удалено {before - len(cleaned)} строк",
            class_name="Cleaner",
            func_name="drop_forbidden_rows",
        )
    return cleaned


# -------------------------- Агрегация данных --------------------------------


def aggregate_facts(
    df: pd.DataFrame,
    key_columns: List[str],
    suffix: str,
    logger: Mapping[str, Any],
    variant_name: str,
) -> pd.DataFrame:
    """Группирует данные по ключу и суммирует факт."""

    grouped = (
        df[key_columns + ["fact_value_clean"]]
        .fillna({"fact_value_clean": 0.0})
        .groupby(key_columns, dropna=False, as_index=False)
        .sum(numeric_only=True)
    )
    renamed = grouped.rename(columns={"fact_value_clean": f"Факт_{suffix}"})
    log_debug(
        logger,
        f"{variant_name}: агрегировано {len(renamed)} строк для суффикса {suffix}",
        class_name="Aggregator",
        func_name="aggregate_facts",
    )
    return renamed


def select_best_manager(
    df: pd.DataFrame,
    key_columns: List[str],
    logger: Mapping[str, Any],
    variant_name: str,
) -> pd.DataFrame:
    """Определяет менеджера с максимальным фактом для ключа."""

    additional_columns = [
        column for column in ("manager_name", "manager_id") if column not in key_columns
    ]
    grouping_columns = key_columns + additional_columns
    grouped = (
        df[grouping_columns + ["fact_value_clean"]]
        .fillna({"fact_value_clean": 0.0})
        .groupby(grouping_columns, dropna=False, as_index=False)
        .sum(numeric_only=True)
    )
    idx = grouped.groupby(key_columns, dropna=False)["fact_value_clean"].idxmax()
    best = grouped.loc[idx, key_columns + additional_columns].copy()
    result = best.copy()
    if "manager_name" in result.columns and "manager_name" not in key_columns:
        result = result.rename(columns={"manager_name": "ВКО"})
    if "manager_id" in key_columns and "manager_id" in result.columns:
        result["Таб. номер ВКО"] = result["manager_id"]
    elif "manager_id" in result.columns:
        result = result.rename(columns={"manager_id": "Таб. номер ВКО"})
    # На выходе каждая строка — конкретный ключ (например client_id + manager_id)
    # и менеджер, который показал максимальный факт.
    log_debug(
        logger,
        f"{variant_name}: выбраны менеджеры для {len(result)} ключей",
        class_name="Aggregator",
        func_name="select_best_manager",
    )
    return result


def build_latest_manager(
    current_best: pd.DataFrame,
    previous_best: pd.DataFrame,
    key_columns: List[str],
    defaults: Mapping[str, Any],
    identifiers: Mapping[str, Any],
    logger: Mapping[str, Any],
    variant_name: str,
) -> pd.DataFrame:
    """Комбинирует менеджеров, отдавая приоритет файлу T-0."""

    default_name = defaults["manager_name"]
    identifier_settings = identifiers["manager_id"]
    default_id = format_identifier(
        defaults["manager_id"],
        total_length=identifier_settings["total_length"],
        fill_char=identifier_settings["fill_char"],
    )

    curr = (
        current_best.set_index(key_columns)
        if not current_best.empty
        else pd.DataFrame(columns=key_columns + ["ВКО", "Таб. номер ВКО"]).set_index(key_columns)
    )
    prev = (
        previous_best.set_index(key_columns)
        if not previous_best.empty
        else pd.DataFrame(columns=key_columns + ["ВКО", "Таб. номер ВКО"]).set_index(key_columns)
    )

    combined = prev.join(
        curr,
        how="outer",
        lsuffix="_prev",
        rsuffix="_curr",
    )
    combined["ВКО_Актуальный"] = combined["ВКО_curr"].combine_first(combined["ВКО_prev"]).fillna(default_name)
    combined["Таб. номер ВКО_Актуальный"] = combined["Таб. номер ВКО_curr"].combine_first(combined["Таб. номер ВКО_prev"]).fillna(default_id)

    result = combined.reset_index()[key_columns + ["ВКО_Актуальный", "Таб. номер ВКО_Актуальный"]]
    log_debug(
        logger,
        f"{variant_name}: определены актуальные менеджеры для {len(result)} ключей",
        class_name="Aggregator",
        func_name="build_latest_manager",
    )
    return result


def assemble_variant_dataset(
    variant_name: str,
    key_columns: List[str],
    current_df: pd.DataFrame,
    previous_df: pd.DataFrame,
    defaults: Mapping[str, Any],
    identifiers: Mapping[str, Any],
    logger: Mapping[str, Any],
) -> pd.DataFrame:
    """Формирует таблицу для конкретного варианта ключа."""

    log_debug(
        logger,
        f"{variant_name}: старт построения набора данных",
        class_name="Aggregator",
        func_name="assemble_variant_dataset",
    )

    agg_current = aggregate_facts(current_df, key_columns, "T0", logger, variant_name)
    agg_previous = aggregate_facts(previous_df, key_columns, "T1", logger, variant_name)
    merged = (
        pd.merge(agg_current, agg_previous, on=key_columns, how="outer")
        .fillna({"Факт_T0": 0.0, "Факт_T1": 0.0})
    )
    merged["Прирост"] = merged["Факт_T0"] - merged["Факт_T1"]

    best_current = select_best_manager(
        current_df, key_columns, logger, variant_name
    ).rename(columns={"ВКО": "ВКО_T0", "Таб. номер ВКО": "Таб. номер ВКО_T0"})
    best_previous = select_best_manager(
        previous_df, key_columns, logger, variant_name
    ).rename(columns={"ВКО": "ВКО_T1", "Таб. номер ВКО": "Таб. номер ВКО_T1"})

    merged = merged.merge(best_current, on=key_columns, how="left")
    merged = merged.merge(best_previous, on=key_columns, how="left")

    latest = build_latest_manager(
        current_best=best_current.rename(columns={"ВКО_T0": "ВКО", "Таб. номер ВКО_T0": "Таб. номер ВКО"}),
        previous_best=best_previous.rename(columns={"ВКО_T1": "ВКО", "Таб. номер ВКО_T1": "Таб. номер ВКО"}),
        key_columns=key_columns,
        defaults=defaults,
        identifiers=identifiers,
        logger=logger,
        variant_name=variant_name,
    )
    merged = merged.merge(latest, on=key_columns, how="left")

    log_debug(
        logger,
        f"{variant_name}: итоговый набор содержит {len(merged)} строк",
        class_name="Aggregator",
        func_name="assemble_variant_dataset",
    )
    return merged


def build_manager_summary(
    variant_df: pd.DataFrame,
    include_tb: bool,
    logger: Mapping[str, Any],
    summary_name: str,
    manager_columns: Mapping[str, str],
) -> pd.DataFrame:
    """Создаёт свод по уникальным ТН+ВКО (+ТБ опционально)."""

    manager_id_col = manager_columns["id"]
    manager_name_col = manager_columns["name"]

    group_columns = [manager_id_col, manager_name_col]
    tb_column_present = include_tb and "tb" in variant_df.columns
    if tb_column_present:
        group_columns.append("tb")

    grouped = (
        variant_df.groupby(group_columns, dropna=False)[["Факт_T0", "Факт_T1", "Прирост"]]
        .sum()
        .reset_index()
    )
    rename_map = {
        manager_id_col: SELECTED_MANAGER_ID_COL,
        manager_name_col: SELECTED_MANAGER_NAME_COL,
    }
    if tb_column_present:
        rename_map["tb"] = "ТБ"
    grouped = grouped.rename(columns=rename_map)

    log_debug(
        logger,
        f"{summary_name}: агрегировано {len(grouped)} строк",
        class_name="Aggregator",
        func_name="build_manager_summary",
    )
    return grouped


def clamp_width(length: int) -> int:
    """Ограничивает ширину столбца в диапазоне 70-200."""

    return max(70, min(length, 200))


def format_excel_sheet(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    """Применяет форматирование листа Excel."""

    if getattr(writer, "engine", "") != "xlsxwriter":
        # openpyxl и другие движки не поддерживают используемое форматирование.
        return

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    header_format = workbook.add_format({"bold": True, "text_wrap": True})
    wrap_format = workbook.add_format({"text_wrap": True})
    number_format = workbook.add_format({"num_format": "#,##0.00", "text_wrap": True})

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(df), max(0, df.shape[1] - 1))

    for col_idx, column in enumerate(df.columns):
        max_len = max((len(str(value)) for value in [column] + df[column].tolist()), default=0) + 2
        width = clamp_width(max_len)
        fmt = number_format if column.startswith("Факт") or column == "Прирост" else wrap_format
        worksheet.set_column(col_idx, col_idx, width, fmt)
    worksheet.set_row(0, None, header_format)


def format_decimal_string(value: float, decimals: int = 5) -> str:
    """Форматирует число вида 0.00000."""

    numeric_value = 0.0 if value is None or pd.isna(value) else float(value)
    return f"{numeric_value:.{decimals}f}"


def build_spod_dataset(
    manager_summary: pd.DataFrame,
    spod_config: Mapping[str, Any],
    contest_config: Mapping[str, Any],
    identifiers: Mapping[str, Any],
    logger: Mapping[str, Any],
) -> pd.DataFrame:
    """Готовит данные для загрузки в СПОД."""

    # 1) Фильтруем приросты по условию (например, >=0), 2) сортируем по убыванию.
    mask = build_filter_mask(
        manager_summary["Прирост"],
        spod_config["fact_value_filter"],
    )
    filtered = manager_summary[mask].copy()
    filtered = filtered.sort_values(by="Прирост", ascending=False)

    log_debug(
        logger,
        f"SPOD: после фильтра {spod_config['fact_value_filter']} осталось {len(filtered)} строк",
        class_name="Exporter",
        func_name="build_spod_dataset",
    )

    dataset = filtered.rename(
        columns={
            SELECTED_MANAGER_ID_COL: "MANAGER_PERSON_NUMBER",
            "Прирост": "FACT_VALUE",
        }
    )["MANAGER_PERSON_NUMBER"].to_frame()

    manager_identifier = identifiers["manager_id"]
    dataset["MANAGER_PERSON_NUMBER"] = dataset["MANAGER_PERSON_NUMBER"].apply(
        lambda value: format_identifier(
            value=value,
            total_length=max(manager_identifier["total_length"], 20),
            fill_char=manager_identifier["fill_char"],
        )
    )
    dataset["CONTEST_CODE"] = contest_config["contest_code"]
    dataset["TOURNAMENT_CODE"] = contest_config["tournament_code"]
    dataset["CONTEST_DATE"] = parse_contest_date(contest_config)
    dataset["PLAN_VALUE"] = format_decimal_string(spod_config["plan_value"])
    dataset["FACT_VALUE"] = filtered["Прирост"].apply(format_decimal_string)
    dataset["priority_type"] = spod_config["priority"]

    log_debug(
        logger,
        f"SPOD: подготовлено {len(dataset)} строк для выгрузки",
        class_name="Exporter",
        func_name="build_spod_dataset",
    )

    return dataset[
        [
            "MANAGER_PERSON_NUMBER",
            "CONTEST_CODE",
            "TOURNAMENT_CODE",
            "CONTEST_DATE",
            "PLAN_VALUE",
            "FACT_VALUE",
            "priority_type",
        ]
    ]


def rename_output_columns(
    df: pd.DataFrame, alias_to_source: Mapping[str, str]
) -> pd.DataFrame:
    """Возвращает DataFrame с русскими заголовками для ключей."""

    renamed = df.copy()
    # Переименовываем лишь ключевые идентификаторы; остальные остаются в машинном виде.
    mapping = {
        alias: alias_to_source.get(alias, alias)
        for alias in ("client_id", "tb", "manager_id")
    }
    renamed = renamed.rename(columns=mapping)
    return renamed


def format_raw_sheet(
    df: pd.DataFrame,
    alias_to_source: Mapping[str, str],
) -> pd.DataFrame:
    """Возвращает DataFrame для исходного листа с читаемыми колонками и типами."""

    printable = df.copy()

    # Переименовываем только те столбцы, которые известны пользователю.
    rename_mapping = {
        alias: alias_to_source.get(alias, alias)
        for alias in df.columns
        if alias in alias_to_source
    }
    printable = printable.rename(columns=rename_mapping)

    # Числовой факт выводим отдельным столбцом с гарантированным float.
    if "fact_value_clean" in printable.columns:
        printable = printable.rename(
            columns={"fact_value_clean": "Факт (число)"}
        )
        printable["Факт (число)"] = printable["Факт (число)"].apply(
            lambda value: 0.0 if value is None or pd.isna(value) else float(value)
        )

    return printable


def build_direct_manager_summary(
    current_df: pd.DataFrame,
    previous_df: pd.DataFrame,
    include_tb: bool,
    logger: Mapping[str, Any],
    summary_name: str,
) -> pd.DataFrame:
    """Суммирует факты по менеджерам напрямую (без ключа клиента)."""

    base_columns = ["manager_id", "manager_name"]
    tb_column_present = include_tb and "tb" in current_df.columns and "tb" in previous_df.columns
    if tb_column_present:
        base_columns.append("tb")

    def aggregate(source_df: pd.DataFrame, suffix: str) -> pd.DataFrame:
        grouped = (
            source_df[base_columns + ["fact_value_clean"]]
            .fillna({"fact_value_clean": 0.0})
            .groupby(base_columns, dropna=False, as_index=False)
            .sum(numeric_only=True)
        )
        return grouped.rename(columns={"fact_value_clean": f"Факт_{suffix}"})

    agg_current = aggregate(current_df, "T0")
    agg_previous = aggregate(previous_df, "T1")

    merged = pd.merge(agg_current, agg_previous, on=base_columns, how="outer").fillna(
        {"Факт_T0": 0.0, "Факт_T1": 0.0}
    )
    merged["Прирост"] = merged["Факт_T0"] - merged["Факт_T1"]

    rename_map = {
        "manager_id": DIRECT_MANAGER_ID_COL,
        "manager_name": DIRECT_MANAGER_NAME_COL,
    }
    if tb_column_present:
        rename_map["tb"] = "ТБ"
    result = merged.rename(columns=rename_map)

    log_debug(
        logger,
        f"{summary_name}: агрегировано {len(result)} строк (прямой подсчёт по менеджеру)",
        class_name="Aggregator",
        func_name="build_direct_manager_summary",
    )
    return result


def build_variant_matrix(
    current_df: pd.DataFrame,
    previous_df: pd.DataFrame,
    defaults: Mapping[str, Any],
    identifiers: Mapping[str, Any],
    logger: Mapping[str, Any],
) -> Dict[int, pd.DataFrame]:
    """Строит матрицу всех 8 вариантов расчета приростов согласно методологии.
    
    Возвращает словарь {номер_варианта: DataFrame} для всех 8 комбинаций:
    1: ВКО, без ТБ, КМ по каждому файлу
    2: ВКО, с ТБ, КМ по каждому файлу
    3: ВКО, без ТБ, последний КМ
    4: ВКО, с ТБ, последний КМ
    5: ИНН, без ТБ, КМ по каждому файлу
    6: ИНН, с ТБ, КМ по каждому файлу
    7: ИНН, без ТБ, последний КМ
    8: ИНН, с ТБ, последний КМ
    """
    
    results: Dict[int, pd.DataFrame] = {}
    
    # Варианты 5-8: ИНН (client_id) - используем существующие функции
    log_info(logger, "Строю варианты 5-8: ИНН")
    
    # Вариант 5: ИНН, без ТБ, КМ по каждому файлу
    log_info(logger, "Строю вариант 5: ИНН, без ТБ, КМ по каждому файлу")
    variant_5_df = assemble_variant_dataset(
        variant_name="V5_ИНН_безТБ_КМ_пофайлу",
        key_columns=["client_id"],
        current_df=current_df,
        previous_df=previous_df,
        defaults=defaults,
        identifiers=identifiers,
        logger=logger,
    )
    # Агрегируем по КМ из T0 (доминантный в каждом файле)
    summary_5 = build_manager_summary(
        variant_df=variant_5_df,
        include_tb=False,
        logger=logger,
        summary_name="V5_SUMMARY",
        manager_columns={"id": "Таб. номер ВКО_T0", "name": "ВКО_T0"},
    )
    results[5] = summary_5
    
    # Вариант 6: ИНН, с ТБ, КМ по каждому файлу
    log_info(logger, "Строю вариант 6: ИНН, с ТБ, КМ по каждому файлу")
    variant_6_df = assemble_variant_dataset(
        variant_name="V6_ИНН_сТБ_КМ_пофайлу",
        key_columns=["client_id", "tb"],
        current_df=current_df,
        previous_df=previous_df,
        defaults=defaults,
        identifiers=identifiers,
        logger=logger,
    )
    summary_6 = build_manager_summary(
        variant_df=variant_6_df,
        include_tb=True,
        logger=logger,
        summary_name="V6_SUMMARY",
        manager_columns={"id": "Таб. номер ВКО_T0", "name": "ВКО_T0"},
    )
    results[6] = summary_6
    
    # Вариант 7: ИНН, без ТБ, последний КМ
    log_info(logger, "Строю вариант 7: ИНН, без ТБ, последний КМ")
    variant_7_df = assemble_variant_dataset(
        variant_name="V7_ИНН_безТБ_КМ_последний",
        key_columns=["client_id"],
        current_df=current_df,
        previous_df=previous_df,
        defaults=defaults,
        identifiers=identifiers,
        logger=logger,
    )
    summary_7 = build_manager_summary(
        variant_df=variant_7_df,
        include_tb=False,
        logger=logger,
        summary_name="V7_SUMMARY",
        manager_columns={"id": "Таб. номер ВКО_Актуальный", "name": "ВКО_Актуальный"},
    )
    results[7] = summary_7
    
    # Вариант 8: ИНН, с ТБ, последний КМ
    log_info(logger, "Строю вариант 8: ИНН, с ТБ, последний КМ")
    variant_8_df = assemble_variant_dataset(
        variant_name="V8_ИНН_сТБ_КМ_последний",
        key_columns=["client_id", "tb"],
        current_df=current_df,
        previous_df=previous_df,
        defaults=defaults,
        identifiers=identifiers,
        logger=logger,
    )
    summary_8 = build_manager_summary(
        variant_df=variant_8_df,
        include_tb=True,
        logger=logger,
        summary_name="V8_SUMMARY",
        manager_columns={"id": "Таб. номер ВКО_Актуальный", "name": "ВКО_Актуальный"},
    )
    results[8] = summary_8
    
    # Варианты 1-4: ВКО (gosb) - аналогичная логика, но с gosb как ключом
    log_info(logger, "Строю варианты 1-4: ВКО")
    
    # Вариант 1: ВКО, без ТБ, КМ по каждому файлу
    log_info(logger, "Строю вариант 1: ВКО, без ТБ, КМ по каждому файлу")
    variant_1_df = assemble_variant_dataset(
        variant_name="V1_ВКО_безТБ_КМ_пофайлу",
        key_columns=["gosb"],
        current_df=current_df,
        previous_df=previous_df,
        defaults=defaults,
        identifiers=identifiers,
        logger=logger,
    )
    summary_1 = build_manager_summary(
        variant_df=variant_1_df,
        include_tb=False,
        logger=logger,
        summary_name="V1_SUMMARY",
        manager_columns={"id": "Таб. номер ВКО_T0", "name": "ВКО_T0"},
    )
    results[1] = summary_1
    
    # Вариант 2: ВКО, с ТБ, КМ по каждому файлу
    log_info(logger, "Строю вариант 2: ВКО, с ТБ, КМ по каждому файлу")
    variant_2_df = assemble_variant_dataset(
        variant_name="V2_ВКО_сТБ_КМ_пофайлу",
        key_columns=["gosb", "tb"],
        current_df=current_df,
        previous_df=previous_df,
        defaults=defaults,
        identifiers=identifiers,
        logger=logger,
    )
    summary_2 = build_manager_summary(
        variant_df=variant_2_df,
        include_tb=True,
        logger=logger,
        summary_name="V2_SUMMARY",
        manager_columns={"id": "Таб. номер ВКО_T0", "name": "ВКО_T0"},
    )
    results[2] = summary_2
    
    # Вариант 3: ВКО, без ТБ, последний КМ
    log_info(logger, "Строю вариант 3: ВКО, без ТБ, последний КМ")
    variant_3_df = assemble_variant_dataset(
        variant_name="V3_ВКО_безТБ_КМ_последний",
        key_columns=["gosb"],
        current_df=current_df,
        previous_df=previous_df,
        defaults=defaults,
        identifiers=identifiers,
        logger=logger,
    )
    summary_3 = build_manager_summary(
        variant_df=variant_3_df,
        include_tb=False,
        logger=logger,
        summary_name="V3_SUMMARY",
        manager_columns={"id": "Таб. номер ВКО_Актуальный", "name": "ВКО_Актуальный"},
    )
    results[3] = summary_3
    
    # Вариант 4: ВКО, с ТБ, последний КМ
    log_info(logger, "Строю вариант 4: ВКО, с ТБ, последний КМ")
    variant_4_df = assemble_variant_dataset(
        variant_name="V4_ВКО_сТБ_КМ_последний",
        key_columns=["gosb", "tb"],
        current_df=current_df,
        previous_df=previous_df,
        defaults=defaults,
        identifiers=identifiers,
        logger=logger,
    )
    summary_4 = build_manager_summary(
        variant_df=variant_4_df,
        include_tb=True,
        logger=logger,
        summary_name="V4_SUMMARY",
        manager_columns={"id": "Таб. номер ВКО_Актуальный", "name": "ВКО_Актуальный"},
    )
    results[4] = summary_4
    
    return results


# ----------------------------- Основной сценарий ----------------------------


def process_project(project_root: Path) -> None:
    """Запускает полный цикл обработки данных."""

    # 1. Собираем настройки (файлы, фильтры, идентификаторы, листы Excel).
    # 2. Загружаем T-0/T-1, чистим их и агрегируем по набору ключей.
    # 3. Строим итоговые своды + выгрузку СПОД и сохраняем Excel/CSV.

    settings = build_settings_tree()
    file_section = settings["files"]
    column_profiles = build_column_profiles(file_section["columns"])
    rename_map = column_profiles["rename_map"]
    alias_to_source = column_profiles["alias_to_source"]
    drop_rules = build_drop_rules(settings["filters"]["drop_rules"])
    defaults = settings["defaults"]
    identifiers = settings["identifiers"]
    spod_config = settings["spod"]
    contest_config = settings["contest"]
    manager_views_config = settings.get("manager_views", [])
    direct_manager_views_config = settings.get("direct_manager_views", [])
    growth_combinations_config = settings.get("growth_combinations", [])
    variant_definitions = settings["variants"]

    # Готовим быстрый индекс по ключам файлов (current / previous).
    file_index = {item["key"]: item for item in file_section["items"]}
    current_meta = file_index["current"]
    previous_meta = file_index["previous"]
    sheet_current = resolve_sheet_name(file_section, "current")
    sheet_previous = resolve_sheet_name(file_section, "previous")

    input_dir = project_root / "IN"
    output_dir = project_root / "OUT"
    log_dir = project_root / "log"
    ensure_directories([input_dir, output_dir, log_dir])

    logger = build_logger(log_dir, spod_config["log_topic"])
    log_info(logger, "Старт обработки проекта YEAR_SPOD_Active_Rost_Ost")

    try:
        current_file = input_dir / current_meta["file_name"]
        previous_file = input_dir / previous_meta["file_name"]

        if not current_file.exists() or not previous_file.exists():
            log_info(
                logger,
                "Ожидаемые файлы отсутствуют в каталоге IN. "
                "Положите исходные XLSX и повторите запуск."
            )
            return

        current_df = read_source_file(
            current_file,
            sheet_current,
            rename_map,
            drop_rules,
            identifiers,
            logger,
        )
        previous_df = read_source_file(
            previous_file,
            sheet_previous,
            rename_map,
            drop_rules,
            identifiers,
            logger,
        )

        variant_tables: Dict[str, pd.DataFrame] = {}
        for variant in variant_definitions:
            name = variant["name"]
            columns = variant["columns"]
            # Для каждого набора ключей строим отдельный лист (ID / ID_TB / ...).
            log_info(logger, f"Формирую лист {name}")
            table = assemble_variant_dataset(
                variant_name=name,
                key_columns=columns,
                current_df=current_df,
                previous_df=previous_df,
                defaults=defaults,
                identifiers=identifiers,
                logger=logger,
            )
            variant_tables[name] = table

        manager_view_tables: Dict[str, pd.DataFrame] = {}
        spod_summary_df: Optional[pd.DataFrame] = None
        for view_cfg in manager_views_config:
            source_variant = view_cfg["source_variant"]
            variant_df = variant_tables[source_variant]
            manager_mode = view_cfg.get("manager_mode", "latest")
            manager_columns = get_manager_columns(manager_mode)
            include_tb = view_cfg.get("include_tb", False)
            summary_df = build_manager_summary(
                variant_df=variant_df,
                include_tb=include_tb,
                logger=logger,
                summary_name=view_cfg["name"],
                manager_columns=manager_columns,
            )
            manager_view_tables[view_cfg["name"]] = summary_df
            if view_cfg.get("use_for_spod"):
                spod_summary_df = summary_df

        if spod_summary_df is None:
            raise ValueError(
                "Не найден manager_view с use_for_spod=True. Укажите хотя бы один сценарий в настройках."
            )

        direct_manager_tables: Dict[str, pd.DataFrame] = {}
        for direct_cfg in direct_manager_views_config:
            summary_df = build_direct_manager_summary(
                current_df=current_df,
                previous_df=previous_df,
                include_tb=direct_cfg.get("include_tb", False),
                logger=logger,
                summary_name=direct_cfg["name"],
            )
            direct_manager_tables[direct_cfg["name"]] = summary_df

        growth_combination_tables: Dict[str, pd.DataFrame] = {}
        for combo_cfg in growth_combinations_config:
            source_type = combo_cfg.get("type")
            source_name = combo_cfg["source"]
            if source_type == "manager_view":
                source_table = manager_view_tables[source_name]
            elif source_type == "direct":
                source_table = direct_manager_tables[source_name]
            else:
                raise ValueError(
                    "Недопустимое значение combination.type. Используйте manager_view или direct."
                )
            growth_combination_tables[combo_cfg["name"]] = source_table.copy()

        # Строим матрицу всех 8 вариантов расчета
        log_info(logger, "Строю матрицу всех вариантов расчета (8 комбинаций)")
        variant_matrix = build_variant_matrix(
            current_df=current_df,
            previous_df=previous_df,
            defaults=defaults,
            identifiers=identifiers,
            logger=logger,
        )
        
        # Выбираем вариант для CSV из настроек
        csv_variant_num = spod_config.get("csv_variant", 1)
        if csv_variant_num not in variant_matrix:
            raise ValueError(f"Вариант {csv_variant_num} не найден в матрице. Доступны варианты 1-8.")
        
        csv_variant_df = variant_matrix[csv_variant_num]
        log_info(logger, f"Для CSV выбран вариант {csv_variant_num}")
        
        # Формируем SPOD датасет из выбранного варианта
        spod_dataset = build_spod_dataset(
            csv_variant_df, spod_config, contest_config, identifiers, logger
        )

        report_suffix = timestamp_suffix()
        excel_name = f"{spod_config['file_prefix']}{report_suffix}.xlsx"
        excel_path = output_dir / excel_name
        log_info(logger, f"Сохраняю Excel-файл {excel_name}")

        excel_engine = resolve_excel_engine(logger)

        with pd.ExcelWriter(excel_path, engine=excel_engine) as writer:
            for sheet_name, table in variant_tables.items():
                printable = rename_output_columns(table, alias_to_source)
                printable.to_excel(writer, sheet_name=sheet_name, index=False)
                format_excel_sheet(writer, sheet_name, printable)

            for sheet_name, summary_table in manager_view_tables.items():
                display_table = summary_table.rename(
                    columns={
                        SELECTED_MANAGER_ID_COL: "Таб. номер ВКО (выбранный)",
                        SELECTED_MANAGER_NAME_COL: "ВКО (выбранный)",
                    }
                )
                display_table.to_excel(writer, sheet_name=sheet_name, index=False)
                format_excel_sheet(writer, sheet_name, display_table)

            for sheet_name, summary_table in direct_manager_tables.items():
                display_table = summary_table.rename(
                    columns={
                        DIRECT_MANAGER_ID_COL: "Таб. номер ВКО (по файлу)",
                        DIRECT_MANAGER_NAME_COL: "ВКО (по файлу)",
                    }
                )
                display_table.to_excel(writer, sheet_name=sheet_name, index=False)
                format_excel_sheet(writer, sheet_name, display_table)

            for sheet_name, combo_table in growth_combination_tables.items():
                combo_table.to_excel(writer, sheet_name=sheet_name, index=False)
                format_excel_sheet(writer, sheet_name, combo_table)

            # Сохраняем все варианты матрицы в Excel
            variant_names = {
                1: "V1_ВКО_безТБ_КМ_пофайлу",
                2: "V2_ВКО_сТБ_КМ_пофайлу",
                3: "V3_ВКО_безТБ_КМ_последний",
                4: "V4_ВКО_сТБ_КМ_последний",
                5: "V5_ИНН_безТБ_КМ_пофайлу",
                6: "V6_ИНН_сТБ_КМ_пофайлу",
                7: "V7_ИНН_безТБ_КМ_последний",
                8: "V8_ИНН_сТБ_КМ_последний",
            }
            for variant_num, variant_df in variant_matrix.items():
                sheet_name = variant_names.get(variant_num, f"VARIANT_{variant_num}")
                variant_df.to_excel(writer, sheet_name=sheet_name, index=False)
                format_excel_sheet(writer, sheet_name, variant_df)

            spod_dataset.to_excel(writer, sheet_name="SPOD", index=False)
            format_excel_sheet(writer, "SPOD", spod_dataset)

            raw_sheets = [
                ("RAW_T0", format_raw_sheet(current_df, alias_to_source)),
                ("RAW_T1", format_raw_sheet(previous_df, alias_to_source)),
            ]
            for sheet_name, raw_table in raw_sheets:
                raw_table.to_excel(writer, sheet_name=sheet_name, index=False)
                format_excel_sheet(writer, sheet_name, raw_table)

        csv_name = f"{spod_config['file_prefix']}_SPOD{report_suffix}.csv"
        csv_path = output_dir / csv_name
        log_info(logger, f"Сохраняю CSV-файл {csv_name}")
        spod_dataset.to_csv(csv_path, index=False, sep=";", quoting=csv.QUOTE_MINIMAL)

        log_info(logger, "Обработка успешно завершена")
    except Exception as exc:
        log_info(logger, f"Обработка завершилась с ошибкой: {exc}")
        stack = traceback.format_exc().replace("\n", " | ")
        log_debug(
            logger,
            f"Трассировка: {stack}",
            class_name="Main",
            func_name="process_project",
        )
        raise


def main() -> None:
    """Точка входа."""

    project_root = Path(__file__).resolve().parent.parent
    process_project(project_root)


if __name__ == "__main__":
    main()
