"""Основной модуль расчёта приростов по клиентам СПОД.

Вся бизнес-логика собрана в одном файле main.py согласно требованиям.
"""

from __future__ import annotations

import csv
import datetime as dt
import operator
import traceback
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Set, Tuple

import pandas as pd
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


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
            # Глобальные параметры имени файлов/логов. Остальные настройки задаются
            # в разделе spod_variants индивидуально для каждого варианта.
            "file_prefix": "YEAR_SPOD_Active_Rost_ost",
            "log_topic": "spod",
        },
        "scenario": {
            # name используется в логах/именах листов.
            "name": "SCENARIO_MAIN",
            # key_mode: client (агрегация по ИНН) или manager (агрегация по табельному номеру).
            "key_mode": "client",
            # include_tb: True → добавляем ТБ в ключ и считаем показатели в разрезе ТерБанков.
            "include_tb": False,
            # manager_assignment: latest (последний доступный КМ) или per_file (T0 и T1 отдельно).
            "manager_assignment": "latest",
            # Имена листов в Excel (можно переопределить).
            "detail_sheet_name": "DETAIL_SCENARIO",
            "summary_sheet_name": "SUMMARY_TN",
            "percentile_sheet_name": "PERCENTILE_TN",
        },
        "spod_variants": [
            {
                "name": "SPOD_SCENARIO",
                "source_type": "scenario_summary",
                "calc_sheet_name": "CALC_SCENARIO",
                "spod_sheet_name": "SPOD_SCENARIO",
                "value_column": "Прирост",
                "fact_value_filter": ">0",
                "plan_value": 0.0,
                "priority": "1",
                "contest_code": "01_2025-2_14-1_2",
                "tournament_code": "t_01_2025-2_14-1_2_1001",
                "contest_date": "31/10/2025",
                "include_in_csv": True,
            },
            {
                "name": "SPOD_SCENARIO_PERCENTILE",
                "source_type": "scenario_percentile",
                "calc_sheet_name": "CALC_SCENARIO_PERC",
                "spod_sheet_name": "SPOD_SCENARIO_PERC",
                "value_column": "Обогнал_всего_%",
                "fact_value_filter": ">=0",
                "plan_value": 0.0,
                "priority": "1",
                "contest_code": "01_2025-2_14-1_2",
                "tournament_code": "t_01_2025-2_14-1_2_1001",
                "contest_date": "31/10/2025",
                "include_in_csv": True,
            },
        ],
        "report_layout": {
            # Управляет тем, какие листы попадают в основной Excel (пустой список = блок отключён).
            # Отсутствие ключа означает сохранение предыдущего поведения и выгрузку всех листов блока.
            # detail/summary/percentile — новые категории отчётов сценария.
            "detail_sheets": ["DETAIL_SCENARIO"],
            "summary_sheets": ["SUMMARY_TN"],
            "percentile_sheets": ["PERCENTILE_TN"],
            "calc_sheets": ["CALC_SCENARIO", "CALC_SCENARIO_PERC"],
            "spod_variants": ["SPOD_SCENARIO", "SPOD_SCENARIO_PERCENTILE"],
            # raw_sheets — очищенные исходники T-0/T-1.
            "raw_sheets": ["RAW_T0", "RAW_T1"],
        },
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


def parse_contest_date(contest_date: str) -> str:
    """Возвращает дату турнира в формате ISO."""

    parsed = dt.datetime.strptime(contest_date, "%d/%m/%Y")
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


def _compute_percentile_pair(series: pd.Series) -> Tuple[pd.Series, pd.Series]:
    """Вспомогательная функция: возвращает (обогнал_%, обогнали_%) для серии."""

    if series.empty:
        empty = pd.Series(0.0, index=series.index)
        return empty, empty

    rank_min = series.rank(method="min", ascending=True)
    rank_max = series.rank(method="max", ascending=True)
    count_equal = rank_max - rank_min + 1
    count_less = rank_min - 1
    count_greater = len(series) - rank_max

    obognal = ((count_less + 0.5 * (count_equal - 1)) / len(series)) * 100
    obognali = ((count_greater + 0.5 * (count_equal - 1)) / len(series)) * 100

    return obognal, obognali


def append_percentile_columns(
    table: pd.DataFrame,
    *,
    value_column: str,
    tb_column: Optional[str] = None,
) -> pd.DataFrame:
    """Добавляет в таблицу колонки процентных рангов (см. Docs/percentile_logic.md)."""

    if value_column not in table.columns:
        raise KeyError(
            f"Колонка '{value_column}' не найдена в таблице для расчёта процентилей."
        )

    prepared = table.copy()
    values = pd.to_numeric(prepared[value_column], errors="coerce").fillna(0.0)

    obognal_all, obognali_all = _compute_percentile_pair(values)
    prepared["Обогнал_всего_%"] = obognal_all
    prepared["Обогнали_меня_всего_%"] = obognali_all

    mask_non_negative = values >= 0
    if mask_non_negative.any():
        obognal_ge0, obognali_ge0 = _compute_percentile_pair(values[mask_non_negative])
        prepared["Обогнал_всего_≥0_%"] = obognal_ge0.reindex(
            prepared.index, fill_value=0.0
        )
        prepared["Обогнали_меня_всего_≥0_%"] = obognali_ge0.reindex(
            prepared.index, fill_value=0.0
        )
    else:
        prepared["Обогнал_всего_≥0_%"] = 0.0
        prepared["Обогнали_меня_всего_≥0_%"] = 0.0

    tb_column_present = tb_column and tb_column in prepared.columns
    tb_columns = [
        "Обогнал_ТерБанк_%",
        "Обогнали_меня_ТерБанк_%",
        "Обогнал_ТерБанк_≥0_%",
        "Обогнали_меня_ТерБанк_≥0_%",
    ]
    if tb_column_present:
        for column in tb_columns:
            prepared[column] = 0.0

        for _, group in prepared.groupby(tb_column):
            subset_values = values.loc[group.index]
            obognal_tb, obognali_tb = _compute_percentile_pair(subset_values)
            prepared.loc[group.index, "Обогнал_ТерБанк_%"] = obognal_tb
            prepared.loc[group.index, "Обогнали_меня_ТерБанк_%"] = obognali_tb

            tb_mask = subset_values >= 0
            if tb_mask.any():
                obognal_tb_ge0, obognali_tb_ge0 = _compute_percentile_pair(
                    subset_values[tb_mask]
                )
                idx = subset_values[tb_mask].index
                prepared.loc[idx, "Обогнал_ТерБанк_≥0_%"] = obognal_tb_ge0
                prepared.loc[idx, "Обогнали_меня_ТерБанк_≥0_%"] = obognali_tb_ge0
    else:
        for column in tb_columns:
            prepared[column] = 0.0

    percent_columns = [
        "Обогнал_всего_%",
        "Обогнали_меня_всего_%",
        "Обогнал_всего_≥0_%",
        "Обогнали_меня_всего_≥0_%",
        "Обогнал_ТерБанк_%",
        "Обогнали_меня_ТерБанк_%",
        "Обогнал_ТерБанк_≥0_%",
        "Обогнали_меня_ТерБанк_≥0_%",
    ]
    for column in percent_columns:
        if column in prepared.columns:
            prepared[column] = prepared[column].round(2)

    return prepared


def build_scenario_keys(key_mode: str, include_tb: bool) -> List[str]:
    """Возвращает список колонок для ключа сценария."""

    mapping = {
        "client": ["client_id"],
        "manager": ["manager_id"],
    }
    if key_mode not in mapping:
        raise ValueError("key_mode должен быть client или manager")
    keys = list(mapping[key_mode])
    if include_tb:
        keys.append("tb")
    return keys


def _ensure_manager_identity(
    value: Any,
    *,
    default_value: str,
    identifiers: Mapping[str, Any],
) -> str:
    """Возвращает табельный номер с учётом обязательной длины."""

    if not value or str(value).strip() == "":
        return format_identifier(
            default_value,
            total_length=identifiers["manager_id"]["total_length"],
            fill_char=identifiers["manager_id"]["fill_char"],
        )
    return format_identifier(
        value,
        total_length=identifiers["manager_id"]["total_length"],
        fill_char=identifiers["manager_id"]["fill_char"],
    )


def build_assignment_table(
    variant_df: pd.DataFrame,
    *,
    key_columns: List[str],
    manager_assignment: str,
    defaults: Mapping[str, Any],
    identifiers: Mapping[str, Any],
    logger: Mapping[str, Any],
    scenario_name: str,
) -> pd.DataFrame:
    """Возвращает таблицу назначений (ключ ↔ выбранный КМ)."""

    if manager_assignment not in {"latest", "per_file"}:
        raise ValueError("manager_assignment должен быть latest или per_file.")

    if variant_df.empty:
        columns = key_columns + [
            SELECTED_MANAGER_ID_COL,
            SELECTED_MANAGER_NAME_COL,
            "Источник",
            "Факт_T0",
            "Факт_T1",
            "Прирост",
        ]
        return pd.DataFrame(columns=columns)

    if manager_assignment == "latest":
        assignments = variant_df[key_columns].copy()
        assignments[SELECTED_MANAGER_ID_COL] = variant_df["Таб. номер ВКО_Актуальный"]
        assignments[SELECTED_MANAGER_NAME_COL] = variant_df["ВКО_Актуальный"]
        assignments["Источник"] = "LATEST"
        assignments["Факт_T0"] = variant_df["Факт_T0"].fillna(0.0)
        assignments["Факт_T1"] = variant_df["Факт_T1"].fillna(0.0)
        assignments["Прирост"] = variant_df["Прирост"].fillna(0.0)
        log_debug(
            logger,
            f"{scenario_name}: назначено {len(assignments)} записей (режим latest)",
            class_name="Scenario",
            func_name="build_assignment_table",
        )
        return assignments

    default_name = defaults["manager_name"]
    default_id = _ensure_manager_identity(
        defaults["manager_id"], default_value=defaults["manager_id"], identifiers=identifiers
    )

    records: List[Dict[str, Any]] = []
    for row in variant_df.itertuples(index=False):
        base = {column: getattr(row, column, None) for column in key_columns}
        growth = getattr(row, "Прирост", 0.0) or 0.0
        fact_t0 = getattr(row, "Факт_T0", 0.0) or 0.0
        fact_t1 = getattr(row, "Факт_T1", 0.0) or 0.0

        manager_t0 = getattr(row, "Таб. номер ВКО_T0", None)
        manager_name_t0 = getattr(row, "ВКО_T0", None) or default_name
        manager_t0 = _ensure_manager_identity(
            manager_t0 or default_id, default_value=default_id, identifiers=identifiers
        )

        manager_t1 = getattr(row, "Таб. номер ВКО_T1", None)
        manager_name_t1 = getattr(row, "ВКО_T1", None) or default_name
        manager_t1 = _ensure_manager_identity(
            manager_t1 or default_id, default_value=default_id, identifiers=identifiers
        )

        if fact_t0 or growth > 0:
            record = {
                **base,
                SELECTED_MANAGER_ID_COL: manager_t0,
                SELECTED_MANAGER_NAME_COL: manager_name_t0,
                "Источник": "T0",
                "Факт_T0": fact_t0,
                "Факт_T1": 0.0,
                "Прирост": max(growth, 0.0),
            }
            records.append(record)

        if fact_t1 or growth < 0:
            record = {
                **base,
                SELECTED_MANAGER_ID_COL: manager_t1,
                SELECTED_MANAGER_NAME_COL: manager_name_t1,
                "Источник": "T1",
                "Факт_T0": 0.0,
                "Факт_T1": fact_t1,
                "Прирост": min(growth, 0.0),
            }
            records.append(record)

    assignments = pd.DataFrame(records)
    if assignments.empty:
        columns = key_columns + [
            SELECTED_MANAGER_ID_COL,
            SELECTED_MANAGER_NAME_COL,
            "Источник",
            "Факт_T0",
            "Факт_T1",
            "Прирост",
        ]
        assignments = pd.DataFrame(columns=columns)

    log_debug(
        logger,
        f"{scenario_name}: назначено {len(assignments)} записей (режим per_file)",
        class_name="Scenario",
        func_name="build_assignment_table",
    )
    return assignments


def build_assignment_summary(
    assignment_df: pd.DataFrame,
    *,
    include_tb: bool,
    logger: Mapping[str, Any],
    summary_name: str,
) -> pd.DataFrame:
    """Суммирует факты/приросты по выбранным менеджерам."""

    group_columns = [SELECTED_MANAGER_ID_COL, SELECTED_MANAGER_NAME_COL]
    tb_column_name: Optional[str] = None
    if include_tb:
        if "ТБ" in assignment_df.columns:
            tb_column_name = "ТБ"
        elif "tb" in assignment_df.columns:
            tb_column_name = "tb"
    if tb_column_name:
        group_columns.append(tb_column_name)

    if assignment_df.empty:
        return pd.DataFrame(
            columns=group_columns + ["Факт_T0", "Факт_T1", "Прирост", "Количество записей"]
        )

    numeric_columns = ["Факт_T0", "Факт_T1", "Прирост"]
    summary = (
        assignment_df.groupby(group_columns, dropna=False)[numeric_columns]
        .sum()
        .reset_index()
    )
    counts = assignment_df.groupby(group_columns, dropna=False).size().reset_index(name="Количество записей")
    summary = summary.merge(counts, on=group_columns, how="left")
    if tb_column_name == "tb":
        summary = summary.rename(columns={"tb": "ТБ"})

    log_debug(
        logger,
        f"{summary_name}: сформирована сводная таблица на {len(summary)} менеджеров",
        class_name="Scenario",
        func_name="build_assignment_summary",
    )
    return summary


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
    """Определяет доминантного менеджера (по сумме факта) для каждого ключа.

    Алгоритм:
    1. Формируется составной ключ (например, client_id или client_id+tb) и
       расширяется парой полей менеджера (ФИО + табельный номер), если те ещё
       не входят в key_columns.
    2. Группируем данные по (ключ, manager_id, manager_name) и суммируем
       fact_value_clean, тем самым получаем суммарный вклад каждого ТН по
       конкретному клиенту/объекту.
    3. Выбираем строку с максимальной суммой. Если суммы равны, pandas idxmax
       вернёт первую попавшуюся — этого достаточно по ТЗ.
    """

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
    """Применяет форматирование листа Excel через openpyxl."""

    workbook = writer.book
    worksheet = workbook[sheet_name]

    if df.empty:
        return

    worksheet.freeze_panes = worksheet["A2"]
    worksheet.auto_filter.ref = worksheet.dimensions

    header_alignment = Alignment(wrap_text=True)
    wrap_alignment = Alignment(wrap_text=True)
    number_alignment = Alignment(wrap_text=True)
    header_font = Font(bold=True)

    for cell in next(worksheet.iter_rows(min_row=1, max_row=1)):
        cell.font = header_font
        cell.alignment = header_alignment

    for col_idx, column in enumerate(df.columns, start=1):
        values = [column] + df[column].tolist()
        max_len = max((len(str(value)) for value in values), default=0) + 2
        width = clamp_width(max_len)
        column_letter = get_column_letter(col_idx)
        worksheet.column_dimensions[column_letter].width = width

        if worksheet.max_row >= 2:
            data_range = worksheet[f"{column_letter}2": f"{column_letter}{worksheet.max_row}"]
            if (
                column.startswith("Факт")
                or column == "Прирост"
                or "Обогнал" in column
                or "Обогнали" in column
            ):
                for cell_tuple in data_range:
                    for item in cell_tuple:
                        item.number_format = "#,##0.00"
                        item.alignment = number_alignment
            else:
                for cell_tuple in data_range:
                    for item in cell_tuple:
                        item.alignment = wrap_alignment


def format_decimal_string(value: float, decimals: int = 5) -> str:
    """Форматирует число вида 0.00000."""

    numeric_value = 0.0 if value is None or pd.isna(value) else float(value)
    return f"{numeric_value:.{decimals}f}"


def build_spod_dataset(
    source_table: pd.DataFrame,
    *,
    value_column: str,
    fact_value_filter: str,
    plan_value: float,
    priority: str,
    contest_code: str,
    tournament_code: str,
    contest_date: str,
    identifiers: Mapping[str, Any],
    logger: Mapping[str, Any],
    dataset_name: str,
) -> pd.DataFrame:
    """Готовит данные для загрузки в СПОД."""

    if value_column not in source_table.columns:
        raise KeyError(
            f"Колонка '{value_column}' отсутствует в источнике '{dataset_name}'."
        )

    mask = build_filter_mask(source_table[value_column], fact_value_filter)
    filtered = source_table[mask].copy()
    filtered = filtered.sort_values(by=value_column, ascending=False)

    log_debug(
        logger,
        f"SPOD '{dataset_name}': после фильтра {fact_value_filter} осталось {len(filtered)} строк",
        class_name="Exporter",
        func_name="build_spod_dataset",
    )

    dataset = filtered.rename(
        columns={
            SELECTED_MANAGER_ID_COL: "MANAGER_PERSON_NUMBER",
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
    dataset["CONTEST_CODE"] = contest_code
    dataset["TOURNAMENT_CODE"] = tournament_code
    dataset["CONTEST_DATE"] = parse_contest_date(contest_date)
    dataset["PLAN_VALUE"] = format_decimal_string(plan_value)
    dataset["FACT_VALUE"] = filtered[value_column].apply(format_decimal_string)
    dataset["priority_type"] = priority

    log_debug(
        logger,
        f"SPOD '{dataset_name}': подготовлено {len(dataset)} строк для выгрузки",
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
    spod_variants_config = settings.get("spod_variants", [])
    report_layout = settings.get("report_layout", {})

    def build_whitelist(key: str) -> Optional[Set[str]]:
        """Возвращает множество разрешённых листов для указанного блока."""

        values = report_layout.get(key)
        if values is None:
            return None
        return set(values)

    detail_sheet_whitelist = build_whitelist("detail_sheets")
    summary_sheet_whitelist = build_whitelist("summary_sheets")
    percentile_sheet_whitelist = build_whitelist("percentile_sheets")
    calc_sheet_whitelist = build_whitelist("calc_sheets")
    spod_variant_whitelist = build_whitelist("spod_variants")
    raw_sheet_whitelist = build_whitelist("raw_sheets")

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

    def should_write(entity_name: str, whitelist: Optional[Set[str]], block_name: str) -> bool:
        """Проверяет необходимость выгрузки листа согласно report_layout."""

        if whitelist is None:
            return True
        if entity_name in whitelist:
            return True
        log_debug(
            logger,
            f"Элемент '{entity_name}' пропущен (report_layout ограничил блок '{block_name}')",
            class_name="Exporter",
            func_name="process_project",
        )
        return False

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

        scenario_cfg = settings["scenario"]
        scenario_name = scenario_cfg.get("name", "SCENARIO")
        key_mode = scenario_cfg.get("key_mode", "client")
        include_tb_scenario = scenario_cfg.get("include_tb", False)
        manager_assignment = scenario_cfg.get("manager_assignment", "latest")
        key_columns = build_scenario_keys(key_mode, include_tb_scenario)

        log_info(
            logger,
            f"Сценарий {scenario_name}: key_mode={key_mode}, include_tb={include_tb_scenario}, "
            f"manager_assignment={manager_assignment}",
        )

        scenario_dataset = assemble_variant_dataset(
            variant_name=scenario_name,
            key_columns=key_columns,
            current_df=current_df,
            previous_df=previous_df,
            defaults=defaults,
            identifiers=identifiers,
            logger=logger,
        )

        assignment_df = build_assignment_table(
            scenario_dataset,
            key_columns=key_columns,
            manager_assignment=manager_assignment,
            defaults=defaults,
            identifiers=identifiers,
            logger=logger,
            scenario_name=scenario_name,
        )

        detail_table = rename_output_columns(assignment_df, alias_to_source)
        summary_table = build_assignment_summary(
            assignment_df,
            include_tb=include_tb_scenario,
            logger=logger,
            summary_name=f"{scenario_name}_SUMMARY",
        )
        percentile_table = append_percentile_columns(
            summary_table,
            value_column="Прирост",
            tb_column="ТБ" if include_tb_scenario else None,
        )

        detail_sheet_name = scenario_cfg.get("detail_sheet_name", "DETAIL_SCENARIO")
        summary_sheet_name = scenario_cfg.get("summary_sheet_name", "SUMMARY_TN")
        percentile_sheet_name = scenario_cfg.get("percentile_sheet_name", "PERCENTILE_TN")

        table_registry: Dict[str, pd.DataFrame] = {
            "scenario_assignment": assignment_df,
            "scenario_detail": detail_table,
            "scenario_summary": summary_table,
            "scenario_percentile": percentile_table,
        }
        raw_tables = {
            "RAW_T0": format_raw_sheet(current_df, alias_to_source),
            "RAW_T1": format_raw_sheet(previous_df, alias_to_source),
        }

        def resolve_table(source_type: str, source_name: Any) -> pd.DataFrame:
            if source_type in table_registry:
                return table_registry[source_type]
            if source_type == "raw":
                if source_name not in raw_tables:
                    raise KeyError(f"Недоступный raw-источник '{source_name}'")
                return raw_tables[source_name]
            raise ValueError(
                "Недопустимый source_type. Доступные значения: "
                "scenario_assignment, scenario_detail, scenario_summary, "
                "scenario_percentile, raw."
            )

        if not spod_variants_config:
            raise ValueError(
                "В настройках отсутствуют spod_variants. Добавьте хотя бы один сценарий."
            )

        spod_dataset_tables: Dict[str, pd.DataFrame] = {}
        csv_frames: List[pd.DataFrame] = []
        calc_sheets_to_write: List[Dict[str, Any]] = []
        spod_sheet_names: Dict[str, str] = {}

        for spod_cfg in spod_variants_config:
            source_type = spod_cfg.get("source_type", "scenario_summary")
            source_name = spod_cfg.get("source_name")
            source_table = resolve_table(source_type, source_name)
            value_column = spod_cfg.get("value_column", "Прирост")
            dataset = build_spod_dataset(
                source_table,
                value_column=value_column,
                fact_value_filter=spod_cfg.get("fact_value_filter", "all"),
                plan_value=spod_cfg.get("plan_value", 0.0),
                priority=spod_cfg.get("priority", "1"),
                contest_code=spod_cfg["contest_code"],
                tournament_code=spod_cfg["tournament_code"],
                contest_date=spod_cfg["contest_date"],
                identifiers=identifiers,
                logger=logger,
                dataset_name=spod_cfg["name"],
            )
            spod_dataset_tables[spod_cfg["name"]] = dataset
            spod_sheet_names[spod_cfg["name"]] = spod_cfg.get(
                "spod_sheet_name", spod_cfg["name"]
            )
            if spod_cfg.get("include_in_csv"):
                csv_frames.append(dataset)
            calc_sheet_name = spod_cfg.get("calc_sheet_name")
            if calc_sheet_name:
                calc_sheets_to_write.append(
                    {
                        "sheet_name": calc_sheet_name,
                        "table": source_table,
                        "owner": spod_cfg["name"],
                    }
                )

        report_suffix = timestamp_suffix()
        excel_name = f"{spod_config['file_prefix']}{report_suffix}.xlsx"
        excel_path = output_dir / excel_name
        log_info(logger, f"Сохраняю Excel-файл {excel_name}")

        log_info(
            logger,
            "Используется движок openpyxl (доступен в базовой поставке Anaconda) для сохранения отчёта.",
        )

        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            written_sheets: Set[str] = set()

            def write_sheet(sheet_name: str, table: pd.DataFrame) -> None:
                if sheet_name in written_sheets:
                    log_debug(
                        logger,
                        f"Лист {sheet_name} уже создан — пропускаю повторную запись",
                        class_name="Exporter",
                        func_name="process_project",
                    )
                    return
                table.to_excel(writer, sheet_name=sheet_name, index=False)
                format_excel_sheet(writer, sheet_name, table)
                written_sheets.add(sheet_name)

            if should_write(detail_sheet_name, detail_sheet_whitelist, "detail_sheets"):
                write_sheet(detail_sheet_name, detail_table)

            if should_write(summary_sheet_name, summary_sheet_whitelist, "summary_sheets"):
                write_sheet(summary_sheet_name, summary_table)

            if should_write(percentile_sheet_name, percentile_sheet_whitelist, "percentile_sheets"):
                write_sheet(percentile_sheet_name, percentile_table)

            for calc_meta in calc_sheets_to_write:
                owner_name = calc_meta["owner"]
                if not should_write(owner_name, spod_variant_whitelist, "spod_variants"):
                    continue
                sheet_name = calc_meta["sheet_name"]
                if not should_write(sheet_name, calc_sheet_whitelist, "calc_sheets"):
                    continue
                write_sheet(sheet_name, calc_meta["table"])

            for spod_name, dataset in spod_dataset_tables.items():
                if not should_write(spod_name, spod_variant_whitelist, "spod_variants"):
                    continue
                sheet_name = spod_sheet_names.get(spod_name, spod_name)
                write_sheet(sheet_name, dataset)

            for sheet_name, raw_table in raw_tables.items():
                if not should_write(sheet_name, raw_sheet_whitelist, "raw_sheets"):
                    continue
                write_sheet(sheet_name, raw_table)

        csv_name = f"{spod_config['file_prefix']}_SPOD{report_suffix}.csv"
        csv_path = output_dir / csv_name
        if csv_frames:
            log_info(logger, f"Сохраняю CSV-файл {csv_name}")
            csv_dataset = pd.concat(csv_frames, ignore_index=True)
            csv_dataset.to_csv(csv_path, index=False, sep=";", quoting=csv.QUOTE_MINIMAL)
        else:
            log_info(
                logger,
                "CSV-файл SPOD не сформирован: нет вариантов с include_in_csv=True",
            )

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
