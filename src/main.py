"""Основной модуль расчёта приростов по клиентам СПОД.

Вся бизнес-логика собрана в одном файле main.py согласно требованиям.
"""

from __future__ import annotations

import csv
import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional

import pandas as pd


# ----------------------------- Конфигурация ---------------------------------


@dataclass
class ColumnConfig:
    """Описание исходных колонок и их единых идентификаторов."""

    tb: str = "ТБ"
    gosb: str = "ГОСБ"
    manager_name: str = "ВКО"
    manager_id: str = "Таб. номер ВКО"
    client_id: str = "ИНН"
    fact: str = "Остаток срочной задолженности по основному долгу"

    def rename_map(self) -> Dict[str, str]:
        """Возвращает словарь для перевода русских колонок в единые имена."""

        return {
            self.tb: "tb",
            self.gosb: "gosb",
            self.manager_name: "manager_name",
            self.manager_id: "manager_id",
            self.client_id: "client_id",
            self.fact: "fact_value",
        }


@dataclass
class ProcessingConfig:
    """Базовые параметры обработки файлов."""

    sheet_name: str = "Sheet1"
    current_file: str = "АКТИВЫ 31-10-2025 (ОСТАТОК-V2).xlsx"
    previous_file: str = "АКТИВЫ 31-12-2024 (ОСТАТОК-V2).xlsx"
    drop_rules: Mapping[str, Iterable[str]] = field(
        default_factory=lambda: {
            "manager_name": ("-", "Серая зона"),
            "manager_id": ("-", "Green_Zone", "Tech_Sib"),
            "client_id": ("Report_id не определен",),
        }
    )
    default_manager_name: str = "Не найден КМ"
    default_manager_id: str = "90000009"
    tn_fill_char: str = "0"
    tn_total_length: int = 8
    inn_fill_char: str = "0"
    inn_total_length: int = 12


SettingsTree = List[Dict[str, Any]]


@dataclass
class OutputConfig:
    """Настройки формирования выходных файлов."""

    file_prefix: str
    log_topic: str
    contest_code: str
    tournament_code: str
    contest_date_initial: str
    plan_value: float
    spod_priority: str = "1"

    def contest_date_iso(self) -> str:
        """Возвращает дату турнира в формате ISO."""

        parsed = dt.datetime.strptime(self.contest_date_initial, "%d/%m/%Y")
        return parsed.strftime("%Y-%m-%d")


# --------------------------- Вспомогательные функции ------------------------


def build_settings_tree() -> SettingsTree:
    """Возвращает иерархию настроек, сгруппированную по темам."""

    return [
        {
            "section": "spod",
            "title": "Выгрузки СПОД и логирование",
            "values": {
                "file_prefix": "YEAR_SPOD_Active_Rost_Ost",
                "log_topic": "spod",
                "plan_value": 0.0,
                "priority": "1",
            },
        },
        {
            "section": "contest",
            "title": "Параметры турнира",
            "values": {
                "contest_code": "01_2025-2_14-1_2",
                "tournament_code": "t_01_2025-2_14-1_2_1001",
                "contest_date": "31/10/2025",
            },
        },
        {
            "section": "defaults",
            "title": "Значения по умолчанию",
            "values": {
                "manager_name": "Не найден КМ",
                "manager_tn": "90000009",
            },
        },
        {
            "section": "identifiers",
            "title": "Преобразование идентификаторов",
            "values": {
                "tn_fill_char": "0",
                "tn_total_length": 8,
                "inn_fill_char": "0",
                "inn_total_length": 12,
            },
        },
    ]


def get_settings_section(tree: SettingsTree, section_name: str) -> Dict[str, Any]:
    """Возвращает словарь значений нужной секции."""

    for section in tree:
        if section["section"] == section_name:
            return section["values"]
    raise KeyError(f"Секция настроек {section_name} не найдена")


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


def normalize_string(value: Any) -> str:
    """Возвращает очищенную строку без None."""

    if value is None:
        return ""
    return str(value).strip()


class StructuredLogger:
    """Простой логгер, удовлетворяющий требованиям к форматированию."""

    def __init__(self, info_path: Path, debug_path: Path) -> None:
        self.info_path = info_path
        self.debug_path = debug_path

    def info(self, message: str) -> None:
        line = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} - [INFO] - {message}"
        print(line)
        with self.info_path.open("a", encoding="utf-8") as info_file:
            info_file.write(f"{line}\n")

    def debug(self, message: str, class_name: str, func_name: str) -> None:
        line = (
            f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} - [DEBUG] - "
            f"{message} [class: {class_name} | def: {func_name}]"
        )
        with self.debug_path.open("a", encoding="utf-8") as debug_file:
            debug_file.write(f"{line}\n")


def build_logger(log_dir: Path, topic: str) -> StructuredLogger:
    """Инициализирует файловый логгер."""

    ensure_directories([log_dir])
    suffix = timestamp_suffix()
    info_path = log_dir / f"INFO_{topic}{suffix}.log"
    debug_path = log_dir / f"DEBUG_{topic}{suffix}.log"
    return StructuredLogger(info_path, debug_path)


# -------------------------- Работа с исходными файлами ----------------------


def read_source_file(
    file_path: Path,
    column_config: ColumnConfig,
    processing: ProcessingConfig,
    logger: StructuredLogger,
) -> pd.DataFrame:
    """Загружает исходный Excel и подготавливает данные."""

    if not file_path.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")

    logger.info(f"Загружаю данные из файла {file_path.name}")
    raw_df = pd.read_excel(file_path, sheet_name=processing.sheet_name)
    renamed = raw_df.rename(columns=column_config.rename_map())

    required_columns = list(column_config.rename_map().values())
    missing = [col for col in required_columns if col not in renamed.columns]
    if missing:
        raise ValueError(
            f"Отсутствуют обязательные колонки {missing} в файле {file_path}"
        )

    prepared = renamed[required_columns].copy()

    for column in ("tb", "gosb", "manager_name"):
        prepared[column] = prepared[column].apply(normalize_string)

    prepared["manager_id"] = prepared["manager_id"].apply(
        lambda value: format_identifier(
            value=value,
            total_length=processing.tn_total_length,
            fill_char=processing.tn_fill_char,
        )
    )
    prepared["client_id"] = prepared["client_id"].apply(
        lambda value: format_identifier(
            value=value,
            total_length=processing.inn_total_length,
            fill_char=processing.inn_fill_char,
        )
    )

    prepared["fact_value_clean"] = prepared["fact_value"].apply(safe_to_float)

    cleaned = drop_forbidden_rows(prepared, processing.drop_rules, logger)
    logger.debug(
        f"После очистки в {file_path.name} осталось строк: {len(cleaned)}",
        class_name="DataLoader",
        func_name="read_source_file",
    )
    return cleaned


def drop_forbidden_rows(
    df: pd.DataFrame,
    drop_rules: Mapping[str, Iterable[str]],
    logger: StructuredLogger,
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
        logger.debug(
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
    logger: StructuredLogger,
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
    logger.debug(
        f"{variant_name}: агрегировано {len(renamed)} строк для суффикса {suffix}",
        class_name="Aggregator",
        func_name="aggregate_facts",
    )
    return renamed


def select_best_manager(
    df: pd.DataFrame,
    key_columns: List[str],
    logger: StructuredLogger,
    variant_name: str,
) -> pd.DataFrame:
    """Определяет менеджера с максимальным фактом для ключа."""

    grouping_columns = key_columns + ["manager_name", "manager_id"]
    grouped = (
        df[grouping_columns + ["fact_value_clean"]]
        .fillna({"fact_value_clean": 0.0})
        .groupby(grouping_columns, dropna=False, as_index=False)
        .sum(numeric_only=True)
    )
    idx = grouped.groupby(key_columns, dropna=False)["fact_value_clean"].idxmax()
    best = grouped.loc[idx, key_columns + ["manager_name", "manager_id"]].copy()
    result = best.rename(
        columns={"manager_name": "ВКО", "manager_id": "Таб. номер ВКО"}
    )
    logger.debug(
        f"{variant_name}: выбраны менеджеры для {len(result)} ключей",
        class_name="Aggregator",
        func_name="select_best_manager",
    )
    return result


def build_latest_manager(
    current_best: pd.DataFrame,
    previous_best: pd.DataFrame,
    key_columns: List[str],
    default_name: str,
    default_id: str,
    logger: StructuredLogger,
    variant_name: str,
) -> pd.DataFrame:
    """Комбинирует менеджеров, отдавая приоритет файлу T-0."""

    curr = current_best.set_index(key_columns) if not current_best.empty else pd.DataFrame(columns=key_columns + ["ВКО", "Таб. номер ВКО"]).set_index(key_columns)
    prev = previous_best.set_index(key_columns) if not previous_best.empty else pd.DataFrame(columns=key_columns + ["ВКО", "Таб. номер ВКО"]).set_index(key_columns)

    combined = prev.join(
        curr,
        how="outer",
        lsuffix="_prev",
        rsuffix="_curr",
    )
    combined["ВКО_Актуальный"] = combined["ВКО_curr"].combine_first(combined["ВКО_prev"]).fillna(default_name)
    combined["Таб. номер ВКО_Актуальный"] = combined["Таб. номер ВКО_curr"].combine_first(combined["Таб. номер ВКО_prev"]).fillna(default_id)

    result = combined.reset_index()[key_columns + ["ВКО_Актуальный", "Таб. номер ВКО_Актуальный"]]
    logger.debug(
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
    processing: ProcessingConfig,
    logger: StructuredLogger,
) -> pd.DataFrame:
    """Формирует таблицу для конкретного варианта ключа."""

    logger.debug(
        f"{variant_name}: старт построения набора данных",
        class_name="Aggregator",
        func_name="assemble_variant_dataset",
    )

    agg_current = aggregate_facts(current_df, key_columns, "T0", logger, variant_name)
    agg_previous = aggregate_facts(previous_df, key_columns, "T1", logger, variant_name)
    merged = pd.merge(agg_current, agg_previous, on=key_columns, how="outer").fillna({"Факт_T0": 0.0, "Факт_T1": 0.0})
    merged["Прирост"] = merged["Факт_T0"] - merged["Факт_T1"]

    best_current = select_best_manager(current_df, key_columns, logger, variant_name).rename(
        columns={"ВКО": "ВКО_T0", "Таб. номер ВКО": "Таб. номер ВКО_T0"}
    )
    best_previous = select_best_manager(previous_df, key_columns, logger, variant_name).rename(
        columns={"ВКО": "ВКО_T1", "Таб. номер ВКО": "Таб. номер ВКО_T1"}
    )

    merged = merged.merge(best_current, on=key_columns, how="left")
    merged = merged.merge(best_previous, on=key_columns, how="left")

    latest = build_latest_manager(
        current_best=best_current.rename(columns={"ВКО_T0": "ВКО", "Таб. номер ВКО_T0": "Таб. номер ВКО"}),
        previous_best=best_previous.rename(columns={"ВКО_T1": "ВКО", "Таб. номер ВКО_T1": "Таб. номер ВКО"}),
        key_columns=key_columns,
        default_name=processing.default_manager_name,
        default_id=format_identifier(
            processing.default_manager_id,
            total_length=processing.tn_total_length,
            fill_char=processing.tn_fill_char,
        ),
        logger=logger,
        variant_name=variant_name,
    )
    merged = merged.merge(latest, on=key_columns, how="left")

    logger.debug(
        f"{variant_name}: итоговый набор содержит {len(merged)} строк",
        class_name="Aggregator",
        func_name="assemble_variant_dataset",
    )
    return merged


def build_manager_summary(
    variant_df: pd.DataFrame,
    include_tb: bool,
    logger: StructuredLogger,
    summary_name: str,
) -> pd.DataFrame:
    """Создаёт свод по уникальным ТН+ВКО (+ТБ опционально)."""

    group_columns = ["Таб. номер ВКО_Актуальный", "ВКО_Актуальный"]
    if include_tb and "tb" in variant_df.columns:
        group_columns.append("tb")

    grouped = (
        variant_df.groupby(group_columns, dropna=False)[["Факт_T0", "Факт_T1", "Прирост"]]
        .sum()
        .reset_index()
    )
    if "tb" in grouped.columns:
        grouped = grouped.rename(columns={"tb": "ТБ"})

    logger.debug(
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
    output_config: OutputConfig,
    processing: ProcessingConfig,
    logger: StructuredLogger,
) -> pd.DataFrame:
    """Готовит данные для загрузки в СПОД."""

    dataset = manager_summary.rename(
        columns={
            "Таб. номер ВКО_Актуальный": "MANAGER_PERSON_NUMBER",
            "Прирост": "FACT_VALUE",
        }
    )["MANAGER_PERSON_NUMBER"].to_frame()

    dataset["MANAGER_PERSON_NUMBER"] = dataset["MANAGER_PERSON_NUMBER"].apply(
        lambda value: format_identifier(
            value=value,
            total_length=max(processing.tn_total_length, 20),
            fill_char=processing.tn_fill_char,
        )
    )
    dataset["CONTEST_CODE"] = output_config.contest_code
    dataset["TOURNAMENT_CODE"] = output_config.tournament_code
    dataset["CONTEST_DATE"] = output_config.contest_date_iso()
    dataset["PLAN_VALUE"] = format_decimal_string(output_config.plan_value)
    dataset["FACT_VALUE"] = manager_summary["Прирост"].apply(format_decimal_string)
    dataset["priority_type"] = output_config.spod_priority

    logger.debug(
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


def rename_output_columns(df: pd.DataFrame, column_config: ColumnConfig) -> pd.DataFrame:
    """Возвращает DataFrame с русскими заголовками для ключей."""

    renamed = df.copy()
    mapping = {
        "client_id": column_config.client_id,
        "tb": column_config.tb,
        "manager_id": column_config.manager_id,
    }
    renamed = renamed.rename(columns=mapping)
    return renamed


# ----------------------------- Основной сценарий ----------------------------


def process_project(project_root: Path) -> None:
    """Запускает полный цикл обработки данных."""

    column_config = ColumnConfig()
    processing = ProcessingConfig()
    settings_tree = build_settings_tree()
    spod_settings = get_settings_section(settings_tree, "spod")
    contest_settings = get_settings_section(settings_tree, "contest")
    default_settings = get_settings_section(settings_tree, "defaults")
    identifier_settings = get_settings_section(settings_tree, "identifiers")

    output_config = OutputConfig(
        file_prefix=spod_settings["file_prefix"],
        log_topic=spod_settings["log_topic"],
        contest_code=contest_settings["contest_code"],
        tournament_code=contest_settings["tournament_code"],
        contest_date_initial=contest_settings["contest_date"],
        plan_value=float(spod_settings["plan_value"]),
        spod_priority=spod_settings["priority"],
    )

    processing.default_manager_name = default_settings["manager_name"]
    processing.default_manager_id = default_settings["manager_tn"]
    processing.tn_fill_char = identifier_settings["tn_fill_char"]
    processing.tn_total_length = int(identifier_settings["tn_total_length"])
    processing.inn_fill_char = identifier_settings["inn_fill_char"]
    processing.inn_total_length = int(identifier_settings["inn_total_length"])

    input_dir = project_root / "IN"
    output_dir = project_root / "OUT"
    log_dir = project_root / "log"
    ensure_directories([input_dir, output_dir, log_dir])

    logger = build_logger(log_dir, output_config.log_topic)
    logger.info("Старт обработки проекта YEAR_SPOD_Active_Rost_Ost")

    current_file = input_dir / processing.current_file
    previous_file = input_dir / processing.previous_file

    if not current_file.exists() or not previous_file.exists():
        logger.info(
            "Ожидаемые файлы отсутствуют в каталоге IN. "
            "Положите исходные XLSX и повторите запуск."
        )
        return

    current_df = read_source_file(current_file, column_config, processing, logger)
    previous_df = read_source_file(previous_file, column_config, processing, logger)

    variant_definitions = {
        "ID": ["client_id"],
        "ID_TB": ["client_id", "tb"],
        "ID_TN": ["client_id", "manager_id"],
        "ID_TB_TN": ["client_id", "tb", "manager_id"],
    }

    variant_tables: Dict[str, pd.DataFrame] = {}
    for name, columns in variant_definitions.items():
        logger.info(f"Формирую лист {name}")
        table = assemble_variant_dataset(
            variant_name=name,
            key_columns=columns,
            current_df=current_df,
            previous_df=previous_df,
            processing=processing,
            logger=logger,
        )
        variant_tables[name] = table

    manager_summary = build_manager_summary(
        variant_tables["ID_TN"],
        include_tb=False,
        logger=logger,
        summary_name="TN_VKO",
    )
    manager_summary_tb = build_manager_summary(
        variant_tables["ID_TB_TN"],
        include_tb=True,
        logger=logger,
        summary_name="TN_VKO_TB",
    )

    spod_dataset = build_spod_dataset(manager_summary, output_config, processing, logger)

    report_suffix = timestamp_suffix()
    excel_name = f"{output_config.file_prefix}{report_suffix}.xlsx"
    excel_path = output_dir / excel_name
    logger.info(f"Сохраняю Excel-файл {excel_name}")

    with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
        for sheet_name, table in variant_tables.items():
            printable = rename_output_columns(table, column_config)
            printable.to_excel(writer, sheet_name=sheet_name, index=False)
            format_excel_sheet(writer, sheet_name, printable)

        manager_summary.to_excel(writer, sheet_name="TN_VKO", index=False)
        format_excel_sheet(writer, "TN_VKO", manager_summary)

        manager_summary_tb.to_excel(writer, sheet_name="TN_VKO_TB", index=False)
        format_excel_sheet(writer, "TN_VKO_TB", manager_summary_tb)

        spod_dataset.to_excel(writer, sheet_name="SPOD", index=False)
        format_excel_sheet(writer, "SPOD", spod_dataset)

    csv_name = f"{output_config.file_prefix}_SPOD{report_suffix}.csv"
    csv_path = output_dir / csv_name
    logger.info(f"Сохраняю CSV-файл {csv_name}")
    spod_dataset.to_csv(csv_path, index=False, sep=";", quoting=csv.QUOTE_MINIMAL)

    logger.info("Обработка успешно завершена")


def main() -> None:
    """Точка входа."""

    project_root = Path(__file__).resolve().parent.parent
    process_project(project_root)


if __name__ == "__main__":
    main()
