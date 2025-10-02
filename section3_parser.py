"""Парсер данных раздела 3 из файла ОНХ.

Модуль предназначен для извлечения сведений об "иных нежилых помещениях"
из произвольного табличного источника (обычно Excel).  В исходных файлах
часто встречаются заголовки с техническими префиксами (например, "AD§"),
неоднозначные обозначения и числовые значения с единицами измерения.

Функция :func:`extract_section3_rows` принимает набор записей (``list`` из
``dict`` или, например, ``pandas.DataFrame``) и возвращает список словарей
в формате, удобном для последующей выгрузки в шаблон раздела 3.
"""

from __future__ import annotations

import math
import re
from typing import Dict, Iterable, Iterator, List, Mapping, MutableMapping, Optional, Sequence

# Порядок полей соответствует структуре раздела 3.
SECTION3_FIELDS: Sequence[str] = (
    "unit_number",
    "purpose",
    "floor",
    "entrance_number",
    "area",
    "room_name",
    "part_area",
    "ceiling_height",
)

# Поля с числовыми значениями (их нужно приводить к float).
NUMERIC_FIELDS = {"area", "part_area", "ceiling_height"}


def extract_section3_rows(source: object) -> List[Dict[str, Optional[object]]]:
    """Извлекает строки для раздела 3 из произвольного табличного источника.

    Parameters
    ----------
    source:
        Любой объект, содержащий табличные данные. Это может быть список
        словарей, словарь списков, ``pandas.DataFrame`` и т.п.  Каждая запись
        должна содержать значения с колонками, соответствующими графам
        раздела 3.

    Returns
    -------
    list of dict
        Список словарей. В каждом словаре присутствуют ключи из
        :data:`SECTION3_FIELDS`.  Для числовых колонок значения
        преобразуются к ``float``. Пустые строки пропускаются.
    """

    records = list(_coerce_records(source))
    if not records:
        return []

    header_map = _build_header_map(record.keys() for record in records)
    if not header_map:
        return []

    rows: List[Dict[str, Optional[object]]] = []
    for record in records:
        normalized = _extract_row(record, header_map)
        if normalized is None:
            continue
        rows.append(normalized)
    return rows


def _coerce_records(source: object) -> Iterator[Mapping[str, object]]:
    """Преобразует произвольный источник в последовательность словарей."""

    if source is None:
        return iter(())

    # Поддержка pandas.DataFrame (необязательный импорт).
    to_dict = getattr(source, "to_dict", None)
    if callable(to_dict):
        try:
            converted = to_dict(orient="records")  # type: ignore[call-arg]
        except TypeError:
            converted = to_dict()
        if isinstance(converted, list):
            return (row for row in converted if isinstance(row, Mapping))

    if isinstance(source, Mapping):
        if all(isinstance(values, Sequence) for values in source.values()):
            length = max((len(values) for values in source.values()), default=0)
            rows = []
            for index in range(length):
                row: Dict[str, object] = {}
                for key, values in source.items():
                    if index < len(values):
                        row[str(key)] = values[index]
                rows.append(row)
            return iter(rows)
        return iter((source,))

    if isinstance(source, Sequence):
        rows = []
        for item in source:
            if isinstance(item, Mapping):
                rows.append(item)
        return iter(rows)

    return iter(())


def _build_header_map(
    header_sequences: Iterable[Iterable[str]],
) -> Dict[str, str]:
    """Формирует отображение ключей раздела на реальные названия колонок."""

    result: Dict[str, str] = {}
    priorities: Dict[str, int] = {}

    for headers in header_sequences:
        for header in headers:
            normalized = _normalize_header(header)
            if not normalized:
                continue
            match = _match_field(normalized)
            if match is None:
                continue
            field, priority = match
            if field in result and priorities[field] <= priority:
                continue
            result[field] = header
            priorities[field] = priority

    return result


def _extract_row(
    record: Mapping[str, object], header_map: Mapping[str, str]
) -> Optional[Dict[str, Optional[object]]]:
    """Преобразует исходный словарь в нормализованную запись."""

    row: Dict[str, Optional[object]] = {}
    meaningful = False

    for field in SECTION3_FIELDS:
        value = None
        header = header_map.get(field)
        if header is not None and header in record:
            value = record.get(header)

        if field in NUMERIC_FIELDS:
            parsed = _parse_number(value)
            if parsed is not None and not math.isnan(parsed):
                row[field] = parsed
                meaningful = True
            else:
                row[field] = None
        else:
            text = _clean_string(value)
            row[field] = text
            if text:
                meaningful = True

    if not meaningful:
        return None

    # Если площадь части отсутствует, но есть общая площадь — используем её.
    if row.get("part_area") is None and row.get("area") is not None:
        row["part_area"] = row["area"]

    return row


def _normalize_header(value: object) -> str:
    """Подготавливает заголовок для сравнения (убирает префиксы, пробелы)."""

    if value is None:
        return ""

    text = str(value)
    if "§" in text:
        text = text.split("§", 1)[1]

    text = (
        text.replace("\xa0", " ")
        .replace("\r", " ")
        .replace("\n", " ")
        .replace("_", " ")
        .replace("№", "номер")
    )
    text = text.lower().replace("ё", "е")
    text = re.sub(r"[^a-z0-9а-я ]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _match_field(normalized_header: str) -> Optional[tuple[str, int]]:
    """Определяет, к какому полю относится заголовок."""

    if not normalized_header:
        return None

    lookup = HEADER_LOOKUP.get(normalized_header)
    if lookup:
        return lookup, 0

    checks: List[tuple[str, int]] = []

    if "условн" in normalized_header and "номер" in normalized_header:
        checks.append(("unit_number", 1))
    if "назнач" in normalized_header:
        checks.append(("purpose", 1))
    if "этаж" in normalized_header or "уровен" in normalized_header:
        checks.append(("floor", 1))
    if "подъезд" in normalized_header or "секци" in normalized_header:
        checks.append(("entrance_number", 1))
    if "наимен" in normalized_header and "помещ" in normalized_header:
        checks.append(("room_name", 1))
    if "высот" in normalized_header and "потол" in normalized_header:
        checks.append(("ceiling_height", 1))
    if "площад" in normalized_header:
        if "част" in normalized_header:
            checks.append(("part_area", 1))
        else:
            checks.append(("area", 2))

    if not checks:
        return None

    return min(checks, key=lambda item: item[1])


def _clean_string(value: object) -> Optional[str]:
    """Подготавливает строковое значение (обрезает пробелы, проверяет пустоту)."""

    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None
    return text


def _parse_number(value: object) -> Optional[float]:
    """Пытается извлечь ``float`` из произвольного значения."""

    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    cleaned = re.sub(r"[^0-9,.-]+", "", text)
    if not cleaned:
        return None
    cleaned = cleaned.replace(",", ".")

    match = re.search(r"[-+]?[0-9]+(?:\.[0-9]+)?", cleaned)
    if not match:
        return None

    try:
        return float(match.group(0))
    except ValueError:
        return None


HEADER_SYNONYMS = {
    "unit_number": (
        "Условный номер",
        "Условный номер помещения",
        "Условный номер блока",
        "Условный номер части",
        "Номер помещения",
    ),
    "purpose": (
        "Назначение",
        "Назначение помещения",
        "Функциональное назначение",
    ),
    "floor": (
        "Этаж расположения",
        "Этаж",
        "Уровень",
        "Этаж/Уровень",
    ),
    "entrance_number": (
        "Номер подъезда",
        "Подъезд",
        "Подъезд / Секция",
        "Секция",
    ),
    "area": (
        "Площадь, м2",
        "Площадь помещения, м2",
        "Общая площадь",
        "Площадь",
    ),
    "room_name": (
        "Наименование помещения",
        "Наименование",
        "Наименование части помещения",
    ),
    "part_area": (
        "Площадь части, м2",
        "Площадь части помещения, м2",
        "Площадь части",
    ),
    "ceiling_height": (
        "Высота потолков, м",
        "Высота потолков",
        "Высота",
    ),
}


def _build_lookup() -> Dict[str, str]:
    lookup: Dict[str, str] = {}
    for field, synonyms in HEADER_SYNONYMS.items():
        for synonym in synonyms:
            normalized = _normalize_header(synonym)
            if normalized:
                lookup[normalized] = field
    return lookup


HEADER_LOOKUP = _build_lookup()


__all__ = [
    "SECTION3_FIELDS",
    "NUMERIC_FIELDS",
    "extract_section3_rows",
]

