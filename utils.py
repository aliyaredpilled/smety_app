# utils.py
"""Вспомогательные функции для обработки Excel‑смет."""

from __future__ import annotations

import logging
import re
from typing import Any, Optional

logger = logging.getLogger(__name__)
_NUMBER_TOLERANCE = 1e-9


def is_likely_empty(value: Any) -> bool:
    """Возвращает *True*, если значение «пустое» (None, пустая строка, пробелы)."""
    if value == 0 and isinstance(value, (int, float)):
        return False
    return value is None or str(value).strip() == ""


def check_merge(ws, row: int, start_col: int, end_col: int) -> Optional[str]:
    """
    Проверяет, объединены ли ячейки (row, start_col‑end_col).

    Возвращает адрес диапазона *A1:C3* либо *None*.
    """
    try:
        for rng in ws.merged_cells.ranges:  # type: ignore[attr-defined]
            if rng.min_row <= row <= rng.max_row:
                if (
                    (rng.min_col == start_col + 1 and rng.max_col == end_col + 1)
                    or (rng.min_col <= start_col + 1 <= end_col + 1 <= rng.max_col)
                ):
                    return rng.coord
    except AttributeError:
        logger.warning(
            "У листа нет атрибута 'merged_cells' — пропускаю проверку merge для строки %s",
            row,
        )
    except Exception as exc:  # noqa: BLE001
        logger.warning(
            "Не удалось проверить merge для строки %s, столбцы %s‑%s: %s",
            row,
            start_col + 1,
            end_col + 1,
            exc,
        )
    return None


def get_start_coord(coord: Optional[str]) -> Optional[str]:
    """Для диапазона 'A1:C3' вернёт 'A1'. Для одиночных ячеек вернёт адрес без изменений."""
    if isinstance(coord, str) and ":" in coord:
        return coord.split(":", 1)[0]
    return coord


def is_zero(value: Any) -> bool:
    """Проверяет, является ли значение числом «0» (строка '0', '0,0', 0 и т. д.)."""
    if is_likely_empty(value):
        return False
    try:
        return abs(float(str(value).replace(",", ".").strip())) < _NUMBER_TOLERANCE
    except (ValueError, TypeError):
        return False


def get_item_id_nature(value: Any) -> str:
    """
    Определяет тип идентификатора позиции: 'integer', 'decimal', 'not_a_number'.
    """
    if is_likely_empty(value):
        return "not_a_number"

    raw = str(value).strip()
    raw = re.sub(r"(?<=[.,])\s*(?=\d)", "", raw)  # '1. 2' -> '1.2'
    numeric = raw.replace(",", ".")

    try:
        num = float(numeric)
        if abs(num - int(num)) < _NUMBER_TOLERANCE:
            return "integer"
        return "decimal" if any(c in raw for c in ",.") else "not_a_number"
    except (ValueError, TypeError):
        # Пробуем вытащить числовой префикс
        match = re.match(r"^(\d+(?:[.,]\d+)?)\b", raw)
        if match:
            try:
                part = float(match.group(1).replace(",", "."))
                return "integer" if abs(part - int(part)) < _NUMBER_TOLERANCE else "decimal"
            except (ValueError, TypeError):
                pass
        return "not_a_number"
