from __future__ import annotations

import pytest

from inventory_app import (
    format_inventory_entries,
    parse_float_field,
    parse_int_field,
)
from inventory_launcher import InventoryEntry


def test_format_inventory_entries_returns_empty_message_when_no_entries() -> None:
    assert format_inventory_entries([]) == "Nenhum item encontrado"
    return None


def test_format_inventory_entries_builds_lines() -> None:
    entries = [InventoryEntry("A", 1.0), InventoryEntry("B", 2.0)]
    assert format_inventory_entries(entries) == "A\t1\nB\t2"
    return None


def test_parse_int_field_accepts_valid_integer() -> None:
    assert parse_int_field("3", "Linha inicial") == 3
    return None


def test_parse_int_field_rejects_non_integer_and_low_values() -> None:
    with pytest.raises(ValueError):
        parse_int_field("abc", "Linha inicial")
    with pytest.raises(ValueError):
        parse_int_field("0", "Linha inicial")
    return None


def test_parse_float_field_accepts_dot_and_comma() -> None:
    assert parse_float_field("0.2", "Pausa tecla") == 0.2
    assert parse_float_field("0,3", "Pausa tecla") == 0.3
    return None


def test_parse_float_field_rejects_invalid_and_negative() -> None:
    with pytest.raises(ValueError):
        parse_float_field("abc", "Pausa tecla")
    with pytest.raises(ValueError):
        parse_float_field("-1", "Espera inicial")
    return None