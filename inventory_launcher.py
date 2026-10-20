from __future__ import annotations

import argparse
from collections.abc import Sequence
from pathlib import Path
from time import sleep
from typing import Literal, NamedTuple, Optional, Protocol

import openpyxl
import pyautogui

ENTER_KEY = "enter"
CODE_COLUMN = 1
COUNT_COLUMN = 2
BRANCH_COLUMN = 3
WAREHOUSE_COLUMN = 4
FIXED_PARAMETERS_ROW = 2
REVERT_COMBO_KEYS = ("shift", "f2")

class InventoryEntry(NamedTuple):
    item_code: str
    count_quantity: float

class AutomationStoppedError(Exception):
    pass


class LaunchResult(NamedTuple):
    status: Literal["concluido", "parado"]
    launched_entries: tuple[InventoryEntry, ...]

class LaunchConfig(NamedTuple):
    sheet_path: Path
    start_row: int
    initial_wait_seconds: float
    key_pause_seconds: float
    simulate: bool


class KeyWriter(Protocol):
    def type_text(self, text: str) -> None: ...

    def press_key(self, key: str) -> None: ...

    def press_combo(self, keys: Sequence[str]) -> None: ...

    def pause_seconds(self, seconds: float) -> None: ...


class GuiKeyboard(Protocol):
    def typewrite(self, message: str, interval: float) -> None: ...

    def press(self, key: str) -> None: ...

    def hotkey(self, *keys: str) -> None: ...


class CellLike(Protocol):
    value: object


class WorksheetLike(Protocol):
    max_row: Optional[int]

    def cell(self, row: int, column: int) -> CellLike: ...


class PyAutoGuiWriter:
    def __init__(self, action_pause_seconds: float, keyboard: GuiKeyboard) -> None:
        self._action_pause_seconds = action_pause_seconds
        self._keyboard = keyboard

    def type_text(self, text: str) -> None:
        if not text:
            return None
        try:
            self._keyboard.typewrite(text, interval=self._action_pause_seconds)
        except pyautogui.FailSafeException as error:
            raise AutomationStoppedError("Automação parada pelo fail-safe") from error
        sleep(self._action_pause_seconds)
        return None

    def press_key(self, key: str) -> None:
        if not key:
            return None
        try:
            self._keyboard.press(key)
        except pyautogui.FailSafeException as error:
            raise AutomationStoppedError("Automação parada pelo fail-safe") from error
        sleep(self._action_pause_seconds)
        return None

    def press_combo(self, keys: Sequence[str]) -> None:
        if not keys:
            return None
        try:
            self._keyboard.hotkey(*keys)
        except pyautogui.FailSafeException as error:
            raise AutomationStoppedError("Automação parada pelo fail-safe") from error
        sleep(self._action_pause_seconds)
        return None

    def pause_seconds(self, seconds: float) -> None:
        if seconds <= 0.0:
            return None
        sleep(seconds)
        return None


def validate_start_row(start_row: int) -> None:
    if start_row < 2:
        raise ValueError("linha-inicial deve ser >= 2 para pular o cabeçalho")
    return None


def is_blank_value(raw_value: object) -> bool:
    if raw_value is None:
        return True
    if isinstance(raw_value, str) and not raw_value.strip():
        return True
    return False


def format_item_code(raw_value: object) -> str:
    if raw_value is None:
        return ""
    if isinstance(raw_value, bool):
        return str(raw_value)
    if isinstance(raw_value, int):
        return str(raw_value)
    if isinstance(raw_value, float) and raw_value.is_integer():
        return str(int(raw_value))
    return str(raw_value).strip()


def validate_count(count_value: float) -> float:
    if count_value < 0:
        raise ValueError("Contagem não pode ser negativa")
    return count_value


def parse_count(raw_value: object) -> float:
    if raw_value is None:
        return 0.0
    if isinstance(raw_value, bool):
        raise ValueError("Contagem não pode ser booleana")
    if isinstance(raw_value, (int, float)):
        return validate_count(float(raw_value))
    text_value = str(raw_value).strip().replace(",", ".")
    if not text_value:
        return 0.0
    try:
        parsed_value = float(text_value)
    except ValueError as error:
        raise ValueError("Contagem deve ser um número não negativo") from error
    return validate_count(parsed_value)


def format_count_text(count_quantity: float) -> str:
    rounded_quantity = round(count_quantity, 4)
    if rounded_quantity.is_integer():
        return str(int(rounded_quantity))
    count_text = f"{rounded_quantity:.4f}".rstrip("0").rstrip(".")
    return count_text.replace(".", ",")


def build_inventory_entry(row_index: int, code_value: object, count_value: object) -> InventoryEntry:
    if is_blank_value(count_value):
        raise ValueError(f"Linha {row_index}: contagem não pode ser vazia")
    item_code = format_item_code(code_value)
    count_quantity = parse_count(count_value)
    if not item_code:
        raise ValueError(f"Linha {row_index}: código não pode ser vazio")
    return InventoryEntry(item_code=item_code, count_quantity=count_quantity)


def collect_worksheet_entries(worksheet: WorksheetLike, start_row: int) -> list[InventoryEntry]:
    last_row = worksheet.max_row
    if last_row is None or last_row < start_row:
        return []
    entries: list[InventoryEntry] = []
    for row_index in range(start_row, last_row + 1):
        code_value: object = worksheet.cell(row=row_index, column=CODE_COLUMN).value
        count_value: object = worksheet.cell(row=row_index, column=COUNT_COLUMN).value
        if is_blank_value(code_value) and is_blank_value(count_value):
            continue
        entries.append(build_inventory_entry(row_index, code_value, count_value))
    return entries


def read_inventory_entries(sheet_path: Path, start_row: int) -> list[InventoryEntry]:
    validate_start_row(start_row)
    workbook = openpyxl.load_workbook(filename=sheet_path, read_only=True)
    worksheet = workbook.active
    if worksheet is None:
        workbook.close()
        raise ValueError("Planilha sem aba ativa")
    entries = collect_worksheet_entries(worksheet, start_row)
    workbook.close()
    return entries


def launch_entry(writer: KeyWriter, entry: InventoryEntry) -> None:
    count_text = format_count_text(entry.count_quantity)
    writer.press_key(ENTER_KEY)
    writer.type_text(entry.item_code)
    writer.press_key(ENTER_KEY)
    writer.press_key(ENTER_KEY)
    writer.press_key(ENTER_KEY)
    writer.press_key(ENTER_KEY)
    writer.type_text(count_text)
    writer.press_key(ENTER_KEY)
    writer.type_text(count_text)
    writer.press_key(ENTER_KEY)
    writer.press_key(ENTER_KEY)
    return None


def launch_inventory_entries(
    writer: KeyWriter,
    entries: Sequence[InventoryEntry],
    initial_wait_seconds: float,
) -> LaunchResult:
    launched_entries: list[InventoryEntry] = []
    try:
        writer.pause_seconds(initial_wait_seconds)
        for entry in entries:
            launch_entry(writer, entry)
            launched_entries.append(entry)
    except AutomationStoppedError:
        return LaunchResult(status="parado", launched_entries=tuple(launched_entries))
    return LaunchResult(status="concluido", launched_entries=tuple(launched_entries))

def read_fixed_parameters(sheet_path: Path) -> tuple[str, str]:
    workbook = openpyxl.load_workbook(filename=sheet_path, read_only=True)
    worksheet = workbook.active
    if worksheet is None:
        workbook.close()
        raise ValueError("Planilha sem aba ativa")
    branch_value: object = worksheet.cell(row=FIXED_PARAMETERS_ROW, column=BRANCH_COLUMN).value
    warehouse_value: object = worksheet.cell(row=FIXED_PARAMETERS_ROW, column=WAREHOUSE_COLUMN).value
    workbook.close()
    branch_code = format_item_code(branch_value)
    warehouse_code = format_item_code(warehouse_value)
    if not branch_code or not warehouse_code:
        raise ValueError("Filial e depósito devem estar preenchidos em C2 e D2")
    return branch_code, warehouse_code


def revert_entry(
    writer: KeyWriter,
    entry: InventoryEntry,
    branch_code: str,
    warehouse_code: str,
) -> None:
    writer.type_text(branch_code)
    writer.press_key(ENTER_KEY)
    writer.type_text(entry.item_code)
    writer.press_key(ENTER_KEY)
    writer.type_text(warehouse_code)
    writer.press_key(ENTER_KEY)
    writer.press_combo(REVERT_COMBO_KEYS)
    writer.press_key(ENTER_KEY)
    return None


def revert_inventory_entries(
    writer: KeyWriter,
    entries: Sequence[InventoryEntry],
    branch_code: str,
    warehouse_code: str,
    initial_wait_seconds: float,
) -> LaunchResult:
    reverted_entries: list[InventoryEntry] = []
    try:
        writer.pause_seconds(initial_wait_seconds)
        for entry in entries:
            revert_entry(writer, entry, branch_code, warehouse_code)
            reverted_entries.append(entry)
    except AutomationStoppedError:
        return LaunchResult(status="parado", launched_entries=tuple(reverted_entries))
    return LaunchResult(status="concluido", launched_entries=tuple(reverted_entries))

def print_inventory_entries(entries: Sequence[InventoryEntry]) -> None:
    for entry in entries:
        print(f"{entry.item_code}\t{format_count_text(entry.count_quantity)}")
    return None


def parse_config() -> LaunchConfig:
    parser = argparse.ArgumentParser(description="Automação de lançamento de inventário")
    parser.add_argument("--planilha", required=True, type=Path)
    parser.add_argument("--linha-inicial", default=3, type=int)
    parser.add_argument("--espera-inicial", default=10.0, type=float)
    parser.add_argument("--pausa-tecla", default=0.2, type=float)
    parser.add_argument("--simular", action="store_true")
    args = parser.parse_args()
    return LaunchConfig(
        sheet_path=Path(args.planilha),
        start_row=int(args.linha_inicial),
        initial_wait_seconds=float(args.espera_inicial),
        key_pause_seconds=float(args.pausa_tecla),
        simulate=bool(args.simular),
    )


def main() -> int:
    config = parse_config()
    entries = read_inventory_entries(config.sheet_path, config.start_row)
    if not entries:
        print("Nenhum item para lançar.")
        return 0
    if config.simulate:
        print_inventory_entries(entries)
        return 0
    writer = PyAutoGuiWriter(config.key_pause_seconds, pyautogui)
    print("Primeiro item deve estar lançado manualmente. Foque o ERP; início automático em breve.")
    launch_inventory_entries(writer, entries, config.initial_wait_seconds)
    print(f"{len(entries)} itens lançados.")
    return 0

def format_result(result: LaunchResult, labels: tuple[str, str, str, str]) -> str:
    done_text, stopped_text, none_text, count_text = labels
    status_text = done_text if result.status == "concluido" else stopped_text
    if not result.launched_entries:
        return f"{status_text}.\n{none_text}"
    last_entry = result.launched_entries[-1]
    return (
        f"{status_text}.\n"
        f"{count_text}: {len(result.launched_entries)}\n"
        f"Último item: {last_entry.item_code} | {format_count_text(last_entry.count_quantity)}"
    )


def format_launch_result(result: LaunchResult) -> str:
    return format_result(
        result,
        ("Lançamento concluído", "Lançamento parado", "Nenhuma linha lançada.", "Linhas lançadas"),
    )


def format_revert_result(result: LaunchResult) -> str:
    return format_result(
        result,
        ("Reversão concluída", "Reversão parada", "Nenhuma linha revertida.", "Linhas revertidas"),
    )

if __name__ == "__main__":
    raise SystemExit(main())
