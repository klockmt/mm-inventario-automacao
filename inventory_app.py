from __future__ import annotations
import sys

import queue
import threading
import tkinter as tk
from collections.abc import Sequence
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pyautogui

from inventory_launcher import (
    AutomationStoppedError,
    InventoryEntry,
    PyAutoGuiWriter,
    format_count_text,
    format_launch_result,
    format_revert_result,
    launch_inventory_entries,
    read_fixed_parameters,
    read_inventory_entries,
    revert_inventory_entries,
)

ENABLE_CONTROLS_MESSAGE = "__ENABLE_CONTROLS__"
RESULT_MESSAGE_PREFIX = "__RESULT__:"
ERROR_MESSAGE_PREFIX = "__ERROR__:"

def format_inventory_entries(entries: Sequence[InventoryEntry]) -> str:
    if not entries:
        return "Nenhum item encontrado"
    lines = [
        f"{entry.item_code}\t{format_count_text(entry.count_quantity)}"
        for entry in entries
    ]
    return "\n".join(lines)


def parse_int_field(raw_text: str, field_label: str) -> int:
    try:
        parsed_value = int(raw_text.strip())
    except ValueError as error:
        raise ValueError(f"{field_label} deve ser um número inteiro") from error
    if parsed_value < 1:
        raise ValueError(f"{field_label} deve ser maior ou igual a 1")
    return parsed_value


def parse_float_field(raw_text: str, field_label: str) -> float:
    normalized_text = raw_text.strip().replace(",", ".")
    try:
        parsed_value = float(normalized_text)
    except ValueError as error:
        raise ValueError(f"{field_label} deve ser um número") from error
    if parsed_value < 0.0:
        raise ValueError(f"{field_label} deve ser maior ou igual a zero")
    return parsed_value

def executable_folder() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def build_default_sheet_path(base_folder: Path) -> Path:
    return base_folder / "inventario.xlsx"

class InventoryApp:
    def __init__(self, root: tk.Tk) -> None:
        self._root = root
        self._root.title("Lançamento de Inventário")
        self._message_queue: queue.Queue[str] = queue.Queue()
        default_sheet_path = build_default_sheet_path(executable_folder())
        self._sheet_path_variable = tk.StringVar(value=str(default_sheet_path))
        self._start_row_variable = tk.StringVar(value="3")
        self._wait_variable = tk.StringVar(value="10")
        self._pause_variable = tk.StringVar(value="0.2")
        self._build_widgets()
        self._poll_message_queue()
        return None

    def run(self) -> None:
        self._root.mainloop()
        return None

    def _build_widgets(self) -> None:
        self._root.columnconfigure(0, weight=1)
        self._build_form()
        self._build_actions()
        self._build_output()
        return None

    def _build_form(self) -> None:
        form_frame = ttk.LabelFrame(self._root, text="Configuração", padding=8)
        form_frame.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        form_frame.columnconfigure(1, weight=1)
        ttk.Label(form_frame, text="Planilha").grid(row=0, column=0, sticky="w")
        self._sheet_path_entry = ttk.Entry(form_frame, textvariable=self._sheet_path_variable)
        self._sheet_path_entry.grid(row=0, column=1, sticky="ew", padx=4)
        self._select_path_button = ttk.Button(form_frame, text="Selecionar", command=self.select_sheet_path)
        self._select_path_button.grid(row=0, column=2, padx=4)
        ttk.Label(form_frame, text="Linha inicial").grid(row=1, column=0, sticky="w", pady=4)
        self._start_row_entry = ttk.Entry(form_frame, textvariable=self._start_row_variable, width=10)
        self._start_row_entry.grid(row=1, column=1, sticky="w", padx=4)
        ttk.Label(form_frame, text="Espera inicial (s)").grid(row=2, column=0, sticky="w")
        self._wait_entry = ttk.Entry(form_frame, textvariable=self._wait_variable, width=10)
        self._wait_entry.grid(row=2, column=1, sticky="w", padx=4)
        ttk.Label(form_frame, text="Pausa tecla (s)").grid(row=3, column=0, sticky="w")
        self._pause_entry = ttk.Entry(form_frame, textvariable=self._pause_variable, width=10)
        self._pause_entry.grid(row=3, column=1, sticky="w", padx=4)
        return None

    def _build_actions(self) -> None:
        actions_frame = ttk.Frame(self._root, padding=8)
        actions_frame.grid(row=1, column=0, sticky="ew", padx=8)
        self._read_test_button = ttk.Button(actions_frame, text="Testar leitura", command=self.run_read_test)
        self._read_test_button.grid(row=0, column=0, padx=4)
        self._type_test_button = ttk.Button(actions_frame, text="Testar digitação", command=self.run_type_test)
        self._type_test_button.grid(row=0, column=1, padx=4)
        self._launch_button = ttk.Button(actions_frame, text="Lançamento automático", command=self.run_launch)
        self._launch_button.grid(row=0, column=2, padx=4)
        self._revert_button = ttk.Button(actions_frame, text="Reverter lançamento", command=self.run_revert)
        self._revert_button.grid(row=0, column=3, padx=4)
        return None

    def _build_output(self) -> None:
        output_frame = ttk.LabelFrame(self._root, text="Resultado", padding=8)
        output_frame.grid(row=2, column=0, sticky="nsew", padx=8, pady=8)
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)
        self._output_text = tk.Text(output_frame, height=12, width=80)
        self._output_text.grid(row=0, column=0, sticky="nsew")
        self._root.rowconfigure(2, weight=1)
        return None

    def select_sheet_path(self) -> None:
        selected_path = filedialog.askopenfilename(filetypes=[("Planilha Excel", "*.xlsx")])
        if selected_path:
            self._sheet_path_variable.set(selected_path)
        return None

    def run_read_test(self) -> None:
        try:
            sheet_path, start_row = self._read_config_from_form()
            entries = read_inventory_entries(sheet_path, start_row)
            self._output_text.delete("1.0", tk.END)
            self._output_text.insert(tk.END, format_inventory_entries(entries))
        except ValueError as error:
            messagebox.showerror("Erro de validação", str(error))
        return None

    def run_type_test(self) -> None:
        self._start_automation(minimize_window=False)
        return None

    def run_launch(self) -> None:
        self._start_automation(minimize_window=True)
        return None

    def run_revert(self) -> None:
        self._start_revert(minimize_window=True)
        return None

    def _read_revert_parameters(self) -> tuple[Sequence[InventoryEntry], str, str, float, float]:
        sheet_path, start_row = self._read_config_from_form()
        wait_seconds = parse_float_field(self._wait_variable.get(), "Espera inicial")
        pause_seconds = parse_float_field(self._pause_variable.get(), "Pausa tecla")
        entries = read_inventory_entries(sheet_path, start_row)
        if not entries:
            raise ValueError("Nenhum item para reverter")
        branch_code, warehouse_code = read_fixed_parameters(sheet_path)
        return entries, branch_code, warehouse_code, wait_seconds, pause_seconds

    def _start_revert(self, minimize_window: bool) -> None:
        try:
            entries, branch_code, warehouse_code, wait_seconds, pause_seconds = self._read_revert_parameters()
            self._set_controls_disabled(True)
            self._append_output(f"Iniciando reversão de {len(entries)} itens.")
            automation_thread = threading.Thread(
                target=self._revert_worker,
                args=(entries, branch_code, warehouse_code, wait_seconds, pause_seconds),
                daemon=True,
            )
            automation_thread.start()
            if minimize_window:
                self._root.iconify()
        except ValueError as error:
            messagebox.showerror("Erro de validação", str(error))
        return None

    def _revert_worker(
        self,
        entries: Sequence[InventoryEntry],
        branch_code: str,
        warehouse_code: str,
        wait_seconds: float,
        pause_seconds: float,
    ) -> None:
        try:
            writer = PyAutoGuiWriter(pause_seconds, pyautogui)
            result = revert_inventory_entries(writer, entries, branch_code, warehouse_code, wait_seconds)
            self._message_queue.put(RESULT_MESSAGE_PREFIX + format_revert_result(result))
        except AutomationStoppedError as error:
            self._message_queue.put(RESULT_MESSAGE_PREFIX + f"Reversão parada. {error}")
        except Exception as error:
            self._message_queue.put(ERROR_MESSAGE_PREFIX + f"Erro durante a reversão: {error}")
        finally:
            self._message_queue.put(ENABLE_CONTROLS_MESSAGE)
        return None

    def _read_config_from_form(self) -> tuple[Path, int]:
        raw_path = self._sheet_path_variable.get().strip()
        if not raw_path:
            raise ValueError("Selecione uma planilha")
        sheet_path = Path(raw_path)
        if not sheet_path.exists():
            raise ValueError("Planilha não encontrada")
        start_row = parse_int_field(self._start_row_variable.get(), "Linha inicial")
        return sheet_path, start_row

    def _read_launch_parameters(self) -> tuple[Sequence[InventoryEntry], float, float]:
        sheet_path, start_row = self._read_config_from_form()
        wait_seconds = parse_float_field(self._wait_variable.get(), "Espera inicial")
        pause_seconds = parse_float_field(self._pause_variable.get(), "Pausa tecla")
        entries = read_inventory_entries(sheet_path, start_row)
        if not entries:
            raise ValueError("Nenhum item para lançar")
        return entries, wait_seconds, pause_seconds

    def _start_automation(self, minimize_window: bool) -> None:
        try:
            entries, wait_seconds, pause_seconds = self._read_launch_parameters()
            self._set_controls_disabled(True)
            self._append_output(f"Iniciando lançamento de {len(entries)} itens.")
            automation_thread = threading.Thread(
                target=self._launch_worker,
                args=(entries, wait_seconds, pause_seconds),
                daemon=True,
            )
            automation_thread.start()
            if minimize_window:
                self._root.iconify()
        except ValueError as error:
            messagebox.showerror("Erro de validação", str(error))
        return None

    def _launch_worker(
        self,
        entries: Sequence[InventoryEntry],
        wait_seconds: float,
        pause_seconds: float,
    ) -> None:
        try:
            writer = PyAutoGuiWriter(pause_seconds, pyautogui)
            result = launch_inventory_entries(writer, entries, wait_seconds)
            self._message_queue.put(RESULT_MESSAGE_PREFIX + format_launch_result(result))
        except AutomationStoppedError as error:
            self._message_queue.put(RESULT_MESSAGE_PREFIX + f"Lançamento parado. {error}")
        except Exception as error:
            self._message_queue.put(ERROR_MESSAGE_PREFIX + f"Erro durante o lançamento: {error}")
        finally:
            self._message_queue.put(ENABLE_CONTROLS_MESSAGE)
        return None

    def _poll_message_queue(self) -> None:
        while not self._message_queue.empty():
            message = self._message_queue.get_nowait()
            if message == ENABLE_CONTROLS_MESSAGE:
                self._set_controls_disabled(False)
            elif message.startswith(RESULT_MESSAGE_PREFIX):
                result_text = message.removeprefix(RESULT_MESSAGE_PREFIX)
                self._append_output(result_text)
                self._root.deiconify()
                messagebox.showinfo("Resultado", result_text)
            elif message.startswith(ERROR_MESSAGE_PREFIX):
                error_text = message.removeprefix(ERROR_MESSAGE_PREFIX)
                self._append_output(error_text)
                self._root.deiconify()
                messagebox.showerror("Erro", error_text)
            else:
                self._append_output(message)
        self._root.after(100, self._poll_message_queue)
        return None

    def _append_output(self, message: str) -> None:
        self._output_text.insert(tk.END, message + "\n")
        self._output_text.see(tk.END)
        return None

    def _set_controls_disabled(self, disabled: bool) -> None:
        control_state = tk.DISABLED if disabled else tk.NORMAL
        self._select_path_button.config(state=control_state)
        self._sheet_path_entry.config(state=control_state)
        self._start_row_entry.config(state=control_state)
        self._wait_entry.config(state=control_state)
        self._pause_entry.config(state=control_state)
        self._read_test_button.config(state=control_state)
        self._type_test_button.config(state=control_state)
        self._launch_button.config(state=control_state)
        self._revert_button.config(state=control_state)
        return None


def main() -> int:
    root = tk.Tk()
    application = InventoryApp(root)
    application.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())