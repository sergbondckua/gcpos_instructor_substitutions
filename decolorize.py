"""
Скрипт знебарвлює всі комірки Excel-таблиці,
крім тих, що помічені збереженими кольорами (жовтий, синій).

Використання:
    python decolorize_except_yellow.py input.xlsx output.xlsx
"""
import logging

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

logger = logging.getLogger()

class ColorFilter:
    """Зберігає набір кольорів, які потрібно залишити без змін."""

    def __init__(self, colors: set[str]):
        self.colors = {c.upper() for c in colors}

    def should_keep(self, cell) -> bool:
        """Повертає True, якщо заливка комірки входить до збережених кольорів."""
        fill = cell.fill
        if fill is None or fill.fill_type in (None, "none"):
            return False
        color = fill.fgColor
        if color is None or color.type != "rgb":
            return False

        rgb = color.rgb.upper()
        rgb_short = rgb.lstrip("FF") if len(rgb) == 8 else rgb
        return rgb in self.colors or rgb_short in self.colors


class ExcelDecolorizer:
    """Відповідає за завантаження, обробку та збереження Excel-файлу."""

    NO_FILL = PatternFill(fill_type=None)

    def __init__(
        self, input_path: str, output_path: str, color_filter: ColorFilter
    ):
        self.input_path = input_path
        self.output_path = output_path
        self.color_filter = color_filter
        self.total_cleared = 0
        self.total_kept = 0

    def process(self) -> None:
        """Основний метод: завантажує файл, обробляє і зберігає."""
        wb = load_workbook(self.input_path)
        for sheet in wb.worksheets:
            self._process_sheet(sheet)
        wb.save(self.output_path)
        self._print_report()

    def _process_sheet(self, sheet) -> None:
        """Обробляє один аркуш."""
        for row in sheet.iter_rows():
            for cell in row:
                self._process_cell(cell)

    def _process_cell(self, cell) -> None:
        """Знебарвлює комірку, якщо її колір не збережений."""
        if self.color_filter.should_keep(cell):
            self.total_kept += 1
        else:
            cell.fill = self.NO_FILL
            self.total_cleared += 1

    def _print_report(self) -> None:
        logger.info("Готово! Збережено: %s", self.output_path)
        logger.info("Знебарвлено комірок: %s", self.total_cleared)
        logger.info("Збережено кольорових: %s", self.total_kept)
