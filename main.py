import logging
import sys
from datetime import datetime, timedelta

from decolorize import ColorFilter, ExcelDecolorizer, logger

logger = logging.getLogger(__name__)

YELLOW_COLORS = {
    "FFFF00",
    "FFFFFF00",
    "FFFF99",
    "FFD700",
    "FFFACD",
    "FFF44F",
}

BLUE_COLORS = {
    "2a6099",
    "FF0000FF",
    "ADD8E6",
    "00BFFF",
    "1F75FE",
    "4472C4",
    "5B9BD5",
    "2E75B6",
    "000080",
    "6699FF",
}


if __name__ == "__main__":
    if len(sys.argv) != 2:
        logger.warning(
            "Приклад використання: python main.py input.xlsx"
        )
        sys.exit(1)

    color_filter = ColorFilter(colors=YELLOW_COLORS | BLUE_COLORS)
    decolorizer = ExcelDecolorizer(
        input_path=sys.argv[1],
        output_path=f"output_files/Розклад {(datetime.now() + timedelta(days=1)).strftime('%d.%m.%Y')}.xlsx",
        color_filter=color_filter,
    )
    decolorizer.process()
