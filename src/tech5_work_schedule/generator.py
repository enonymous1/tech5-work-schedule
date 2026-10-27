"""Render a validated Config into a TECH-5 Excel workbook.

Layout (v2.0.0, changed from the original):

- Row 1: project name. Row 2: RFP reference and duration. The original
  rendered neither into the workbook; the project name survived only in
  a filename truncated to 20 characters.
- Rows 3-4: the two-row header (N° / Activity / Deliverable, period
  numbers, Total), as before.
- Data from row 5, then a blank row, the legend, and the three World
  Bank guidance notes, as before.

Markers follow the legend: phases are █, activities are X, and
deliverables (and any milestone) are a red ♦. The original gave a
Deliverable without ``is_milestone`` an X, contradicting its own legend,
and bolded any description containing the word "workshop"; both
behaviours are gone. The workbook is built from scratch with openpyxl;
the original's blank one-cell "template" carried no formatting and is
not shipped or read.
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .config import Config, load_config

FONT_NAME = "Gill Sans"
FONT_SIZE = 10
TITLE_ROW = 1
SUBTITLE_ROW = 2
HEADER_ROW_1 = 3
HEADER_ROW_2 = 4
FIRST_DATA_ROW = 5
SHEET_TITLE = "TECH-5"

HEADER_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
PHASE_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
DELIVERABLE_FILL = PatternFill(
    start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
)
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

PHASE_MARKER = "█"
ACTIVITY_MARKER = "X"
MILESTONE_MARKER = "♦"


def _font(
    *, bold: bool = False, color: str | None = None, size: int = FONT_SIZE
) -> Font:
    return Font(name=FONT_NAME, size=size, bold=bold, color=color)


def _marker_for(item_type: str, is_milestone: bool) -> tuple[str, Font]:
    if is_milestone or item_type == "Deliverable":
        return MILESTONE_MARKER, _font(bold=True, color="FF0000")
    if item_type == "Phase":
        return PHASE_MARKER, _font()
    return ACTIVITY_MARKER, _font()


def _render_header(ws: Worksheet, config: Config) -> int:
    """Render the title, subtitle, and header rows. Returns the total column."""
    total_duration = config.meta.total_duration
    total_col = 3 + total_duration

    title = ws.cell(row=TITLE_ROW, column=1, value=config.meta.project_name)
    title.font = _font(bold=True, size=12)
    ws.merge_cells(
        start_row=TITLE_ROW, start_column=1, end_row=TITLE_ROW, end_column=total_col
    )

    subtitle_parts = [f"Duration: {total_duration} {config.meta.time_unit_label}"]
    if config.meta.rfp_reference:
        subtitle_parts.insert(0, f"RFP Reference: {config.meta.rfp_reference}")
    subtitle = ws.cell(row=SUBTITLE_ROW, column=1, value="  |  ".join(subtitle_parts))
    subtitle.font = _font()
    ws.merge_cells(
        start_row=SUBTITLE_ROW,
        start_column=1,
        end_row=SUBTITLE_ROW,
        end_column=total_col,
    )

    for row in (HEADER_ROW_1, HEADER_ROW_2):
        for col in range(1, total_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = HEADER_FILL
            cell.font = _font()

    ws.merge_cells(
        start_row=HEADER_ROW_1, start_column=1, end_row=HEADER_ROW_2, end_column=1
    )
    num = ws.cell(row=HEADER_ROW_1, column=1, value="N°")
    num.alignment = Alignment(horizontal="center", vertical="center")
    num.font = _font(bold=True)

    ws.merge_cells(
        start_row=HEADER_ROW_1, start_column=2, end_row=HEADER_ROW_2, end_column=2
    )
    activity = ws.cell(row=HEADER_ROW_1, column=2, value="Activity / Deliverable")
    activity.alignment = Alignment(horizontal="center", vertical="center")
    activity.font = _font(bold=True)

    ws.merge_cells(
        start_row=HEADER_ROW_1,
        start_column=3,
        end_row=HEADER_ROW_1,
        end_column=total_col,
    )
    months_title = ws.cell(
        row=HEADER_ROW_1, column=3, value=config.meta.time_unit_label
    )
    months_title.alignment = Alignment(horizontal="center", vertical="center")
    months_title.font = _font(bold=True)

    for i in range(1, total_duration + 1):
        cell = ws.cell(row=HEADER_ROW_2, column=2 + i, value=i)
        cell.alignment = Alignment(horizontal="center")
        cell.font = _font(bold=True)

    total_cell = ws.cell(row=HEADER_ROW_2, column=total_col, value="Total")
    total_cell.alignment = Alignment(horizontal="center", wrap_text=False)
    total_cell.font = _font(bold=True)
    return total_col


def _render_items(ws: Worksheet, config: Config, total_col: int) -> int:
    """Render schedule rows. Returns the last data row."""
    row = FIRST_DATA_ROW
    start_col = 3
    for item in config.items:
        id_cell = ws.cell(row=row, column=1, value=item.id)
        desc_cell = ws.cell(row=row, column=2, value=item.description)
        id_cell.font = _font()
        desc_cell.font = _font()

        if item.type == "Phase":
            id_cell.font = _font(bold=True)
            desc_cell.font = _font(bold=True)
            for col in range(1, total_col + 1):
                ws.cell(row=row, column=col).fill = PHASE_FILL
        elif item.type == "Deliverable":
            id_cell.font = _font(bold=True)
            desc_cell.font = _font(bold=True)
            for col in range(1, total_col + 1):
                ws.cell(row=row, column=col).fill = DELIVERABLE_FILL
        elif item.type == "Sub-Activity":
            desc_cell.alignment = Alignment(indent=1)

        if item.total_text:
            total_cell = ws.cell(row=row, column=total_col, value=item.total_text)
            total_cell.alignment = Alignment(horizontal="center", wrap_text=False)
            total_cell.font = (
                _font(bold=True) if item.type in ("Phase", "Deliverable") else _font()
            )

        marker, marker_font = _marker_for(item.type, item.is_milestone)
        for period in item.active_periods:
            cell = ws.cell(row=row, column=start_col + period - 1, value=marker)
            cell.font = marker_font
            cell.alignment = Alignment(horizontal="center")
        row += 1
    return row - 1


def _render_legend_and_notes(ws: Worksheet, total_col: int, last_data_row: int) -> None:
    row = last_data_row + 2
    legend_label = ws.cell(row=row, column=2, value="Legend:")
    legend_label.font = _font(bold=True)

    legend_items = [
        (PHASE_MARKER, "  Phase Duration", None),
        (ACTIVITY_MARKER, "  Activity Duration", None),
        (
            MILESTONE_MARKER,
            "  Deliverable / Benchmark / Approval",
            _font(bold=True, color="FF0000"),
        ),
    ]
    for symbol, text, custom_font in legend_items:
        row += 1
        sym_cell = ws.cell(row=row, column=1, value=symbol)
        sym_cell.alignment = Alignment(horizontal="right")
        sym_cell.font = custom_font if custom_font is not None else _font()
        text_cell = ws.cell(row=row, column=2, value=text)
        text_cell.alignment = Alignment(horizontal="left")
        text_cell.font = _font()

    row += 2
    note2 = CellRichText(
        TextBlock(
            InlineFont(rFont=FONT_NAME, sz=FONT_SIZE),
            "2 Duration of activities shall be indicated ",
        ),
        TextBlock(
            InlineFont(rFont=FONT_NAME, sz=FONT_SIZE, u="single"),
            "in a form of a bar chart.",
        ),
        TextBlock(InlineFont(rFont=FONT_NAME, sz=FONT_SIZE), "."),
    )
    notes: list[tuple[str | CellRichText, int]] = [
        (
            "1 List the deliverables with the breakdown for activities required "
            "to produce them and other benchmarks such as the Client’s "
            "approvals. For phased assignments, indicate the activities, "
            "delivery of reports, and benchmarks separately for each phase.",
            45,
        ),
        (note2, 25),
        ("3. Include a legend, if necessary, to help read the chart.", 20),
    ]
    for content, height in notes:
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=total_col)
        cell = ws.cell(row=row, column=2, value=content)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if isinstance(content, str):
            cell.font = _font()
        ws.row_dimensions[row].height = height
        row += 1


def _apply_sizing_and_print(
    ws: Worksheet, config: Config, total_col: int, last_data_row: int
) -> None:
    ws.column_dimensions["A"].width = 8

    max_desc_width = 40
    max_total_width = 10
    for r in range(HEADER_ROW_1, last_data_row + 1):
        desc_val = ws.cell(row=r, column=2).value
        if desc_val:
            max_desc_width = max(max_desc_width, len(str(desc_val)) + 2)
        total_val = ws.cell(row=r, column=total_col).value
        if total_val:
            max_total_width = max(max_total_width, len(str(total_val)) + 6)
    ws.column_dimensions["B"].width = min(max_desc_width, 120)

    for i in range(1, config.meta.total_duration + 1):
        ws.column_dimensions[get_column_letter(2 + i)].width = 3.5
    ws.column_dimensions[get_column_letter(total_col)].width = max_total_width

    ws.freeze_panes = "C5"
    page_setup_props = ws.sheet_properties.pageSetUpPr
    if page_setup_props is not None:
        page_setup_props.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{TITLE_ROW}:{HEADER_ROW_2}"


def build_workbook(config: Config) -> Workbook:
    """Build a TECH-5 workbook from a validated config."""
    wb = openpyxl.Workbook()
    ws = wb.active
    assert isinstance(ws, Worksheet)
    ws.title = SHEET_TITLE
    wb.properties.title = f"TECH-5 Work Schedule: {config.meta.project_name}"

    total_col = _render_header(ws, config)
    last_data_row = _render_items(ws, config, total_col)

    for r in range(HEADER_ROW_1, last_data_row + 1):
        for c in range(1, total_col + 1):
            ws.cell(row=r, column=c).border = THIN_BORDER

    _render_legend_and_notes(ws, total_col, last_data_row)
    _apply_sizing_and_print(ws, config, total_col, last_data_row)
    return wb


def slugify(project_name: str) -> str:
    """Make a project name safe for a filename, without truncating it to 20 chars."""
    slug = re.sub(r"[^A-Za-z0-9]+", "_", project_name).strip("_")
    return slug[:60] or "Output"


def generate(config_path: str | Path, output: str | Path | None = None) -> Path:
    """Generate a workbook from a config file and return its path.

    With no ``output``, the workbook goes to ``output/`` under the current
    working directory, named ``TECH5_<project>_<timestamp>.xlsx``. With an
    explicit ``output``, that exact path is used and parent directories
    are created.
    """
    config = load_config(config_path)
    wb = build_workbook(config)
    if output is None:
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = (
            output_dir / f"TECH5_{slugify(config.meta.project_name)}_{timestamp}.xlsx"
        )
    else:
        output_path = Path(output)
        if output_path.parent != Path(""):
            output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
