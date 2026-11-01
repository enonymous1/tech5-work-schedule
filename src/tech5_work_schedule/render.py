"""Render a validated Config into a TECH-5 Excel workbook.

Layout (v2.0.0, changed from the original):

- Row 1: project name. Row 2: RFP reference and duration. The original
  rendered neither into the workbook; the project name survived only in
  a filename truncated to 20 characters.
- Rows 3-4: the two-row header (N° / Activity / Deliverable, period
  numbers, Total), as before.
- Data from row 5, then a blank row, the legend, and the three World
  Bank guidance notes, as before.

Markers follow the legend (see ``markers.marker_for``): phases are █,
activities are X, and deliverables (and any milestone) are a red ♦.
The original gave a Deliverable without ``is_milestone`` an X,
contradicting its own legend, and bolded any description containing
the word "workshop"; both behaviours are gone. The workbook is built
from scratch with openpyxl; the original's blank one-cell "template"
carried no formatting and is not shipped or read.

Styling comes from a ``Theme`` and row/column numbers from a ``Layout``,
so this module holds no style constants of its own. File output lives
in ``io``; ``generator`` re-exports both for backwards compatibility.
"""

from __future__ import annotations

import math

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .config import Config
from .layout import (
    HEADER_ROW_1,
    HEADER_ROW_2,
    SUBTITLE_ROW,
    TITLE_ROW,
    Layout,
)
from .markers import ACTIVITY_MARKER, MILESTONE_MARKER, PHASE_MARKER, marker_for
from .theme import DEFAULT_THEME, Theme

SHEET_TITLE = "TECH-5"

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def _font(
    theme: Theme,
    *,
    bold: bool = False,
    color: str | None = None,
    size: int | None = None,
) -> Font:
    return Font(
        name=theme.font_name, size=size or theme.font_size, bold=bold, color=color
    )


def _render_header(ws: Worksheet, config: Config, layout: Layout, theme: Theme) -> None:
    """Render the title, subtitle, and header rows."""
    total_col = layout.total_col
    total_duration = config.meta.total_duration

    title = ws.cell(row=TITLE_ROW, column=1, value=config.meta.project_name)
    title.font = _font(theme, bold=True, size=theme.title_size)
    ws.merge_cells(
        start_row=TITLE_ROW, start_column=1, end_row=TITLE_ROW, end_column=total_col
    )

    subtitle_parts = [f"Duration: {total_duration} {config.meta.time_unit_label}"]
    if config.meta.rfp_reference:
        subtitle_parts.insert(0, f"RFP Reference: {config.meta.rfp_reference}")
    subtitle = ws.cell(row=SUBTITLE_ROW, column=1, value="  |  ".join(subtitle_parts))
    subtitle.font = _font(theme)
    ws.merge_cells(
        start_row=SUBTITLE_ROW,
        start_column=1,
        end_row=SUBTITLE_ROW,
        end_column=total_col,
    )

    header_fill = theme.fill(theme.header_fill)
    for row in (HEADER_ROW_1, HEADER_ROW_2):
        for col in range(1, total_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = header_fill
            cell.font = _font(theme)

    ws.merge_cells(
        start_row=HEADER_ROW_1, start_column=1, end_row=HEADER_ROW_2, end_column=1
    )
    num = ws.cell(row=HEADER_ROW_1, column=1, value="N°")
    num.alignment = Alignment(horizontal="center", vertical="center")
    num.font = _font(theme, bold=True)

    ws.merge_cells(
        start_row=HEADER_ROW_1, start_column=2, end_row=HEADER_ROW_2, end_column=2
    )
    activity = ws.cell(row=HEADER_ROW_1, column=2, value="Activity / Deliverable")
    activity.alignment = Alignment(horizontal="center", vertical="center")
    activity.font = _font(theme, bold=True)

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
    months_title.font = _font(theme, bold=True)

    for i in range(1, total_duration + 1):
        cell = ws.cell(row=HEADER_ROW_2, column=2 + i, value=i)
        cell.alignment = Alignment(horizontal="center")
        cell.font = _font(theme, bold=True)

    total_cell = ws.cell(row=HEADER_ROW_2, column=total_col, value="Total")
    total_cell.alignment = Alignment(horizontal="center", wrap_text=False)
    total_cell.font = _font(theme, bold=True)


def _render_items(ws: Worksheet, config: Config, layout: Layout, theme: Theme) -> int:
    """Render schedule rows. Returns the last data row."""
    total_col = layout.total_col
    phase_fill = theme.fill(theme.phase_fill)
    deliverable_fill = theme.fill(theme.deliverable_fill)
    for index, item in enumerate(config.items):
        row = layout.data_row(index)
        id_cell = ws.cell(row=row, column=1, value=item.id)
        desc_cell = ws.cell(row=row, column=2, value=item.description)
        id_cell.font = _font(theme)
        # Wrapping (v2.1.0) replaces the old column-width cap: a long
        # description now wraps onto extra lines instead of being cut
        # off at the column edge, and the row height grows to fit it.
        desc_cell.font = _font(theme)
        desc_cell.alignment = Alignment(wrap_text=True, vertical="center")

        if item.type == "Phase":
            id_cell.font = _font(theme, bold=True)
            desc_cell.font = _font(theme, bold=True)
            for col in range(1, total_col + 1):
                ws.cell(row=row, column=col).fill = phase_fill
        elif item.type == "Deliverable":
            id_cell.font = _font(theme, bold=True)
            desc_cell.font = _font(theme, bold=True)
            for col in range(1, total_col + 1):
                ws.cell(row=row, column=col).fill = deliverable_fill
        elif item.type == "Sub-Activity":
            desc_cell.alignment = Alignment(indent=1, wrap_text=True, vertical="center")

        if item.total_text:
            total_cell = ws.cell(row=row, column=total_col, value=item.total_text)
            total_cell.alignment = Alignment(horizontal="center", wrap_text=False)
            total_cell.font = (
                _font(theme, bold=True)
                if item.type in ("Phase", "Deliverable")
                else _font(theme)
            )

        marker = marker_for(item.type, item.is_milestone)
        marker_font = _font(theme, bold=marker.bold, color=marker.color)
        for period in item.active_periods:
            cell = ws.cell(
                row=row, column=layout.month_col + period - 1, value=marker.char
            )
            cell.font = marker_font
            cell.alignment = Alignment(horizontal="center")
    return layout.last_data_row(len(config.items))


def _render_legend_and_notes(
    ws: Worksheet, layout: Layout, theme: Theme, last_data_row: int
) -> list[tuple[int, str]]:
    """Render the legend and guidance notes. Returns (row, plain text) per note.

    Heights are not set here: they depend on the final column widths,
    which are only known after sizing. ``_apply_sizing_and_print``
    estimates the wrapped line count for each returned note and sets
    its row height, so a note on a narrow schedule gets a taller row
    instead of being clipped at the old fixed heights.
    """
    row = last_data_row + 2
    legend_label = ws.cell(row=row, column=2, value="Legend:")
    legend_label.font = _font(theme, bold=True)

    legend_items = [
        (PHASE_MARKER, "  Phase Duration", None),
        (ACTIVITY_MARKER, "  Activity Duration", None),
        (
            MILESTONE_MARKER,
            "  Deliverable / Benchmark / Approval",
            _font(theme, bold=True, color=theme.milestone_color),
        ),
    ]
    for symbol, text, custom_font in legend_items:
        row += 1
        sym_cell = ws.cell(row=row, column=1, value=symbol)
        sym_cell.alignment = Alignment(horizontal="right")
        sym_cell.font = custom_font if custom_font is not None else _font(theme)
        text_cell = ws.cell(row=row, column=2, value=text)
        text_cell.alignment = Alignment(horizontal="left")
        text_cell.font = _font(theme)

    row += 2
    note2 = CellRichText(
        TextBlock(
            InlineFont(rFont=theme.font_name, sz=theme.font_size),
            "2 Duration of activities shall be indicated ",
        ),
        TextBlock(
            InlineFont(rFont=theme.font_name, sz=theme.font_size, u="single"),
            "in a form of a bar chart.",
        ),
    )
    notes: list[tuple[str | CellRichText, str]] = [
        (
            "1 List the deliverables with the breakdown for activities required "
            "to produce them and other benchmarks such as the Client’s "
            "approvals. For phased assignments, indicate the activities, "
            "delivery of reports, and benchmarks separately for each phase.",
            "1 List the deliverables with the breakdown for activities required "
            "to produce them and other benchmarks such as the Client’s "
            "approvals. For phased assignments, indicate the activities, "
            "delivery of reports, and benchmarks separately for each phase.",
        ),
        (
            note2,
            "2 Duration of activities shall be indicated in a form of a bar chart.",
        ),
        (
            "3. Include a legend, if necessary, to help read the chart.",
            "3. Include a legend, if necessary, to help read the chart.",
        ),
    ]
    note_rows: list[tuple[int, str]] = []
    for content, plain_text in notes:
        ws.merge_cells(
            start_row=row, start_column=2, end_row=row, end_column=layout.total_col
        )
        cell = ws.cell(row=row, column=2, value=content)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if isinstance(content, str):
            cell.font = _font(theme)
        note_rows.append((row, plain_text))
        row += 1
    return note_rows


def _estimated_lines(text: str, width_chars: float) -> int:
    """Rough wrapped line count for ``text`` in a column ``width_chars`` wide.

    This is an estimate, not a renderer: openpyxl cannot know the
    viewer's font metrics, so it errs on the tall side (a full line per
    width, rounded up, minimum one). A slightly tall row is harmless;
    a clipped note is not.
    """
    if width_chars <= 0:
        return 1
    return max(1, math.ceil(len(text) / width_chars))


def _apply_sizing_and_print(
    ws: Worksheet,
    config: Config,
    layout: Layout,
    last_data_row: int,
    theme: Theme,
    note_rows: list[tuple[int, str]],
) -> None:
    ws.column_dimensions["A"].width = theme.id_col_width

    max_desc_width = theme.desc_col_min_width
    max_total_width = theme.total_col_min_width
    for r in range(HEADER_ROW_1, last_data_row + 1):
        desc_val = ws.cell(row=r, column=2).value
        if desc_val:
            max_desc_width = max(max_desc_width, len(str(desc_val)) + 2)
        total_val = ws.cell(row=r, column=layout.total_col).value
        if total_val:
            max_total_width = max(max_total_width, len(str(total_val)) + 6)
    desc_width = min(max_desc_width, theme.desc_col_max_width)
    ws.column_dimensions["B"].width = desc_width

    for i in range(1, config.meta.total_duration + 1):
        ws.column_dimensions[get_column_letter(2 + i)].width = theme.month_col_width
    ws.column_dimensions[get_column_letter(layout.total_col)].width = max_total_width

    # Row heights, now that widths are known: wrapped descriptions and
    # notes grow their rows instead of being clipped. One line is
    # ~15 points at the default font size.
    for index, item in enumerate(config.items):
        lines = _estimated_lines(item.description, desc_width)
        if lines > 1:
            ws.row_dimensions[layout.data_row(index)].height = lines * 15
    merged_note_width = (
        desc_width
        + theme.month_col_width * config.meta.total_duration
        + max_total_width
    )
    for row, plain_text in note_rows:
        ws.row_dimensions[row].height = max(
            15, _estimated_lines(plain_text, merged_note_width) * 15
        )

    ws.freeze_panes = layout.freeze_panes
    page_setup_props = ws.sheet_properties.pageSetUpPr
    if page_setup_props is not None:
        page_setup_props.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = layout.print_title_rows


def build_workbook(config: Config, theme: Theme = DEFAULT_THEME) -> Workbook:
    """Build a TECH-5 workbook from a validated config.

    ``theme`` defaults to the v2.0.0 look; pass a custom Theme to change
    colours, fonts, or widths without touching rendering code.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    assert isinstance(ws, Worksheet)
    ws.title = SHEET_TITLE
    wb.properties.title = f"TECH-5 Work Schedule: {config.meta.project_name}"

    layout = Layout(config.meta.total_duration)
    _render_header(ws, config, layout, theme)
    last_data_row = _render_items(ws, config, layout, theme)

    for r in range(HEADER_ROW_1, last_data_row + 1):
        for c in range(1, layout.total_col + 1):
            ws.cell(row=r, column=c).border = THIN_BORDER

    note_rows = _render_legend_and_notes(ws, layout, theme, last_data_row)
    _apply_sizing_and_print(ws, config, layout, last_data_row, theme, note_rows)
    return wb
