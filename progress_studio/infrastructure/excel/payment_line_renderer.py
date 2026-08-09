from __future__ import annotations

import shutil
import tempfile
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Border, Side

from progress_studio.config.payment_theme import PAYMENT_LINE_COLORS, PAYMENT_LINE_STYLE
from progress_studio.domain.payment_models import PaymentLineRenderResult, PaymentResolvedPeriod, PaymentResolvedPoint
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError, PaymentWorkbookSnapshotter


class PaymentLineRenderer:
    """Render one sparse Payment period as a cell-border staircase.

    The renderer intentionally knows nothing about payment percentages or Plan
    distributions.  It receives already-resolved row/column boundaries and only
    paints cell borders.  No Shape objects, pixel coordinates, or drawing anchors
    are used, so the line remains attached to the worksheet grid while zooming.
    """

    MAIN_SHEET = "main"
    PAYMENT_SHEET = "Payment"

    def render_single_period(
        self,
        source_workbook: Path,
        output_workbook: Path,
        period: PaymentResolvedPeriod,
    ) -> PaymentLineRenderResult:
        source = Path(source_workbook)
        output = Path(output_workbook)
        if not period.points:
            raise PaymentWorkbookError(f"{period.period_id} has no resolved Payment points to render.")

        color = PAYMENT_LINE_COLORS.get(period.period_id, "C00000")
        output.parent.mkdir(parents=True, exist_ok=True)
        keep_vba = source.suffix.lower() == ".xlsm"

        with tempfile.NamedTemporaryFile(
            prefix="payment_line_", suffix=output.suffix, dir=output.parent, delete=False
        ) as handle:
            temp_path = Path(handle.name)

        try:
            wb = load_workbook(source, keep_vba=keep_vba)
            try:
                if self.MAIN_SHEET not in wb.sheetnames:
                    raise PaymentWorkbookError("Worksheet 'main' was not found.")
                if self.PAYMENT_SHEET in wb.sheetnames:
                    wb.remove(wb[self.PAYMENT_SHEET])

                main = wb[self.MAIN_SHEET]
                payment = wb.copy_worksheet(main)
                payment.title = self.PAYMENT_SHEET
                payment.freeze_panes = main.freeze_panes
                payment.sheet_view.showGridLines = main.sheet_view.showGridLines
                payment.auto_filter.ref = main.auto_filter.ref

                line = Side(style=PAYMENT_LINE_STYLE, color=color)
                points = tuple(sorted(period.points, key=lambda p: p.activity_row))
                self._paint_points_and_segments(payment, points, line)
                wb.save(temp_path)
            finally:
                wb.close()
            shutil.move(str(temp_path), str(output))
        except PaymentWorkbookError:
            raise
        except Exception as exc:
            raise PaymentWorkbookError(f"Payment line could not be rendered: {exc}") from exc
        finally:
            temp_path.unlink(missing_ok=True)

        return PaymentLineRenderResult(
            source_workbook=source,
            output_workbook=output,
            payment_sheet=self.PAYMENT_SHEET,
            period_id=period.period_id,
            rendered_points=len(period.points),
            color=color,
        )

    def _paint_points_and_segments(self, ws, points: tuple[PaymentResolvedPoint, ...], line: Side) -> None:
        # A single point is still visible as one vertical boundary on its activity row.
        for point in points:
            self._vertical_boundary(ws, point.activity_row, point.activity_row, self._boundary(point), line)

        # Staircase contract: from each target, walk horizontally on that activity
        # row to the next target's boundary, then vertically down to the next row.
        # All geometry is expressed only in worksheet row/column indices.
        for current, nxt in zip(points, points[1:]):
            b1 = self._boundary(current)
            b2 = self._boundary(nxt)
            self._horizontal_boundary(ws, current.activity_row, b1, b2, line)
            self._vertical_boundary(ws, current.activity_row, nxt.activity_row, b2, line)

    @staticmethod
    def _boundary(point: PaymentResolvedPoint) -> int:
        # Boundary N means the grid line immediately before column N.  Therefore
        # the right edge of column C is the same boundary as the left edge of C+1.
        return point.timescale_column if point.boundary_edge == "left" else point.timescale_column + 1

    @staticmethod
    def _replace_border(cell, *, left=None, right=None, top=None, bottom=None) -> None:
        old = copy(cell.border)
        cell.border = Border(
            left=left if left is not None else old.left,
            right=right if right is not None else old.right,
            top=top if top is not None else old.top,
            bottom=bottom if bottom is not None else old.bottom,
            diagonal=old.diagonal,
            diagonal_direction=old.diagonal_direction,
            diagonalUp=old.diagonalUp,
            diagonalDown=old.diagonalDown,
            outline=old.outline,
            vertical=old.vertical,
            horizontal=old.horizontal,
        )

    def _vertical_boundary(self, ws, row1: int, row2: int, boundary: int, line: Side) -> None:
        start, end = sorted((row1, row2))
        # Prefer the left border of the cell to the right of the boundary.  At the
        # worksheet's far-right edge, fall back to the right border of the last cell.
        if boundary <= ws.max_column:
            for row in range(start, end + 1):
                self._replace_border(ws.cell(row, boundary), left=line)
        else:
            col = max(boundary - 1, 1)
            for row in range(start, end + 1):
                self._replace_border(ws.cell(row, col), right=line)

    def _horizontal_boundary(self, ws, row: int, boundary1: int, boundary2: int, line: Side) -> None:
        if boundary1 == boundary2:
            return
        left = min(boundary1, boundary2)
        right = max(boundary1, boundary2)
        # Bottom borders create a crisp horizontal step without introducing any
        # drawing object. Boundaries [left, right] span cells left..right-1.
        for col in range(left, right):
            if 1 <= col <= ws.max_column:
                self._replace_border(ws.cell(row, col), bottom=line)
