from __future__ import annotations

import shutil
import tempfile
from copy import copy
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Border, Side

from progress_studio.config.payment_theme import PAYMENT_LINE_COLORS, PAYMENT_LINE_STYLE
from progress_studio.domain.payment_models import (
    PaymentLineRenderResult,
    PaymentMultiLineRenderResult,
    PaymentResolvedPeriod,
    PaymentResolvedPoint,
)
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError


class PaymentLineRenderer:
    """Cell-based vertical-backbone Payment renderer.

    Each Payment period gets one vertical backbone at its Payment Date on the
    weekly timescale. Only sparse resolved Activity points are rendered. A short
    horizontal branch connects the backbone to each resolved requirement point.
    No Shapes, drawing anchors, or pixel coordinates are created.
    """

    MAIN_SHEET = "main"
    PAYMENT_SHEET = "Payment"
    HEADER_ROW = 4

    def render_periods(
        self,
        source_workbook: Path,
        output_workbook: Path,
        periods: tuple[PaymentResolvedPeriod, ...],
    ) -> PaymentMultiLineRenderResult:
        source = Path(source_workbook)
        output = Path(output_workbook)
        active = tuple(period for period in periods if period.points)
        if not active:
            raise PaymentWorkbookError("No resolved Payment points were available to render.")

        output.parent.mkdir(parents=True, exist_ok=True)
        keep_vba = source.suffix.lower() == ".xlsm"
        with tempfile.NamedTemporaryFile(
            prefix="payment_backbone_", suffix=output.suffix, dir=output.parent, delete=False
        ) as handle:
            temp_path = Path(handle.name)

        colors: list[tuple[str, str]] = []
        rendered_points = 0
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

                timeline = self._timescale_boundaries(payment)
                if not timeline:
                    raise PaymentWorkbookError("Weekly timescale dates were not found on the main sheet.")

                used_boundaries: set[int] = set()
                min_boundary = timeline[0][0]
                max_boundary = timeline[-1][0] + 1
                for period in active:
                    color = PAYMENT_LINE_COLORS.get(period.period_id, "7F7F7F")
                    line = Side(style=PAYMENT_LINE_STYLE, color=color)
                    endpoint = Side(style="thick", color=color)
                    preferred = self._payment_boundary(period.payment_date, timeline)
                    backbone = self._allocate_backbone_boundary(
                        preferred, used_boundaries, min_boundary, max_boundary
                    )
                    used_boundaries.add(backbone)
                    self._paint_header_marker(payment, period, backbone, line)
                    self._paint_vertical_backbone(payment, period, backbone, line, endpoint)
                    colors.append((period.period_id, color))
                    rendered_points += len(period.points)

                wb.save(temp_path)
            finally:
                wb.close()
            shutil.move(str(temp_path), str(output))
        except PaymentWorkbookError:
            raise
        except Exception as exc:
            raise PaymentWorkbookError(f"Payment lines could not be rendered: {exc}") from exc
        finally:
            temp_path.unlink(missing_ok=True)

        return PaymentMultiLineRenderResult(
            source_workbook=source,
            output_workbook=output,
            payment_sheet=self.PAYMENT_SHEET,
            period_ids=tuple(period.period_id for period in active),
            rendered_points=rendered_points,
            rendered_periods=len(active),
            colors=tuple(colors),
        )

    def render_single_period(
        self,
        source_workbook: Path,
        output_workbook: Path,
        period: PaymentResolvedPeriod,
    ) -> PaymentLineRenderResult:
        """Compatibility wrapper: single period rendered with the backbone engine."""
        result = self.render_periods(source_workbook, output_workbook, (period,))
        color = result.colors[0][1]
        return PaymentLineRenderResult(
            source_workbook=result.source_workbook,
            output_workbook=result.output_workbook,
            payment_sheet=result.payment_sheet,
            period_id=period.period_id,
            rendered_points=result.rendered_points,
            color=color,
        )

    def _paint_vertical_backbone(
        self,
        ws,
        period: PaymentResolvedPeriod,
        backbone_boundary: int,
        line: Side,
        endpoint: Side,
    ) -> None:
        points = tuple(sorted(period.points, key=lambda p: p.activity_row))
        if not points:
            return
        first_row = points[0].activity_row
        last_row = points[-1].activity_row
        self._vertical_boundary(ws, first_row, last_row, backbone_boundary, line)

        for point in points:
            target_boundary = self._boundary(point)
            self._horizontal_boundary(ws, point.activity_row, backbone_boundary, target_boundary, line)
            # A stronger one-row cap marks the actual resolved % target without
            # writing into the weekly Plan cells.
            self._vertical_boundary(ws, point.activity_row, point.activity_row, target_boundary, endpoint)

    def _paint_header_marker(
        self,
        ws,
        period: PaymentResolvedPeriod,
        boundary: int,
        line: Side,
    ) -> None:
        """Carry the backbone through the timescale header and attach a lightweight note."""
        self._vertical_boundary(ws, 1, self.HEADER_ROW, boundary, line)
        marker_col = boundary if boundary <= ws.max_column else ws.max_column
        marker_cell = ws.cell(self.HEADER_ROW, marker_col)
        date_text = period.payment_date.isoformat() if period.payment_date else "no date"
        marker_cell.comment = Comment(
            f"{period.period_id} Payment backbone\nPayment date: {date_text}",
            "Progress Studio",
        )

    @staticmethod
    def _allocate_backbone_boundary(
        preferred: int,
        used: set[int],
        minimum: int,
        maximum: int,
    ) -> int:
        """Keep coincident payments readable using the nearest free cell boundary.

        First choice is the resolved Payment-Date boundary. For collisions, use
        the opposite edge of that weekly cell, then the nearest neighbouring edge.
        """
        preferred = min(max(preferred, minimum), maximum)
        if preferred not in used:
            return preferred
        distance = 1
        while distance <= (maximum - minimum + 1):
            for candidate in (preferred + distance, preferred - distance):
                if minimum <= candidate <= maximum and candidate not in used:
                    return candidate
            distance += 1
        return preferred

    def _timescale_boundaries(self, ws) -> tuple[tuple[int, date], ...]:
        result: list[tuple[int, date]] = []
        for col in range(1, ws.max_column + 1):
            value = ws.cell(self.HEADER_ROW, col).value
            if isinstance(value, datetime):
                result.append((col, value.date()))
            elif isinstance(value, date):
                result.append((col, value))
        return tuple(result)

    @staticmethod
    def _payment_boundary(payment_date: date | None, timeline: tuple[tuple[int, date], ...]) -> int:
        # Boundary before the first weekly bucket whose date is on/after Payment Date.
        # If the Payment Date is beyond the last bucket, use the last bucket's right edge.
        if payment_date is None:
            return timeline[0][0]
        for col, bucket_date in timeline:
            if bucket_date >= payment_date:
                return col
        return timeline[-1][0] + 1

    @staticmethod
    def _boundary(point: PaymentResolvedPoint) -> int:
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
        for col in range(left, right):
            if 1 <= col <= ws.max_column:
                self._replace_border(ws.cell(row, col), bottom=line)
