from __future__ import annotations

from typing import Iterable

from PIL import Image
from PySide6.QtCore import QPointF, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QImage, QPainter, QPen, QPixmap
from PySide6.QtWidgets import QWidget


def pil_to_qimage(image: Image.Image) -> QImage:
    rgba = image.convert("RGBA")
    data = rgba.tobytes("raw", "RGBA")
    return QImage(data, rgba.width, rgba.height, QImage.Format.Format_RGBA8888).copy()


class ImageSelectionCanvas(QWidget):
    selection_finished = Signal(object)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setMinimumSize(520, 520)
        self.setMouseTracking(True)
        self._pixmap: QPixmap | None = None
        self._image_size = (1, 1)
        self._draw_rect = QRectF()
        self._start: QPointF | None = None
        self._current_rect: QRectF | None = None
        self._rects: list[tuple[str, tuple[float, float, float, float]]] = []
        self.selected_name: str | None = None

    def set_image(self, image: Image.Image) -> None:
        self._image_size = image.size
        self._pixmap = QPixmap.fromImage(pil_to_qimage(image))
        self._current_rect = None
        self.update()

    def set_rects(self, rects: Iterable[tuple[str, tuple[float, float, float, float]]]) -> None:
        self._rects = list(rects)
        self.update()

    def selected_rect(self) -> tuple[float, float, float, float] | None:
        if not self._current_rect:
            return None
        return self._normal_rect(self._current_rect)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor("#f5f7fb"))
        if not self._pixmap:
            painter.setPen(QColor("#64748b"))
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "Load a sample PDF or image")
            return

        self._draw_rect = self._fit_rect()
        painter.drawPixmap(self._draw_rect.toRect(), self._pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        for name, rect in self._rects:
            widget_rect = self._norm_to_widget(rect)
            color = QColor("#16a34a") if name == self.selected_name else QColor("#2563eb")
            painter.setPen(QPen(color, 2))
            painter.drawRect(widget_rect)
            label_rect = QRectF(widget_rect.left(), max(self._draw_rect.top(), widget_rect.top() - 22), 220, 20)
            painter.fillRect(label_rect, QColor(255, 255, 255, 210))
            painter.setPen(color)
            painter.drawText(label_rect, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, f" {name}")

        if self._current_rect:
            painter.setPen(QPen(QColor("#dc2626"), 2, Qt.PenStyle.DashLine))
            painter.drawRect(self._current_rect.normalized())

    def mousePressEvent(self, event) -> None:  # noqa: N802
        if not self._pixmap or event.button() != Qt.MouseButton.LeftButton:
            return
        point = QPointF(event.position())
        if not self._draw_rect.contains(point):
            return
        self._start = point
        self._current_rect = QRectF(point, point)
        self.update()

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        if not self._start:
            return
        point = self._clamp_to_image(QPointF(event.position()))
        self._current_rect = QRectF(self._start, point).normalized()
        self.update()

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if not self._start or event.button() != Qt.MouseButton.LeftButton:
            return
        point = self._clamp_to_image(QPointF(event.position()))
        self._current_rect = QRectF(self._start, point).normalized()
        self._start = None
        rect = self.selected_rect()
        if rect and rect[2] > 0.002 and rect[3] > 0.002:
            self.selection_finished.emit(rect)
        self.update()

    def _fit_rect(self) -> QRectF:
        if not self._pixmap:
            return QRectF()
        margin = 16
        available = self.rect().adjusted(margin, margin, -margin, -margin)
        image_width, image_height = self._image_size
        scale = min(available.width() / image_width, available.height() / image_height)
        width = image_width * scale
        height = image_height * scale
        left = available.left() + (available.width() - width) / 2
        top = available.top() + (available.height() - height) / 2
        return QRectF(left, top, width, height)

    def _clamp_to_image(self, point: QPointF) -> QPointF:
        return QPointF(
            min(max(point.x(), self._draw_rect.left()), self._draw_rect.right()),
            min(max(point.y(), self._draw_rect.top()), self._draw_rect.bottom()),
        )

    def _normal_rect(self, rect: QRectF) -> tuple[float, float, float, float]:
        rect = rect.normalized()
        x = (rect.left() - self._draw_rect.left()) / self._draw_rect.width()
        y = (rect.top() - self._draw_rect.top()) / self._draw_rect.height()
        width = rect.width() / self._draw_rect.width()
        height = rect.height() / self._draw_rect.height()
        x = min(max(x, 0.0), 1.0)
        y = min(max(y, 0.0), 1.0)
        width = min(max(width, 0.0), 1.0 - x)
        height = min(max(height, 0.0), 1.0 - y)
        return x, y, width, height

    def _norm_to_widget(self, rect: tuple[float, float, float, float]) -> QRectF:
        x, y, width, height = rect
        return QRectF(
            self._draw_rect.left() + x * self._draw_rect.width(),
            self._draw_rect.top() + y * self._draw_rect.height(),
            width * self._draw_rect.width(),
            height * self._draw_rect.height(),
        )
