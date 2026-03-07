from PyQt5 import QtWidgets, QtGui, QtCore


class ThumbSignals(QtCore.QObject):
    thumbReady = QtCore.pyqtSignal(str, QtGui.QImage)


# ===== Worker load ảnh trong thread pool =====
class ThumbWorker(QtCore.QRunnable):
    def __init__(self, path, size, signal):
        super().__init__()
        self.path = path
        self.size = size
        self.signal = signal

    def run(self):
        img = QtGui.QImage(self.path)
        if img.isNull():
            return
        img = img.scaled(
            self.size[0], self.size[1],
            QtCore.Qt.KeepAspectRatio,
            QtCore.Qt.SmoothTransformation
        )
        self.signal.emit(self.path, img)


class ImageCache:
    _cache = {}
    signals = ThumbSignals()

    @classmethod
    def init(cls, tableview):
        cls.signals.thumbReady.connect(lambda path, img: cls._update(path, img, tableview))

    @classmethod
    def _update(cls, path, img, tableview):
        cls._cache[path] = QtGui.QPixmap.fromImage(img)
        tableview.viewport().update() 

    @classmethod
    def get_pixmap(cls, path, size=(80, 80)):
        if path not in cls._cache:
            cls._cache[path] = QtGui.QPixmap()  # placeholder
            worker = ThumbWorker(path, size, cls.signals.thumbReady)
            QtCore.QThreadPool.globalInstance().start(worker)
        return cls._cache[path]

class StockItemDelegate(QtWidgets.QStyledItemDelegate):
    clicked = QtCore.pyqtSignal(str, QtCore.QModelIndex)

    def __init__(self, buttons=("+",), parent=None):
        super().__init__(parent)
        self._buttons = {}  
        self._hovered = None
        self._button_names = buttons

    def paint(self, painter, option, index):
        painter.save()
        rect = option.rect.adjusted(5, 5, -5, -5)

        if option.state & QtWidgets.QStyle.State_Selected:
            painter.fillRect(option.rect, option.palette.highlight())
        elif option.state & QtWidgets.QStyle.State_MouseOver:
            painter.fillRect(option.rect, QtGui.QColor(240, 240, 240))

        x_offset = 0
        if index.column() == 0:
            data = index.data(QtCore.Qt.UserRole) or {}
            if "image" in data:
                pix = ImageCache.get_pixmap(data["image"], size=(80,80))
                if not pix.isNull():
                    thumb = pix.scaled(80, 80, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
                    painter.drawPixmap(rect.left(), rect.top(), thumb)
                    x_offset = 90
            painter.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))
            painter.setPen(QtGui.QColor("black"))
            painter.drawText(rect.left() + x_offset, rect.top() + 15, data.get("name", ""))
            painter.setFont(QtGui.QFont("Arial", 10))
            painter.drawText(rect.left() + x_offset, rect.top() + 30, f"Code: {data.get('code', '')}")

        elif index.column() == 3:
            stock = index.data()
            painter.setFont(QtGui.QFont("Arial", 10))
            color = QtGui.QColor("#ff0000") if stock != "0.0" else QtGui.QColor("black")
            painter.setPen(color)
            painter.drawText(rect, QtCore.Qt.AlignCenter, str(stock))
        elif index.column() == 10:
            key = (index.row(), index.column())
            count = len(self._button_names)
            if count > 0:
                w = rect.width() // count - (count + 20)
                h = rect.height() - 50
                btn_rects = {}
                for i, name in enumerate(self._button_names):
                    x = rect.left() + 10 + i * (w + 5)
                    y = rect.top() + 25
                    r = QtCore.QRect(x, y, w, h)
                    btn_rects[name] = r
                    hovered = (self._hovered == (name, key))
                    self._drawButton(painter, r, name, "#FFFFFF", "#ff6600", "black", hovered)
                self._buttons[key] = btn_rects
            else:
                if key in self._buttons:
                    del self._buttons[key]
        else:
            super().paint(painter, option, index)

        painter.restore()

    def _drawButton(self, painter, rect, text, bg, hover, text_color, hovered=False):
        painter.save()
        painter.setBrush(QtGui.QColor(bg))
        border_color = QtGui.QColor(hover if hovered else "#CCCCCC")
        painter.setPen(QtGui.QPen(border_color, 1))
        painter.drawRoundedRect(rect, 3, 3)
        painter.setPen(QtGui.QColor(text_color))
        painter.setFont(QtGui.QFont("Arial", 8, QtGui.QFont.Bold))
        painter.drawText(rect, QtCore.Qt.AlignCenter, text)
        painter.restore()

    def editorEvent(self, event, model, option, index):
        key = (index.row(), index.column())
        btns = self._buttons.get(key, {})

        if event.type() == QtCore.QEvent.MouseMove:
            pos = event.pos()
            for name, rect in btns.items():
                if rect.contains(pos):
                    if self._hovered != (name, key):
                        self._hovered = (name, key)
                        option.widget.viewport().update()
                    return True
            if self._hovered:
                self._hovered = None
                option.widget.viewport().update()
            return True

        elif event.type() == QtCore.QEvent.MouseButtonRelease:
            pos = event.pos()
            for name, rect in btns.items():
                if rect.contains(pos):
                    self.clicked.emit(name, index)
                    return True

        return super().editorEvent(event, model, option, index)

    def sizeHint(self, option, index):
        return QtCore.QSize(200, 80)
