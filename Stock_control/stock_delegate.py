from PyQt5 import QtWidgets, QtGui, QtCore
import os


class ThumbSignals(QtCore.QObject):
    thumbReady = QtCore.pyqtSignal(object, QtGui.QImage)

    def __init__(self):
        super().__init__()


class ThumbWorker(QtCore.QRunnable):
    def __init__(self, path, size):
        super().__init__()
        self.path = path
        self.size = size
        # self.signal = signal

    def run(self):
        img = QtGui.QImage(self.path)
        if img.isNull():
            return
        img = img.scaled(
            self.size[0], self.size[1],
            QtCore.Qt.KeepAspectRatio,
            QtCore.Qt.SmoothTransformation
        )
        ImageCache.signals.thumbReady.emit(self.path, img)


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
    def create_placeholder(cls, size=(80, 80)):
        pix = QtGui.QPixmap(*size)
        pix.fill(QtGui.QColor("#f0f0f0"))

        painter = QtGui.QPainter(pix)
        painter.setPen(QtGui.QPen(QtGui.QColor("#b0b0b0"), 1))
        margin = 15
        r = pix.rect().adjusted(margin, margin, -margin, -margin)
        painter.drawRoundedRect(r, 4, 4)
        painter.drawEllipse(r.left()+5, r.top()+5, 8, 8)
        # painter.drawRect(pix.rect().adjusted(10,10,-10,-10))
        painter.drawText(
            pix.rect(),
            QtCore.Qt.AlignCenter,
            "NO\nIMAGE"
        )
        painter.end()

        return pix
    
    @classmethod
    def get_pixmap(cls, path, size=(80, 80)):
        if path not in cls._cache:
            cls._cache[path] = cls.create_placeholder(size)
            if path and os.path.exists(path):
                worker = ThumbWorker(path, size)
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
        header_text = index.model().headerData(
            index.column(),
            QtCore.Qt.Horizontal,
            QtCore.Qt.DisplayRole
        )
        if header_text == "Spare part":
            data = index.data(QtCore.Qt.UserRole) or {}
            if "image" in data:
                pix = ImageCache.get_pixmap(data["image"], size=(80,80))
                if not pix.isNull():
                    thumb = pix.scaled(80, 80, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
                    img_rect = QtCore.QRect(
                                            rect.left(),
                                            rect.top(),
                                            80,
                                            80
                                        )
                    x = img_rect.left() + (img_rect.width() - thumb.width()) // 2
                    y = img_rect.top() + (img_rect.height() - thumb.height()) // 2  - 5
                    painter.drawPixmap(x, y, thumb)
                    x_offset = 90
            painter.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))
            painter.setPen(QtGui.QColor("black"))
            painter.drawText(rect.left() + x_offset, rect.top() + 15, data.get("name", ""))
            painter.setFont(QtGui.QFont("Arial", 10))
            painter.drawText(rect.left() + x_offset, rect.top() + 30, f"Code: {data.get('code', '')}")

        elif header_text == "Stock\nstatus":
            stock_status = index.data()
            painter.setFont(QtGui.QFont("Arial", 10, QtGui.QFont.Bold))
            if stock_status == "Urgent":
                color = QtGui.QColor("#ff0000")
            elif stock_status == "Below Min Stock":
                color = QtGui.QColor("#ff6600")
            elif stock_status == "Overstock":
                color = QtGui.QColor("#0000ff")
            else:
                color = QtGui.QColor("black")
            painter.setPen(color)
            painter.drawText(rect, QtCore.Qt.AlignCenter, str(stock_status) if stock_status != "Below Min Stock" else "Below\nMin Stock")
        # elif index.column() == 12:
        #     key = (index.row(), index.column())
        #     count = len(self._button_names)
        #     if count > 0:
        #         w = rect.width() // count - (count + 20)
        #         h = rect.height() - 50
        #         btn_rects = {}
        #         for i, name in enumerate(self._button_names):
        #             x = rect.left() + 10 + i * (w + 5)
        #             y = rect.top() + 25
        #             r = QtCore.QRect(x, y, w, h)
        #             btn_rects[name] = r
        #             hovered = (self._hovered == (name, key))
        #             self._drawButton(painter, r, name, "#FFFFFF", "#ff6600", "black", hovered)
        #         self._buttons[key] = btn_rects
        #     else:
        #         if key in self._buttons:
        #             del self._buttons[key]
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

class StockAlertDelegate(QtWidgets.QStyledItemDelegate):

    def paint(self, painter, option, index):

        header = index.model().headerData(
            index.column(),
            QtCore.Qt.Horizontal,
            QtCore.Qt.DisplayRole
        )

        if header != "Status":
            return super().paint(painter, option, index)

        painter.save()
        painter.setRenderHint(QtGui.QPainter.Antialiasing)

        status = index.data()

        if status == "Urgent":
            bg = QtGui.QColor("#FFDADA")
            border = QtGui.QColor("#E74C3C")
            text = QtGui.QColor("#9C1000")

        elif status == "Below Min Stock":
            bg = QtGui.QColor("#FFEDD6")
            border = QtGui.QColor("#F39C12")
            text = QtGui.QColor("#A76700")

        elif status == "Overstock":
            bg = QtGui.QColor("#D2E6FF")
            border = QtGui.QColor("#3498DB")
            text = QtGui.QColor("#21618C")

        else:
            bg = QtGui.QColor("#EEEEEE")
            border = QtGui.QColor("#BBBBBB")
            text = QtGui.QColor("#666666")

        badge = option.rect.adjusted(8, 6, -8, -6)

        painter.setPen(QtGui.QPen(border, 1))
        painter.setBrush(bg)
        painter.drawRoundedRect(badge, 8, 8)
        painter.setPen(text)
        painter.drawText(badge, QtCore.Qt.AlignCenter, status)
        painter.restore()