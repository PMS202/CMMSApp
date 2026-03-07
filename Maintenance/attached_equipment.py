from PyQt5 import QtWidgets, QtCore, QtGui

class DynamicSuggestion(QtWidgets.QStyledItemDelegate):
    def __init__(self, database, parent=None, dep = None, year = None):
        super().__init__(parent)
        self.database = database
        self.cache = {}
        self.debounce_timer = QtCore.QTimer()
        self.debounce_timer.setSingleShot(True)
        self.pending_prefix = ""
        self.pending_editor = None
        self.dep = dep
        self.year = year

    def createEditor(self, parent, option, index):
        editor = QtWidgets.QLineEdit(parent)

        self.completer_model = QtCore.QStringListModel()
        self.completer = SmartCompleter(self.completer_model, editor)
        self.completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.completer.setFilterMode(QtCore.Qt.MatchContains)
        self.completer.setEditor(editor)
        editor.setCompleter(self.completer)
        self.completer.activated[str].connect(lambda _: self._append_semicolon(editor))
        editor.textEdited.connect(lambda: self._on_text_edited(editor))
        return editor
    
    def _append_semicolon(self, editor):
        def do_append():
            text = editor.text().strip()
            if not text.endswith(';'):
                editor.setText(text + '; ')
                editor.setCursorPosition(len(editor.text()))
        QtCore.QTimer.singleShot(0, do_append)

    def setEditorData(self, editor, index):
        value = index.model().data(index, QtCore.Qt.EditRole)
        editor.setText(value if value else "")

    def setModelData(self, editor, model, index):
        model.setData(index, editor.text(), QtCore.Qt.EditRole)

    def _on_text_edited(self, editor):
        text = editor.text()
        if not text:
            return

        current_prefix = text.split(';')[-1].strip()
        if len(current_prefix) < 2:
            return

        self.pending_prefix = current_prefix
        self.pending_editor = editor
        self.debounce_timer.stop()
        self.debounce_timer.timeout.connect(self._trigger_search)
        self.debounce_timer.start(300)

    def _trigger_search(self):
        prefix = self.pending_prefix
        editor = self.pending_editor

        if prefix in self.cache:
            results = self.cache[prefix]
        else:
            results = self._query_machine_codes(prefix)
            self.cache[prefix] = results

        self._update_completer(editor, prefix, results)

    def _update_completer(self, editor, prefix, results):
        self.completer_model.setStringList(results)
        self.completer.setCompletionPrefix(prefix)
        self.completer.complete()

    def _query_machine_codes(self, prefix):
        try:
            sql = """
                SELECT DISTINCT ( m.machine_code )
                FROM `Maintenance_plan` as mp
                JOIN `Machines` as m
                ON mp.machine_id = m.machine_id
                JOIN `Production_Lines` as p
                ON mp.line_id = p.line_id
                JOIN `Departments` as d
                ON p.department_id = d.department_id
                JOIN `Months_Years` as my
                ON mp.month_year_id = my.month_year_id
                WHERE m.machine_code LIKE :text AND d.department_name = :dep AND my.year = :year
                ORDER BY m.machine_code ASC
                LIMIT 5;
            """
            result = self.database.query(sql = sql, params = {'text':f"%{prefix}%",'dep':self.dep,'year':self.year})
            return [r[0] for r in result]
        except Exception as e:
            return []

class SmartCompleter(QtWidgets.QCompleter):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.editor = None  

    def setEditor(self, editor):
        self.editor = editor

    def pathFromIndex(self, index):
        completion = super().pathFromIndex(index)
        if not self.editor:
            return completion
        current_text = self.editor.text()
        parts = [p.strip() for p in current_text.split(';')]
        if parts:
            parts[-1] = completion
        else:
            parts = [completion]
        return '; '.join(parts)
