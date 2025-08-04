from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QTableWidget, QTableWidgetItem

def setup_table_functionality(self, table):
        def copy():
            selected = table.selectedIndexes()
            if not selected:
                return
            
            selected.sort()
            rows = selected[-1].row() - selected[0].row() + 1
            cols = selected[-1].column() - selected[0].column() + 1
            
            text = ""
            for i in range(rows):
                for j in range(cols):
                    if j > 0:
                        text += "\t"
                    item = table.item(selected[0].row() + i, selected[0].column() + j)
                    if item is not None:
                        text += item.text()
                text += "\n"
            
            QApplication.clipboard().setText(text)
        
        def paste():
            selected = table.selectedIndexes()
            if not selected:
                return

            row = selected[0].row()
            col = selected[0].column()

            clipboard = QApplication.clipboard()
            text = clipboard.text()
            rows = text.split('\n')

            for i, r in enumerate(rows):
                if not r.strip():
                    continue
                cells = r.split('\t')
                for j, c in enumerate(cells):
                    if row + i >= table.rowCount() or col + j >= table.columnCount():
                        continue

                    item = table.item(row + i, col + j)
                    if item is not None:
                        item.setText(c)
                        item.setTextAlignment(Qt.AlignCenter)
                    else:
                        new_item = QTableWidgetItem(c)
                        new_item.setTextAlignment(Qt.AlignCenter)
                        table.setItem(row + i, col + j, new_item)

        def keyPressEvent(event):
            if event.key() == Qt.Key_C and event.modifiers() == Qt.ControlModifier:
                copy()
            elif event.key() == Qt.Key_V and event.modifiers() == Qt.ControlModifier:
                paste()
            else:
                QTableWidget.keyPressEvent(table, event)
        
        table.copy = copy
        table.paste = paste
        table.keyPressEvent = keyPressEvent