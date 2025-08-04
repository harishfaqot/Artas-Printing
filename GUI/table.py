from PyQt5.QtWidgets import QApplication, QTableWidget, QTableWidgetItem
from PyQt5.QtCore import Qt, QModelIndex
import sys

class TableWidgetWithCopyPaste(QTableWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
    
    def keyPressEvent(self, event):
        # Copy action (Ctrl+C)
        if event.key() == Qt.Key_C and event.modifiers() == Qt.ControlModifier:
            self.copy()
        # Paste action (Ctrl+V)
        elif event.key() == Qt.Key_V and event.modifiers() == Qt.ControlModifier:
            self.paste()
        # Cut action (Ctrl+X)
        elif event.key() == Qt.Key_X and event.modifiers() == Qt.ControlModifier:
            self.cut()
        else:
            super().keyPressEvent(event)
    
    def copy(self):
        selected = self.selectedIndexes()
        if not selected:
            return
        
        # Sort by row and column to get proper order
        selected.sort()
        rows = selected[-1].row() - selected[0].row() + 1
        cols = selected[-1].column() - selected[0].column() + 1
        
        # Create text with tab separation and newlines for rows
        text = ""
        for i in range(rows):
            for j in range(cols):
                if j > 0:
                    text += "\t"
                item = self.item(selected[0].row() + i, selected[0].column() + j)
                if item is not None:
                    text += item.text()
            text += "\n"
        
        QApplication.clipboard().setText(text)
    
    def paste(self):
        selected = self.selectedIndexes()
        if not selected:
            return
        
        # Get the first selected cell position
        row = selected[0].row()
        col = selected[0].column()
        
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        rows = text.split('\n')
        
        for i, r in enumerate(rows):
            if not r:
                continue
            cells = r.split('\t')
            for j, c in enumerate(cells):
                if row + i >= self.rowCount() or col + j >= self.columnCount():
                    continue
                item = self.item(row + i, col + j)
                if item is not None:
                    item.setText(c)
                else:
                    self.setItem(row + i, col + j, QTableWidgetItem(c))
    
    def cut(self):
        self.copy()
        for item in self.selectedItems():
            item.setText("")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    table = TableWidgetWithCopyPaste(5, 5)
    table.setHorizontalHeaderLabels(['A', 'B', 'C', 'D', 'E'])
    table.setVerticalHeaderLabels(['1', '2', '3', '4', '5'])

    
    table.resize(600, 400)
    table.show()
    
    sys.exit(app.exec_())