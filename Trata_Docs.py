import sys
import pandas as pd
import re
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QPushButton, QFileDialog, QTextEdit, QVBoxLayout, QWidget, QMessageBox
from PyQt5.QtCore import Qt, QAbstractTableModel

APP_VERSION = "1.0.1"

class PandasModel(QAbstractTableModel):
    def __init__(self, data, headers):
        QAbstractTableModel.__init__(self)
        self._data = data
        self._headers = headers

    def rowCount(self, parent=None):
        return len(self._data.index)

    def columnCount(self, parent=None):
        return len(self._data.columns)

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return self._headers[section]
            elif orientation == Qt.Vertical:
                return str(section + 1)
        return None


class ExcelViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Excel Viewer {APP_VERSION}")
        self.setGeometry(100, 100, 800, 600)

        # Create widgets
        self.table_view = QTableView(self)
        self.load_file_button = QPushButton("Carregar Arquivo", self)
        self.treat_file_button = QPushButton("Tratar Arquivo", self)
        self.save_file_button = QPushButton("Salvar Arquivo", self)
        self.file_path_text = QTextEdit(self)
        self.file_path_text.setReadOnly(True)
        self.file_path_text.setFixedHeight(30)
        self.treat_file_button.setStyleSheet("background-color : lightblue")

        # Create layout and add widgets
        layout = QVBoxLayout()
        layout.addWidget(self.file_path_text)
        layout.addWidget(self.table_view)
        layout.addWidget(self.load_file_button)
        layout.addWidget(self.treat_file_button)
        layout.addWidget(self.save_file_button)

        # Set layout for the main window
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Connect signals to slots
        self.load_file_button.clicked.connect(self.load_file)
        self.treat_file_button.clicked.connect(self.treat_file)
        self.save_file_button.clicked.connect(self.save_file)

    def load_file(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
        if filename:
            self.file_path_text.setPlainText(filename)
            self.load_excel_file(filename)

    def load_excel_file(self, filename):
        try:
            # Load the data from the Excel file using Pandas
            data = pd.read_excel(filename, skiprows=2)
            data = data[['Nome Completo', 'E-mail', 'Lista de Distribuição']]
            headers = data.columns.tolist()
            self.table_view.setModel(PandasModel(data, headers))
            self.table_view.setModel(PandasModel(data, headers))
            self.table_view.setColumnWidth(0,200)
            self.table_view.setColumnWidth(1,200)
            self.table_view.setColumnWidth(2,200)
        except pd.errors.ParserError:
            error_message = "Arquivo fora do padrão indicado"
            QMessageBox.warning(self, "Error", error_message)
        except Exception as e:
                # Show an error message for any other exception
                error_message = f"Um erro ocorreu ao tentar abrir o arquivo"
                QMessageBox.warning(self, "Error", error_message)

    def treat_file(self):
        try:
            # Treat the data from the Excel file
            data = self.table_view.model()._data
            data = data[['Nome Completo', 'E-mail', 'Lista de Distribuição']]
            data = (data.set_index(['Nome Completo', 'E-mail'])
                    .stack()
                    .str.split(r'[,;|]', expand=True)
                    .apply(lambda x: x.str.strip())
                    .stack()
                    .unstack(-2)
                    .reset_index(-1, drop=True)
                    .reset_index())
            headers = data.columns.tolist()
            self.table_view.setModel(PandasModel(data, headers))
            self.table_view.setColumnWidth(0,200)
            self.table_view.setColumnWidth(1,200)
            self.table_view.setColumnWidth(2,200)
        except AttributeError:
            error_message = "Nenhum arquivo carregado, carregue um arquivo primeiro"
            QMessageBox.warning(self, "Error", error_message)
        except Exception as e:
                    # Show an error message for any other exception
                    error_message = f"Um erro ocorreu ao tentar tratar o arquivo:\n{str(e)}"
                    QMessageBox.warning(self, "Error", error_message)    
    def save_file(self):
        filename, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
        if filename:
            try:
                # Save the data to Excel file using Pandas
                data = self.table_view.model()._data
                data.to_excel(filename, index=False)
            except ArithmeticError:
                error_message = "Não há dados a serem salvos. Por favor, carregue e trate um arquivo primeiro"
                QMessageBox.warning(self, "Error", error_message)
            except Exception as e:
                # Show an error message for any other exception
                error_message = f"Um erro ocorreu ao tentar salvar o arquivo:\n{str(e)}"
                QMessageBox.warning(self, "Error", error_message)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    excel_viewer = ExcelViewer()
    excel_viewer.show()
    sys.exit(app.exec_())