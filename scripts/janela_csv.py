import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
    QFileDialog, QProgressBar, QLabel, QHBoxLayout
)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Leitor de CSV com PyQt5')
        self.setGeometry(100, 100, 800, 600)
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.layout.addWidget(self.progress_bar)

        # Tabela vazia
        self.table = QTableWidget()
        self.layout.addWidget(self.table)

        # Botões
        self.button_layout = QHBoxLayout()
        self.btn_select_file = QPushButton('Selecionar arquivo CSV')
        self.btn_select_folder = QPushButton('Selecionar pasta de saída')
        self.button_layout.addWidget(self.btn_select_file)
        self.button_layout.addWidget(self.btn_select_folder)
        self.layout.addLayout(self.button_layout)

        # Sinal dos botões será implementado depois

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
