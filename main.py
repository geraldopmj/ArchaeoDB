from PySide6.QtWidgets import QApplication, QMainWindow
from ui import Ui_MainWindow
import sys
import os

if __name__ == "__main__":
    app = QApplication(sys.argv)  
    base_dir = os.path.dirname(os.path.abspath(__file__ ))  
    db_path = os.path.join(base_dir, "Data")
    os.makedirs(os.path.dirname(db_path), exist_ok=True)

    MainWindow = QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()

    sys.exit(app.exec())
