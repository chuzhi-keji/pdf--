# main.py
import sys
from PyQt6.QtWidgets import QApplication
# Assuming gui folder is at the same level as main.py or in sys.path
from gui.main_window import MainWindow

if __name__ == '__main__':
    app = QApplication(sys.argv)

    # Optional: Set an application icon (replace 'app_icon.png' with your actual icon path)
    # app_icon_path = os.path.join(os.path.dirname(__file__), 'resources', 'app_icon.png')
    # if os.path.exists(app_icon_path):
    #    app.setWindowIcon(QIcon(app_icon_path))

    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())