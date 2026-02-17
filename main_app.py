import os
import sys
base_directory = os.path.dirname(sys.argv[0])
plugin_path = os.path.join(base_directory, "platforms")
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = plugin_path
import warnings
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QFileDialog # Engine and container.
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QObject, Signal, QThread, QFile, QIODevice
from pathlib import Path
import backend_functions

# Working container for dynamic display of progress messages.
class Worker(QObject):
    progress = Signal(str)
    finished = Signal(str)

    def __init__(self, input_file, output_folder):
        super().__init__()
        self.input_file = input_file
        self.output_folder = output_folder

    def run(self):
        def report_to_gui(message: str):
            self.progress.emit(message)

        try:
            result = backend_functions.extract_words(self.input_file, self.output_folder, progress = report_to_gui)
            self.finished.emit(result)
        except Exception as e:
            self.finished.emit(f"Error message: {e}")

# Organizes all button functions together.
class MainApp(QMainWindow):
    # Standard setup from the parent function.
    def __init__(self):
        super().__init__()

        loader = QUiLoader()
        ui_path = os.path.join(base_directory, 'mainwindow.ui')
        ui_file = QFile(ui_path)
        if not ui_file.open(QIODevice.ReadOnly):
            raise RuntimeError(f"UI file not found at {ui_path}.\n")
        self.window = loader.load(ui_file)
        ui_file.close()
        self.setCentralWidget(self.window) # Put the loaded window into the container.


        self.excel_btn = self.window.findChild(QPushButton, 'excelBtn')
        self.excel_btn.clicked.connect(self.excel_clicked)

        self.folder_btn = self.window.findChild(QPushButton, 'folderBtn')
        self.folder_btn.clicked.connect(self.folder_clicked)

        self.function_btn = self.window.findChild(QPushButton, 'functionBtn')
        self.function_btn.clicked.connect(self.function_clicked)

        self.exit_btn = self.window.findChild(QPushButton, 'exitBtn')
        self.exit_btn.clicked.connect(self.exit_clicked)


    def excel_clicked(self):
        file_path, _ = QFileDialog.getOpenFileName(self.window, 'Select an Excel file', '', "Excel files (*.xlsx)")
        if file_path:
            excel_label = self.window.findChild(QLabel, 'fileLabel')
            excel_label.setText(f"Selected <<{Path(file_path).name}>>.")
            self.input_file = file_path

    def folder_clicked(self):
        folder_path = QFileDialog.getExistingDirectory(self.window, 'Select an output folder', '')
        if folder_path:
            excel_label = self.window.findChild(QLabel, 'folderLabel')
            excel_label.setText(f"Selected <<{Path(folder_path).name}>>.")
            self.output_folder = folder_path

    def function_clicked(self):
        progress_label = self.window.findChild(QLabel, 'progressLabel')

        if not hasattr(self, 'input_file') or not hasattr(self, 'output_folder'):
            progress_label.setText("❌ Please choose an Excel file and an output folder before running extraction.")
            return

        try:
            self.thread = QThread()
            self.worker = Worker(self.input_file, self.output_folder)
            self.worker.moveToThread(self.thread)

            self.thread.started.connect(self.worker.run)
            self.worker.progress.connect(progress_label.setText)
            self.worker.finished.connect(progress_label.setText)
            self.worker.finished.connect(self.thread.quit)
            self.worker.finished.connect(self.worker.deleteLater)
            self.thread.finished.connect(self.thread.deleteLater)

            self.thread.start()

        except Exception as e:
            progress_label.setText(f"❌ Error message: {e}")

    def exit_clicked(self):
        self.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec())
