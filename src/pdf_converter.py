import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QListWidget, QPushButton, QFileDialog, QLabel, QProgressDialog
from PyQt5.QtGui import QDropEvent, QDragEnterEvent
from PyQt5.QtCore import Qt
from docx2pdf import convert
import threading

class Application(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DOCX to PDF Converter")
        self.file_paths = []
        self.process_cancelled = False
        self.setAcceptDrops(True)
        self.create_widgets()

    def create_widgets(self):
        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        self.description = QLabel("Quickly add docx files: Drag and drop or browse.", self)
        main_layout.addWidget(self.description)

        layout12 = QHBoxLayout()
        main_layout.addLayout(layout12)

        self.list_widget = QListWidget(self)
        layout12.addWidget(self.list_widget)

        input_layout = QVBoxLayout()
        input_layout.setAlignment(Qt.AlignTop)
        layout12.addLayout(input_layout)

        self.browse_button = QPushButton("Browse", self)
        self.browse_button.clicked.connect(self.select_files)
        input_layout.addWidget(self.browse_button)

        self.remove_button = QPushButton("Remove", self)
        self.remove_button.clicked.connect(self.remove_paths)
        input_layout.addWidget(self.remove_button)

        self.clear_button = QPushButton("Clear", self)
        self.clear_button.clicked.connect(self.clear_list)
        input_layout.addWidget(self.clear_button)

        input_layout.addSpacing(80)
        
        self.convert_button = QPushButton("Convert", self)
        self.convert_button.clicked.connect(self.convert_button_event)
        input_layout.addWidget(self.convert_button)


    def select_files(self):
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("DOCX Files (*.docx)")
        file_dialog.setWindowTitle("Select DOCX files")

        if file_dialog.exec(): 
            for file in file_dialog.selectedFiles():
                self.add_file_path(file)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            for url in urls:
                if not url.toLocalFile().endswith('.docx'):
                    return
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            dragged_file_paths = [url.toLocalFile() for url in event.mimeData().urls() if url.toLocalFile().endswith('.docx')]
            for file in dragged_file_paths:
                self.add_file_path(file)

    def add_file_path(self, file_path):
        if file_path not in self.file_paths:
                self.file_paths.append(file_path)
                self.list_widget.addItem(file_path)      

    def convert_button_event(self):
        if self.file_paths:
            target_folder = QFileDialog.getExistingDirectory(self, "Select Folder")
            if target_folder:
                thread = threading.Thread(target=self.convert_to_pdf, args=(target_folder,))
                thread.start()
                self.start_progress_dialog(len(self.file_paths))

    def convert_to_pdf(self, target_folder):
        for file_path in self.file_paths:
            filename = os.path.basename(file_path).replace(".docx","")
            target_path = f"{target_folder}/{filename}.pdf"
            counter = 1
            while os.path.exists(target_path):
                target_path = f"{target_folder}/{filename}({counter}).pdf"
                counter += 1
            self.progress_dialog.setLabelText(file_path)
            
            convert(file_path, target_path)
            if self.process_cancelled:
                break
            self.increment_progress_bar()

    def remove_paths(self):
        for item in self.list_widget.selectedItems():
            self.list_widget.takeItem(self.list_widget.row(item))
            self.file_paths.remove(item.text())


    def clear_list(self):
        self.list_widget.clear()
        self.file_paths = []

    def start_progress_dialog(self, max):
        self.progress_dialog = QProgressDialog()
        self.progress_dialog.setWindowTitle('DOCX to PDF Converter')
        self.progress_dialog.setLabelText('Processing...')
        self.progress_dialog.setWindowModality(Qt.NonModal)
        self.progress_dialog.setWindowFlags(Qt.CustomizeWindowHint | Qt.WindowTitleHint)
        self.progress_dialog.setMinimum(0)
        self.progress_dialog.setMaximum(max)
        self.progress_dialog.setAutoClose(True)
        self.progress_dialog.setValue(0)

        cancel_button = QPushButton("Cancel", self.progress_dialog)
        cancel_button.clicked.connect(self.cancel_button_event)
        self.progress_dialog.setCancelButton(cancel_button)

        self.progress_dialog.resize(400, self.progress_dialog.height())

        self.progress_dialog.exec_()

    def increment_progress_bar(self):
        self.progress_dialog.setValue(self.progress_dialog.value() + 1)

    def cancel_button_event(self):
        self.process_cancelled = True

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Application()
    window.setGeometry(200, 200, 400, 200)
    window.show()
    sys.exit(app.exec_())
