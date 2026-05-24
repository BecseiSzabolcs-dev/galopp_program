import mimetypes
import os

from PyQt6.QtWidgets import QDialog, QFileDialog, QLabel, QPushButton, QVBoxLayout


class Addon:
    def __init__(self, data):
        if isinstance(data, str):
            try:
                id, name, ftype, path = data.split(";")

                self.id = id
                self.name = name
                self.type = ftype
                self.path = path

            except Exception as e:
                print(e)
                self.id = ""
                self.name = ""
                self.type = ""
                self.path = ""
        if isinstance(data, dict):
            # ['Id','Name','Type',"Path"]
            try:
                self.id = data["Id"]
                self.name = data["Name"]
                self.type = data["Type"]
                self.path = data["Path"]
            except Exception as e:
                print(f"Error: {e}")

    def get_type(self):
        if self.type.startswith("video"):
            return "video"
        elif self.type.startswith("image"):
            return "image"
        elif self.type.startswith("sound"):
            return "sound"
        else:
            return None

    def __str__(self):
        try:
            return f"{self.id};{self.name};{self.type};{self.path}"
        except:
            return None


class AddonSel(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI()
        self.file_data_list = []
        self.show()

    def initUI(self):
        self.setWindowTitle("Media to PPTX Converter")
        self.setGeometry(300, 300, 400, 150)

        layout = QVBoxLayout()

        self.label = QLabel("Select a folder containing images/videos", self)
        layout.addWidget(self.label)

        # Button to select a folder
        self.btn_folder = QPushButton("Select Folder", self)
        self.btn_folder.clicked.connect(self.process_folder)
        layout.addWidget(self.btn_folder)

        # Button to select a single file (optional)
        self.btn_file = QPushButton("Select Single File", self)
        self.btn_file.clicked.connect(self.process_single_file)
        layout.addWidget(self.btn_file)

        self.setLayout(layout)

    def get_files(self):
        return self.file_data_list

    def process_files(self, file_list):

        for i, file_path in enumerate(file_list, 1):
            mime, _ = mimetypes.guess_type(file_path)
            if mime and (mime.startswith("video") or mime.startswith("image")):
                data_rec = {
                    "Id": i,
                    "Name": file_path.split("/")[-1],
                    "Type": mime,
                    "Path": file_path,
                }
                self.file_data_list.append(Addon(data_rec))

        """
        for record in self.file_data_list:
            print(record)
        """

        self.accept()

    def process_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder_path:
            # Get all files in the folder
            files = [os.path.join(folder_path, f) for f in os.listdir(folder_path)]
            self.process_files(files)

    def process_single_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Media File",
            "",
            "Media Files (*.mp4 *.jpg *.png *.jpeg *.mov)",
        )
        if file_path:
            self.process_files([file_path])


if __name__ == "__main__":
    addonsel = AddonSel()

    for rec in addonsel.get_files():
        print(rec)
