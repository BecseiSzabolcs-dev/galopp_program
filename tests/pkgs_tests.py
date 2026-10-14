import sys
import importlib.util
import subprocess

class PKG_TEST:
    def __init__(self):
        self.packages = {
            "requests": "requests",
            "numpy": "numpy",
            "pandas": "pandas",
            "lxml": "lxml",
            "beautifulsoup4": "bs4",
            "certifi": "certifi",
            "charset-normalizer": "charset_normalizer",
            "idna": "idna",
            "opencv-python": "cv2",
            "packaging": "packaging",
            "pillow": "PIL",
            "PyQt6": "PyQt6",
            "python-pptx": "pptx",
            "PyPDF2": "PyPDF2",
        }

        # Track missing PyPI packages
        missing_packages = []
        for pip_name, import_name in self.packages.items():
            if importlib.util.find_spec(import_name) is None:
                missing_packages.append(pip_name)

        if missing_packages:
            print(f"Missing packages:\n\t {'\n\t '.join(missing_packages)}")
            self.install_missing_packages(missing_packages)

    def install_missing_packages(self, missing_packages):
        for package in missing_packages:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

if __name__ == "__main__":
    PKG_TEST()       


