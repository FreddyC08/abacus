import docx
from docx.opc.exceptions import PackageNotFoundError


class Document:
    """Represents a document with tables, paragraphs, and words"""

    def __init__(self, file_path):
        self.doc = None

        try:
            if not file_path.endswith(".docx"):
                print("WRN: File type is not .docx - appending extension")
                file_path += ".docx"
            self.doc = docx.Document(file_path)
        except PackageNotFoundError:
            print(f"ERR: Package {file_path} not found. Please ensure the filename is correct and the file is closed in other programs.")
        except FileNotFoundError:
            print(f"ERR: File {file_path} not found on disk")
        except Exception as e:
            print(f"ERR: {e}")
            return

        self.tables = []
        self.paras = []
        self.words = []
