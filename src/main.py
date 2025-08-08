import sys
from parsers.document_parser import DocumentParser
import questionary


if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = input("Enter file path to document: ")
    print()

    selected_options = questionary.checkbox(
        "Select objects to include",
        choices=[
            {"name": "Headers", "checked": False},
            {"name": "Table headers", "checked": False},
            {"name": "Table bodies", "checked": False},
            {"name": "Captions", "checked": False},
            {"name": "References", "checked": False},
            {"name": "Appendix", "checked": False},
        ],
    ).ask()

    document_dao = DocumentParser(file_path, selected_options)
    if document_dao.document_loaded:
        document_dao.print_word_count()
