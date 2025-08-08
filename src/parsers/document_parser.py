from models.document import Document


class DocumentParser:
    """Data access object for interacting with .docx files"""

    def __init__(self, file_path, options):
        valid_keys = (
            "Headers",
            "Table headers",
            "Table bodies",
            "Captions",
            "References",
            "Appendix",
        )
        self.selected_options = {
            key: (key in options if options else False) for key in valid_keys
        }
        print(f"self.selected_options: {self.selected_options}")
        self.document = Document(file_path)
        if self.document.doc is None:
            return

        self.initialise_document()

    @property
    def document_loaded(self):
        return self.document.doc is not None

    def initialise_document(self):
        self.document.tables = self.document.doc.tables
        self.document.paras = list(self.set_paragraphs())
        self.document.words = list(self.set_words())

    def set_paragraphs(self):
        counter = 0

        while counter < len(self.document.doc.paragraphs):
            paragraph = self.document.doc.paragraphs[counter]
            counter += 1

            # At an appendix, stop count
            if paragraph.text.strip().lower().startswith(("appendix", "appendices")):
                break

            if self.selected_options["Headers"] and (
                not paragraph.style.name.startswith("Heading")
                and not paragraph.style.name.startswith("Title")
                and not paragraph.style.name.startswith("Subtitle")
                and paragraph.text != ""
                and not paragraph.text.startswith("Figure")
            ):
                yield paragraph

            if self.selected_options["Headers"]:
                # Include all paragraphs (including headings)
                if paragraph.text.strip() != "":
                    yield paragraph
            else:
                # Exclude headings/titles/subtitles
                if (
                    not paragraph.style.name.startswith("Heading")
                    and not paragraph.style.name.startswith("Title")
                    and not paragraph.style.name.startswith("Subtitle")
                    and paragraph.text.strip() != ""
                    and not paragraph.text.startswith("Figure")
                ):
                    yield paragraph

        for table in self.document.tables:
            for index, row in enumerate(table.rows):
                if index == 0 and self.selected_options["Table headers"]:
                    continue
                if not self.selected_options["Table bodies"]:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if paragraph.text != "":
                                yield paragraph

        for table in self.document.tables:
            for index, row in enumerate(table.rows):
                # Skip header row if option says so
                if index == 0 and not self.selected_options.get("Table headers", True):
                    continue

                # Skip table bodies if option says so
                if not self.selected_options.get("Table bodies", True):
                    continue

                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if paragraph.text.strip():
                            yield paragraph

    def set_words(self):
        for paragraph in self.document.paras:
            paragraph = paragraph.text
            words = paragraph.split()
            for word in words:
                if word not in {".", "–"}:
                    yield word

    def print_word_count(self):
        print(f"Words: {len(self.document.words)}")
        # print(f"Included words: {self.document.words}")
