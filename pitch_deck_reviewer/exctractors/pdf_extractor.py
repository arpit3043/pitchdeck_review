import fitz


class PDFExtractor:
    def extract_text_from_pdf(self, pdf_path):
        """Extracts text from a given PDF file."""
        try:
            doc = fitz.open(pdf_path)
            slides_content = []
            for i, page in enumerate(doc):
                text = page.get_text("text")
                slides_content.append({"slide_number": i + 1, "text": text.strip()})
            return slides_content
        except Exception as e:
            print(f"Error extracting text from PDF: {e}")
            return []
