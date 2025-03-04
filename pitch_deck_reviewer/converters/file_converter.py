import os
import sys
import subprocess

try:
    import comtypes.client

    COMTYPES_AVAILABLE = True
except ImportError:
    COMTYPES_AVAILABLE = False


class FileConverter:
    def __init__(self, output_dir="converted_pdfs"):
        self.output_dir = output_dir
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

    def convert_to_pdf(self, file_path):
        """Converts a file to PDF based on its extension."""
        try:
            base_name = os.path.basename(file_path)
            pdf_path = os.path.join(
                self.output_dir, base_name.rsplit(".", 1)[0] + ".pdf"
            )

            if file_path.endswith(".ppt") or file_path.endswith(".pptx"):
                return self._convert_ppt_to_pdf(file_path, pdf_path)
            elif file_path.endswith(".doc") or file_path.endswith(".docx"):
                return self._convert_doc_to_pdf(file_path, pdf_path)
            elif file_path.endswith(".xls") or file_path.endswith(".xlsx"):
                return self._convert_excel_to_pdf(file_path, pdf_path)
            elif file_path.endswith(".pdf"):
                return file_path  # If it's already a PDF, return the path directly
            else:
                print(f"Unsupported file format: {file_path}")
                return None
        except Exception as e:
            print(f"Error converting file to PDF: {e}")
            return None

    def _convert_ppt_to_pdf(self, ppt_path, pdf_path):
        """Converts a PPT or PPTX file to a PDF."""
        if sys.platform == "win32" and COMTYPES_AVAILABLE:
            return self._convert_ppt_windows(ppt_path, pdf_path)
        else:
            return self._convert_ppt_libreoffice(ppt_path, pdf_path)

    def _convert_ppt_windows(self, ppt_path, pdf_path):
        try:
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1
            presentation = powerpoint.Presentations.Open(os.path.abspath(ppt_path))
            presentation.SaveAs(
                os.path.abspath(pdf_path), 32
            )  # 32 is the format for PDF
            presentation.Close()
            powerpoint.Quit()
            return pdf_path
        except Exception as e:
            print(f"Error converting PPT to PDF on Windows: {e}")
            return None

    def _convert_ppt_libreoffice(self, ppt_path, pdf_path):
        try:
            subprocess.run(
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    self.output_dir,
                    ppt_path,
                ],
                check=True,
            )
            return pdf_path
        except Exception as e:
            print(f"Error converting PPT to PDF: {e}")
            return None

    def _convert_doc_to_pdf(self, doc_path, pdf_path):
        """Converts a DOC or DOCX file to a PDF."""
        try:
            subprocess.run(
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    self.output_dir,
                    doc_path,
                ],
                check=True,
            )
            return pdf_path
        except Exception as e:
            print(f"Error converting DOC to PDF: {e}")
            return None

    def _convert_excel_to_pdf(self, excel_path, pdf_path):
        """Converts an XLS or XLSX file to a PDF."""
        try:
            subprocess.run(
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    self.output_dir,
                    excel_path,
                ],
                check=True,
            )
            return pdf_path
        except Exception as e:
            print(f"Error converting Excel to PDF: {e}")
            return None
