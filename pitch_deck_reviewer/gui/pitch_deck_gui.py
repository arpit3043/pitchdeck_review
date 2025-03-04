import tkinter as tk
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
import webbrowser
import os


class PitchDeckGUI(TkinterDnD.Tk):
    def __init__(self, file_converter, pdf_extractor, pitch_deck_reviewer):
        """GUI for Pitch Deck Reviewer"""
        super().__init__()
        self.title("Pitch Deck Reviewer")
        self.geometry("500x400")

        self.label = tk.Label(
            self, text="Drag and drop your file here", font=("Arial", 14)
        )
        self.label.pack(pady=10)

        self.select_button = tk.Button(
            self, text="Select File", command=self.select_file
        )
        self.select_button.pack(pady=10)

        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self.process_dropped_file)

        self.file_converter = file_converter
        self.pdf_extractor = pdf_extractor
        self.pitch_deck_reviewer = pitch_deck_reviewer

    def select_file(self):
        """Opens file dialog for selecting PPT, PPTX, DOC, DOCX, XLS, XLSX, or PDF files."""
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Supported Files", "*.ppt;*.pptx;*.doc;*.docx;*.xls;*.xlsx;*.pdf")
            ]
        )
        if file_path:
            self.process_file(file_path)

    def process_dropped_file(self, event):
        """Handles drag-and-drop file selection."""
        file_path = event.data.strip("{}")
        self.process_file(file_path)

    def process_file(self, file_path):
        """Processes the selected file (PPT, PPTX, DOC, DOCX, XLS, XLSX, or PDF)."""
        try:
            self.label.config(text="Processing file...")
            self.update_idletasks()

            # Show loading buffer
            self.show_loading_buffer()

            # Process the file in a separate thread to keep the GUI responsive
            threading.Thread(
                target=self._process_file_thread, args=(file_path,)
            ).start()
        except Exception as e:
            self.label.config(text=f"Error: {e}")

    def _process_file_thread(self, file_path):
        try:
            pdf_path = self.file_converter.convert_to_pdf(file_path)

            if not pdf_path:
                self.label.config(text="Error: Conversion to PDF failed!")
                self.hide_loading_buffer()
                return

            extracted_content = self.pdf_extractor.extract_text_from_pdf(pdf_path)
            if not extracted_content:
                self.label.config(text="Error: Failed to extract text from PDF!")
                self.hide_loading_buffer()
                return

            self.pitch_deck_reviewer.slides_content = extracted_content
            feedback = self.pitch_deck_reviewer.analyze_pitch_deck()

            if "error" in feedback:
                self.label.config(text="Error: No content extracted from the file!")
            else:
                report_path = "pitch_deck_report.md"
                with open(report_path, "w", encoding="utf-8") as report_file:
                    report_file.write(str(feedback))

                self.label.config(
                    text=f"✅ Analysis Complete! Report saved: {report_path}"
                )
                self.hide_loading_buffer()
                self.after(2000, self.close_and_open_report, report_path)
        except Exception as e:
            self.label.config(text=f"Error: {e}")
            self.hide_loading_buffer()

    def show_loading_buffer(self):
        """Shows a loading buffer on the GUI."""
        self.loading_label = tk.Label(self, text="Loading...", font=("Arial", 14))
        self.loading_label.pack(pady=10)
        self.update_idletasks()

    def hide_loading_buffer(self):
        """Hides the loading buffer from the GUI."""
        if hasattr(self, "loading_label"):
            self.loading_label.destroy()
            self.update_idletasks()

    def close_and_open_report(self, report_path):
        """Closes the GUI and opens the MD file."""
        self.destroy()
        webbrowser.open(f"file://{os.path.abspath(report_path)}")
