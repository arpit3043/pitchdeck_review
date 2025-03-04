from converters.file_converter import FileConverter
from exctractors.pdf_extractor import PDFExtractor
from feedback.ai_feedback_provider import AIFeedbackProvider
from reviewers.pitch_deck_reviewer import PitchDeckReviewer
from gui.pitch_deck_gui import PitchDeckGUI
from config.config import API_KEY

if __name__ == "__main__":
    try:
        file_converter = FileConverter()
        pdf_extractor = PDFExtractor()
        feedback_provider = AIFeedbackProvider(API_KEY)
        pitch_deck_reviewer = PitchDeckReviewer(feedback_provider)
        app = PitchDeckGUI(file_converter, pdf_extractor, pitch_deck_reviewer)
        app.mainloop()
    except Exception as e:
        print(f"An error occurred: {e}")
