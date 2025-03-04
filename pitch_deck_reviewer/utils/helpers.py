import os
import webbrowser
import time

# Constants
OUTPUT_DIR = "converted_pdfs"
API_RETRY_DELAY = 5


# Utility Functions
def create_output_dir(output_dir=OUTPUT_DIR):
    """Creates the output directory if it does not exist."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)


def retry_with_delay(func, delay=API_RETRY_DELAY, retries=3):
    """Retries a function with a delay between attempts."""
    for attempt in range(retries):
        try:
            return func()
        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            time.sleep(delay)
    raise Exception("All retry attempts failed.")


def get_file_extension(file_path):
    """Returns the file extension of the given file path."""
    return os.path.splitext(file_path)[1].lower()


def is_supported_file(file_path):
    """Checks if the file is supported for conversion."""
    supported_extensions = [".ppt", ".pptx", ".doc", ".docx", ".xls", ".xlsx", ".pdf"]
    return get_file_extension(file_path) in supported_extensions


def open_file_in_browser(file_path):
    """Opens the file in the default web browser."""
    webbrowser.open(f"file://{os.path.abspath(file_path)}")
