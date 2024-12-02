import os
from PIL import Image
from docx import Document
from pptx import Presentation
import pandas as pd
import fitz
import pdfplumber


def read_rules_file(rules_path, logging):
    """Read rules file and return a DataFrame."""
    try:
        if rules_path.endswith(('.xls', '.xlsx')):
            logging.info(f"Reading Excel file: {rules_path}")
            return pd.read_excel(rules_path)
        elif rules_path.endswith('.csv'):
            logging.info(f"Reading CSV file: {rules_path}")
            return pd.read_csv(rules_path)
        else:
            logging.error(f"Unsupported file format for rules file: {rules_path}")
            return None
    except Exception as e:
        logging.error(f"Error reading rules file: {e}")
        return None


def extract_text_from_pdf(pdf_path, logging):
    text = ""
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            text += page.get_text("text")
        if text.strip():
            logging.info("Text extracted using PyMuPDF.")
            return text
    except Exception as e:
        logging.error(f"Error with PyMuPDF: {e}")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text += page.extract_text()
        if text.strip():
            logging.info("Text extracted using pdfplumber.")
            return text
    except Exception as e:
        logging.error(f"Error with pdfplumber: {e}")

    logging.warning("No text found, attempting OCR...")
    return "No text extracted."


def extract_text_from_ppt(ppt_path, logging):
    presentation = Presentation(ppt_path)
    extracted_text = ""
    for slide_number, slide in enumerate(presentation.slides):
        extracted_text += f"Slide {slide_number + 1}:\n"
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                extracted_text += shape.text + "\n"

        table_text = extract_table_text(slide, logging)
        if table_text:
            extracted_text += "Table Content:\n" + table_text

        extracted_text += "\n"
    logging.info(f"Text extracted from PPT: {ppt_path}")
    return extracted_text


def extract_table_text(slide, logging):
    table_text = ""
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    table_text += cell.text.strip() + " "
                table_text += "\n"
    logging.debug("Table content extracted.")
    return table_text


def extract_text_from_docx(docx_path, logging):
    document = Document(docx_path)
    extracted_text = ""

    for para in document.paragraphs:
        extracted_text += para.text + "\n"

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                extracted_text += cell.text + " "
            extracted_text += "\n"

    logging.info(f"Text extracted from DOCX: {docx_path}")
    return extracted_text


def list_files_in_directory(directory_path, logging):
    if not os.path.isdir(directory_path):
        logging.error(f"The directory '{directory_path}' does not exist.")
        return []
    logging.info(f"Listing all files in directory: {directory_path}")
    all_files = []
    for root, _, files in os.walk(directory_path):
        for file in files:
            all_files.append(os.path.join(root, file))
    return all_files


def search_words_in_text(text, keywords, logging):
    found_words = {}
    for word in keywords:
        found_words[word] = word in text
    logging.debug(f"Search results for words: {found_words}")
    return found_words

def read_file(file_path, logging):
    logging.info(f"Attempting to read file: {file_path}")
    try:
        if file_path.endswith('.csv'):
            data = pd.read_csv(file_path)
            logging.info("File successfully read as CSV.")
        elif file_path.endswith(('.xls', '.xlsx')):
            data = pd.read_excel(file_path)
            logging.info("File successfully read as Excel.")
        else:
            logging.error("Unsupported file format. Please provide a CSV or Excel file.")
            raise ValueError("Unsupported file format.")
        return data
    except Exception as e:
        logging.error(f"Error reading file: {e}")
        raise