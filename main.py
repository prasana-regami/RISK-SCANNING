import os
import json
import logging
import argparse
from tqdm import tqdm

import pandas as pd
from extraction import (
    read_file,
    read_rules_file,
    extract_text_from_pdf,
    extract_text_from_ppt,
    extract_text_from_docx,
    list_files_in_directory,
    search_words_in_text,
)


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("process.log")    ]
)

PROCESSORS = {
    ".pdf": lambda file_path, logging: extract_text_from_pdf(file_path, logging),
    ".ppt": lambda file_path, logging: extract_text_from_ppt(file_path, logging),
    ".pptx": lambda file_path, logging: extract_text_from_ppt(file_path, logging),
    ".docx": lambda file_path, logging: extract_text_from_docx(file_path, logging),
    ".json": lambda file_path, logging: logging.info(f"Processing JSON: {file_path}") or json.dumps(json.load(open(file_path, "r")), indent=4),
    ".txt": lambda file_path, logging: logging.info(f"Processing TXT: {file_path}") or open(file_path, "r").read(),
}

def terminal_args():
    parser = argparse.ArgumentParser(
        description="Process files with specified rules.")
    parser.add_argument("-d", "--directory", required=True,
                        help="Path to the directory to check for files.")
    parser.add_argument("-r", "--rules", required=True,
                        help="Path to the rules file.")
    parser.add_argument("-o", "--output", required=True,
                        help="Path to the output directory.")

    args = parser.parse_args()

    # Check directory for files
    directory_path = args.directory
    files = list_files_in_directory(directory_path,logging)
    if not files:
        logging.warning(f"No files found in '{directory_path}'.")
        return None

    logging.info(f"Files found in '{directory_path}': {len(files)}")

    # Check rules file
    rules_path = args.rules
    if not os.path.exists(rules_path):
        logging.error(f"The rules file '{rules_path}' does not exist.")
        return None

    # Check or create output directory
    output_directory = args.output
    if not os.path.exists(output_directory):
        logging.warning(
            f"The output directory '{output_directory}' does not exist. Creating it now.")
        os.makedirs(output_directory)

    rules_df = read_rules_file(rules_path, logging)
    if rules_df is None:
        logging.error("Failed to load rules file.")
        return None

    return {"input": directory_path, "rules": rules_path, "output": output_directory, "files": files}


def process_file(file_path, logging):
    """Process a file based on its extension."""
    _, ext = os.path.splitext(file_path.lower())
    if ext in PROCESSORS:
        logging.info(f"Processing {ext.upper()[1:]}: {os.path.basename(file_path)}")
        return PROCESSORS[ext](file_path, logging)
    else:
        logging.warning(f"Skipping unsupported file type: {os.path.basename(file_path)}")
        return None
    
def save_results_to_excel(results, output_file):
    # Convert results to a DataFrame
    df = pd.DataFrame(results, columns=["File Name", "File Path", "Search Keyword", "Status"])
    df.to_excel(output_file, index=False)
    logging.info(f"Results saved to {output_file}")


def main():
    input_args = terminal_args()
    if not input_args:
        logging.error("Invalid input arguments. Exiting.")
        return

    directory_path = input_args.get("input")
    rules = input_args.get("rules")
    output_directory = input_args.get("output")
    files = input_args.get("files")

    try:
        data = read_file(rules, logging)
        logging.info("read the rules data")  # Preview the data
    except Exception as e:
        logging.error(f"Failed to process the file: {e}")

    all_results = []
    keyword_data = data.get('keywords')
    with tqdm(total=len(files), desc="Scanning", unit="file") as pbar:
        for file in files:
            file_path = os.path.join(directory_path, file)
            
            # Process file and extract text
            extracted_text = process_file(file_path, logging)

            if extracted_text:
                # Search for words in extracted text
                keywords = list(keyword_data) if not keyword_data.empty else []

                word_results = search_words_in_text(extracted_text, keywords, logging)
                
                # If any words are found, store the results
                for word in word_results:
                    all_results.append([file.split('\\')[-1], file_path, word, "matched"])

                # If no words are found, still log the result as not matched
                if not word_results:
                    for word in word_results:
                        all_results.append([file.split('\\')[-1], file_path, word, "not matched"])
                
                logging.info(f"{file} - Scanning completed")

            # Update the progress bar after scanning each file
            pbar.update(1)
    
    
    # Save the results to Excel
    output_file = os.path.join(output_directory, "search_results.xlsx")
    save_results_to_excel(all_results, output_file)


if __name__ == "__main__":
    main()
