import os
import json
import logging
import argparse
from tqdm import tqdm
from datetime import datetime

import pandas as pd
from extraction import (
    read_file,
    read_excel_file,
    read_rules_file,
    extract_text_from_pdf,
    extract_text_from_ppt,
    extract_text_from_docx,
    list_files_in_directory,
    search_words_in_text,
    extract_text_from_log,
    extract_text_from_outlook,
    extract_text_from_eml,
    extract_text_from_yaml,
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
    ".csv": lambda file_path, logging: read_excel_file(file_path, logging),
    ".xls": lambda file_path, logging: read_excel_file(file_path, logging), 
    ".xlsx": lambda file_path, logging: read_excel_file(file_path, logging),
    ".log" : lambda file_path, logging : extract_text_from_log(file_path,logging),
    ".msg": lambda file_path, logging: extract_text_from_outlook(file_path, logging),
    ".eml": lambda file_path, logging: extract_text_from_eml(file_path, logging),
    ".yaml": lambda file_path, logging: extract_text_from_yaml(file_path, logging),
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

    directory_path = args.directory
    files = list_files_in_directory(directory_path,logging)
    if not files:
        logging.warning(f"No files found in '{directory_path}'.")
        return None

    logging.info(f"Files found in '{directory_path}': {len(files)}")

    rules_path = args.rules
    if not os.path.exists(rules_path):
        logging.error(f"The rules file '{rules_path}' does not exist.")
        return None

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
    df = pd.DataFrame(results, columns=["File Name", "File Path", "Search Keyword", "Status"])
    df.to_excel(output_file, index=False)
    logging.info(f"Results saved to {output_file}")

      
def create_json_report(input_path, output_path, rules_path, processed_files, matched_files, unmatched_files, file_word_counts):
    """
    Create a JSON report after processing files.

    Args:
        input_path (str): Path to the input directory or file.
        output_path (str): Path to the output directory or file.
        rules_path (str): Path to the rules directory or file.
        processed_files (list): List of all processed file paths.
        matched_files (list): List of matched file paths.
        unmatched_files (list): List of unmatched file paths.
        file_word_counts (dict): Dictionary containing word counts for matched and unmatched keywords in each file.
    """
    report = {
        "input_path": input_path,
        "output_path": output_path,
        "rules_path": rules_path,
        "processed_files_count": len(processed_files),
        "matched_files_count": len(matched_files),
        "unmatched_files_count": len(unmatched_files),
        "processed_files": processed_files,
        "matched_files": matched_files,
        "unmatched_files": unmatched_files,
        "file_word_counts": file_word_counts,  
    }

    report_file = os.path.join(output_path, "processing_report.json")

    with open(report_file, 'w') as json_file:
        json.dump(report, json_file, indent=4)

    logging.info(f"Report saved to {report_file}")

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
        logging.info("read the rules data")  
    except Exception as e:
        logging.error(f"Failed to process the file: {e}")

    all_results = []
    keyword_data = data.get('keywords')
    
    processed_files = set() 
    matched_files = set()  
    unmatched_files = set()  
    file_word_counts = {}  
    
    with tqdm(total=len(files), desc="Scanning", unit="file") as pbar:
        for file in files:
            file_path = os.path.join(directory_path, file)
            
            extracted_text = process_file(file_path, logging)
            processed_files.add(file_path)  

            if extracted_text:
                keywords = list(keyword_data) if not keyword_data.empty else []

                word_results = search_words_in_text(extracted_text, keywords, logging)
                
                matched_count = 0
                unmatched_count = 0
                file_matched = False  
                
                for word, result in word_results.items():
                    if result == True:
                        all_results.append([file.split('\\')[-1], file_path, word, "matched"])
                        matched_files.add(file_path)  
                        matched_count += 1
                        file_matched = True  
                    elif result == False:
                        all_results.append([file.split('\\')[-1], file_path, word, "not matched"])
                        unmatched_count += 1

                
                file_word_counts[os.path.basename(file_path)] = {
                    "matched_count": matched_count,
                    "unmatched_count": unmatched_count
                }
                
                if not file_matched:
                    unmatched_files.add(file_path) 
                
                logging.info(f"{file} - Scanning completed")

            pbar.update(1)
    
   
    date_time_string = datetime.now().strftime("%Y-%m-%d_%HH_%MM")
    output_file = os.path.join(output_directory, f"Output_{date_time_string}.xlsx")
    save_results_to_excel(all_results, output_file)

    create_json_report(
        input_path=directory_path,
        output_path=output_directory,
        rules_path=rules,
        processed_files=list(processed_files),
        matched_files=list(matched_files),
        unmatched_files=list(unmatched_files),
        file_word_counts=file_word_counts 
    )


if __name__ == "__main__":
    main()
