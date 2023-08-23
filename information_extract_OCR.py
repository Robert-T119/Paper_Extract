import os
import re
import pandas as pd
import requests
from tabulate import tabulate
from pdf2image import convert_from_path
import pytesseract
from PIL import Image

def extract_potential_dois(text):
    """Extract potential DOIs using a more general regex pattern."""
    doi_pattern = r"10\.\d{4,9}/\S+"
    matches = re.findall(doi_pattern, text)
    return matches

def clean_doi(doi):
    """Clean the extracted DOI by removing the trailing dot (if present)."""
    return doi.rstrip('.')

def extract_text_from_pdf_using_ocr(file_path):
    """Extract text from the first page of a PDF using OCR after converting PDF to image."""
    images = convert_from_path(file_path, dpi=300, first_page=1, last_page=1)
    if images:
        text = pytesseract.image_to_string(images[0])
        return text
    return ""

def extract_doi_from_pdf_using_ocr(file_path):
    """Extract DOIs from the first page of a PDF using OCR."""
    text = extract_text_from_pdf_using_ocr(file_path)
    dois = extract_potential_dois(text)
    return clean_doi(dois[0]) if dois else "N/A"

def abstract_from_inverted_index(inverted_index):
    word_positions = [(word, pos) for word, positions in inverted_index.items() for pos in positions]
    word_positions.sort(key=lambda x: x[1])
    abstract = ' '.join(word for word, pos in word_positions)
    return abstract

def get_paper_info(doi):
    if doi == "N/A":
        return ["N/A"] * 7
    
    base_url = "https://api.openalex.org/works/"
    full_doi = "https://doi.org/" + doi
    response = requests.get(base_url + full_doi)

    if response.status_code != 200:
        return ["N/A"] * 7

    data = response.json()

    authors = ', '.join([authorship['author']['display_name'] for authorship in data['authorships']]) if 'authorships' in data else "N/A"
    publication_date = data.get('publication_date', "N/A")
    title = data.get('title', "N/A")
    concepts = ', '.join([concept['display_name'] for concept in data['concepts']]) if 'concepts' in data else "N/A"
    
    abstract = abstract_from_inverted_index(data['abstract_inverted_index']) if 'abstract_inverted_index' in data else "N/A"
    referenced_works = ', '.join(data['referenced_works']) if 'referenced_works' in data else "N/A"
    related_works = ', '.join(data['related_works']) if 'related_works' in data else "N/A"

    return authors, publication_date, title, concepts, abstract, referenced_works, related_works

directory = '/Users/bohui/Projects/Paper_filter/Test_paper'

table_data = []
for filename in os.listdir(directory):
    if filename.endswith('.pdf'):
        file_path = os.path.join(directory, filename)
        doi = extract_doi_from_pdf_using_ocr(file_path)
        paper_info = get_paper_info(doi)
        table_data.append([filename, doi, *paper_info])

table_headers = ["File Name", "DOI", "Authors", "Publication Date", "Title", "Concepts", "Abstract", "Referenced Works", "Related Works"]
table = tabulate(table_data, headers=table_headers, tablefmt="pretty")
print(table)

df = pd.DataFrame(table_data, columns=table_headers)
try:
    df.to_excel("/Users/bohui/Projects/Paper_filter/Extracted_info.xlsx", index=False)
except Exception as e:
    print(f"Exception occurred: {e}")