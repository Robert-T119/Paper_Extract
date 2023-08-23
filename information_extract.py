import os
import re
import PyPDF2
import requests
import pandas as pd
from tabulate import tabulate
import pytesseract
from PIL import Image
from pdf2image import convert_from_path

def extract_doi_from_image(file_path):
    images = convert_from_path(file_path, last_page=1)
    if not images:
        return None
    text = pytesseract.image_to_string(images[0])
    # The refined regex pattern for DOI extraction
    doi_pattern = r"(10\.\d{4,9}/[-._;()/:A-Z0-9]+)"
    matches = re.findall(doi_pattern, text.upper())
    # If multiple DOIs are found, return the longest one (assuming it's the most complete)
    raw_doi = max(matches, key=len, default=None) if matches else None
    return refine_doi(raw_doi) if raw_doi else None

def refine_doi(doi):
    # Remove any trailing periods from the DOI
    return doi.rstrip('.')

def clean_doi(doi):
    # Use regex to extract valid DOI pattern
    doi_pattern = r"10[.][0-9]{4,}(?:[.][0-9]+)*\/(?:(?![\"&\'<>])\S)+"
    matches = re.findall(doi_pattern, doi)
    return matches[0] if matches else None

def extract_dois_from_pdf(file_path):
    with open(file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        dois = []
        for page in pdf_reader.pages:
            text = page.extract_text()
            doi_pattern = r"\b(10[.][0-9]{4,}(?:[.][0-9]+)*\/(?:(?![\"&\'<>])\S)+)\b"
            matches = re.findall(doi_pattern, text)
            if matches:
                # Clean and validate the DOI
                cleaned_doi = clean_doi(matches[0])
                if cleaned_doi:
                    dois.append(cleaned_doi)
        if dois:
            return dois[0]
    return extract_doi_from_image(file_path)

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

# directory = 'C:\\Users\\robert.tang\\Ceres\\Z-Model_fine_tuning\\Testing_paper'
directory = '/Users/bohui/Projects/Paper_Filter/Test_paper'

table_data = []
for filename in os.listdir(directory):
    if filename.endswith('.pdf'):
        file_path = os.path.join(directory, filename)
        doi = extract_dois_from_pdf(file_path)
        paper_info = get_paper_info(doi)
        table_data.append([filename, doi, *paper_info])

table_headers = ["File Name", "DOI", "Authors", "Publication Date", "Title", "Concepts", "Abstract", "Referenced Works", "Related Works"]
table = tabulate(table_data, headers=table_headers, tablefmt="pretty")
print(table)

df = pd.DataFrame(table_data, columns=table_headers)
df.to_excel("/Users/bohui/Projects/Paper_Filter/Extracted_info.xlsx", index=False)









