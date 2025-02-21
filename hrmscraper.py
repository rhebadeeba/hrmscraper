import requests
import PyPDF2
import re
from io import BytesIO
import pandas as pd
import os
import spacy

nlp = spacy.load('en_core_web_sm')

def extract_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        return text
    
def extract_page_range(input_pdf_path, start_page, end_page):
    with open(input_pdf_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        text = ""
        for page_number in range(start_page - 1, end_page):
            text += reader.pages[page_number].extract_text() + "\n"
        return text

def is_sj_related(paragraph, keywords):
    paragraph_lower = paragraph.lower()
    return any(keyword in paragraph_lower for keyword in keywords)

def is_climate_related(paragraph, keywords):
    paragraph_lower = paragraph.lower()
    return any(keyword in paragraph_lower for keyword in keywords)

def filter_sj_related_paragraphs(paragraphs, keywords):
    filtered_paragraphs = []
    for paragraph in paragraphs:
        lines = paragraph.split('\n')
        filtered_lines = []
        for line in lines:
            if is_header(line):
                continue
            filtered_lines.append(line)
        if len(filtered_lines) < 3:             # check how many lines are in paragraph
            continue                            # filter out if less than 3 lines?
        filtered_paragraph = " ".join(filtered_lines).strip()
        if is_sj_related(filtered_paragraph, keywords) and not is_climate_related(filtered_paragraph, climate_keywords):
            filtered_paragraphs.append(filtered_paragraph)
    return filtered_paragraphs

def is_header(phrase):
    if len(phrase.split()) < 4:
        return True
    
    if phrase.isupper():
        return True
    
    if "Vertex 2023 Corporate Responsibility Report" in phrase:
        return True
    
    if "2022 Inclusion, Diversity and Equity at Vertex Factsheet" in phrase:
        return True
    
    if "Intel Corporate Responsibility Report" in phrase:
        return True
    
    if "Reference Indices Key Performance Indicator" in phrase:
        return True
    
    if "All Rights Reserved" in phrase:
        return True
    
    if "appendix" in phrase.lower():
        return True
    
    if "Engaging and Developing Employees" in phrase:
        return True
    
    if "HELPING PEOPLE THRIVE" in phrase:
        return True
    
    if "|     Our people     |" in phrase:
        return True
    
    if "|" in phrase:
        return True
    
    if "Grant Thornton 2023 ESG Report" in phrase:
        return True
    
    
    header_patterns = [
        r'^[A-Z\s]+$',
        r'^\d+(\.\d+)*\s+'
    ]
    for pattern in header_patterns:
        if re.match(pattern, phrase):
            return True

    return False
    
def search_keywords_in_paragraphs(text, keywords):
    paragraphs = re.split(r'\.\s*\n|\n{2,}', text)
    results = []
    for paragraph in paragraphs:
        for keyword in keywords:
            if (re.search(rf'\b{keyword}\b', paragraph, re.IGNORECASE)):
                results.append(paragraph.replace('\n', ' ').strip())
                break
    return results

def remove_illegal_characters(text):
    return re.sub('[\x00-\x1F\x7F-\x9F]', '', text)

def add_to_spreadsheet(results, filename):
    cleaned_results = [[url, remove_illegal_characters(paragraph)] for url, paragraph in results]

    df = pd.DataFrame(cleaned_results, columns=['Document', 'Text'])    # create dataframe from list of paragraphs
    df.to_excel(filename, index=False, engine='openpyxl')                      # save to excel


urls = ['BIIB 10k 2018.pdf', 'VRTX 10k 2018.pdf', 'JNJ 10k 2018.pdf', 'ABT 10k 2018.pdf', 
        'ELY 10k 2018.pdf', 'BAC 10k 2018.pdf', 'PARA 10k 2018.pdf', 'INTC 10k 2018.pdf', 
        'DAL 10k 2018.pdf', 'CRM 10k 2018.pdf', 'MAR 10k 2018.pdf', 'NOC 10k 2018.pdf', 
        'GIS 10k 2018.pdf', 'EXPE 10k 2018.pdf', 'JPM 10k 2018.pdf', 'MSFT 10k 2018.pdf', 
        'ADP 10k 2018.pdf', 'AXP 10k 2018.pdf', 'COF 10k 2018.pdf', 'CL 10k 2018.pdf', 
        'GM 10k 2018.pdf', 'BMS 10k 2018.pdf', 'HPE 10k 2018.pdf', 
        'KHC 10k 2018.pdf', 'VZ 10k 2018.pdf', 'ZTS 10k 2018.pdf', 'SYF 10k 2018.pdf', 
        'GILD 10k 2018.pdf', 'ABBV 10k 2018.pdf', 'ACN 10k 2018.pdf', 'ALLY 10k 2018.pdf', 
        'BAH 10k 2018.pdf', 'CAH 10k 2018.pdf', 'ELAN 10k 2018.pdf', 'NDAQ 10k 2018.pdf', 
        'MCO 10k 2018.pdf', 'MS 10k 2018.pdf', 'PNC 10k 2018.pdf', 'PRU 10k 2018.pdf']
filename = '10k data.xlsx'
results = []
keywords = ['employee']
# 'justice', 'trust', 'respect', 'diversity', 'equity', 'inclusion', 'non-discrimination', 
#                 'equal opportunity', 'pay equity', 'workplace accessibility', 'harassment', 
#                 'employee', 'balance', 'inclusive leadership', 'advocacy', 
#                 'transparency', 'bias training', 'equitable', 'empowerment', 'DEI', 'inclusive', 
#                 'people', 'human', 'talent', 'hire', 'hiring', 'compensate', 'compensation', 'train', 'training', 'safety',
#                 'reimburse'
climate_keywords = ['climate', 'sustainability', 'renewable', 'environment', 'green', 'page', 'contents', 'table']

for pdf_name in urls:
    input_pdf = pdf_name
    text = extract_from_pdf(pdf_name)
    
    paragraphs = re.split(r'\.\s*\n|\n{2,}', text)
    keyword_paragraphs = filter_sj_related_paragraphs(paragraphs, keywords)
    print(f'{pdf_name} done')
    
    for paragraph in keyword_paragraphs:
        results.append([pdf_name, paragraph])
    
add_to_spreadsheet(results, filename)