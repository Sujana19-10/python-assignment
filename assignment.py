# -*- coding: utf-8 -*-
"""
Created on Sun Apr 14 10:14:28 2024

@author: baswa
"""

import os
import re
import pandas as pd
from docx import Document
import PyPDF2

def extract_information(file_path):
    email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    phone_regex = r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]'
    
    if file_path.endswith('.pdf'):
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ''
            for page_num in range(len(reader.pages)):
                text += reader.pages[page_num].extract_text()
    elif file_path.endswith('.docx'):
        doc = Document(file_path)
        paragraphs = [p.text for p in doc.paragraphs]
        text = '\n'.join(paragraphs)
    else:
        return None, None, None

    email = re.search(email_regex, text)
    phone = re.search(phone_regex, text)
    return email.group() if email else None, phone.group() if phone else None, text

def process_cv_directory(directory):
    data = []
    for filename in os.listdir(directory):
        if filename.endswith('.pdf') or filename.endswith('.docx'):
            file_path = os.path.join(directory, filename)
            email, phone, text = extract_information(file_path)
            data.append({'File': filename, 'Email': email, 'Phone': phone, 'Text': text})
    return data

def create_excel(data, output_path):
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False)

if __name__ == "__main__":
    input_directory = r'D:\Sample2'  # Update with the correct directory path
    output_file = 'output.xlsx'

    cv_data = process_cv_directory(input_directory)
    create_excel(cv_data, output_file)
