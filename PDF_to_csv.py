import re
import csv
import os
from PyPDF2 import PdfReader
from PIL import Image
from io import BytesIO

# Function to extract text from PDF
def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text()
    return text

# Function to extract images from PDF starting from page 2
def extract_images_from_pdf(pdf_path, output_dir):
    reader = PdfReader(pdf_path)
    image_paths = []
    for page_num, page in enumerate(reader.pages[1:], start=2):  # Start from page 2
        if '/XObject' in page['/Resources']:
            x_objects = page['/Resources']['/XObject'].get_object()
            for obj_name in x_objects:
                obj = x_objects[obj_name]
                if obj['/Subtype'] == '/Image':
                    size = (obj['/Width'], obj['/Height'])
                    data = obj.get_data()
                    img = Image.open(BytesIO(data))
                    img_path = os.path.join(output_dir, f"image_page{page_num}_{obj_name[1:]}.png")
                    img.save(img_path)
                    image_paths.append(img_path)
    return image_paths

# Function to parse extracted text
def parse_text(text):
    data = []

    # Step-by-step pattern matching for each field
    id_name_pattern = re.compile(r'(\d+)\.\sনাম:\s(.+?)\n')
    voter_no_pattern = re.compile(r'ভাটার নং:\s(.+?)\n')
    father_pattern = re.compile(r'িপতা:\s(.+?)\n')
    mother_pattern = re.compile(r'মাতা:\s(.+?)\n')
    profession_pattern = re.compile(r'Ïপশা:\s(.+?),')
    DOB = re.compile(r'জĥ তািরখ:(\d{2}/\d{2}/\d{4})\n')
    address_pattern = re.compile(r'িঠকানা:\s(.+?),\s(.+?),\s(.+?)\n(.+?)\n')

    # Find all matches for each pattern
    ids_names = id_name_pattern.findall(text)
    voter_nos = voter_no_pattern.findall(text)
    fathers = father_pattern.findall(text)
    mothers = mother_pattern.findall(text)
    profession_dobs = profession_pattern.findall(text)
    addresses = address_pattern.findall(text)
    DOB = DOB.findall(text)

    # Print matches for verification
    print("IDs and Names:", ids_names)
    print("Voter Nos:", voter_nos)
    print("Fathers:", fathers)
    print("Mothers:", mothers)
    print("Professions:", profession_dobs)
    print("DOB:", DOB)
    print("Addresses:", addresses)

    # Combine matched data
    for i in range(len(ids_names)):
        entry = {
            'ID': ids_names[i][0],
            'Name': ids_names[i][1],
            'Voter No': voter_nos[i] if i < len(voter_nos) else None,
            'Father': fathers[i] if i < len(fathers) else None,
            'Mother': mothers[i] if i < len(mothers) else None,
            'Profession': profession_dobs[i] if i < len(profession_dobs) else None,
            'DOB': DOB[i] if i < len(DOB) else None,
            'Address': f"{addresses[i][0]}, {addresses[i][1]}, {addresses[i][2]}" if i < len(addresses) else None,
            'Image': None  # Placeholder for image path
        }
        data.append(entry)

    return data

# Use the relative path of the PDF file
pdf_path = r'Sample PDF to CSV.pdf'  # Path to the PDF file
output_dir = r'images'  # Directory to save extracted images

# Ensure output directory exists
os.makedirs(output_dir, exist_ok=True)

# Extract text from PDF
extracted_text = extract_text_from_pdf(pdf_path)
print("Extracted Text:\n", extracted_text)  # Print extracted text for verification

# Extract images from PDF starting from page 2
image_paths = extract_images_from_pdf(pdf_path, output_dir)
print("Extracted Images:", image_paths)

# Parse text data
parsed_data = parse_text(extracted_text)
# print("Parsed Data:\n", parsed_data)  # Print parsed data for verification

for i in range(min(len(parsed_data), len(image_paths))):
    parsed_data[i]['Image'] = image_paths[i]

# Write to CSV in Bengali
csv_file = 'ladies_info_bangla_with_images.csv'

try:
    with open(csv_file, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=['ID', 'Name', 'Voter No', 'Father', 'Mother', 'Profession', 'DOB', 'Address', 'Image'])
        writer.writeheader()
        for entry in parsed_data:
            writer.writerow(entry)
    print(f"CSV file '{csv_file}' created successfully.")
except PermissionError as e:
    print(f"PermissionError: {e}. Ensure the file is not open in any other program and you have the correct permissions.")
