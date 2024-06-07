import pypandoc
import os
import fitz
import pandas as pd
from bs4 import BeautifulSoup
from azure.cognitiveservices.vision.computervision import ComputerVisionClient
from azure.cognitiveservices.vision.computervision.models import OperationStatusCodes
from msrest.authentication import CognitiveServicesCredentials
import extract_msg
import time
from dotenv import load_dotenv
load_dotenv()

#Set up cognitive credentials
subscription_key = os.getenv('subscription_key')
endpoint = os.getenv('endpoint')
computervision_client = ComputerVisionClient(endpoint, CognitiveServicesCredentials(subscription_key))


def extract_doc(file_name):
    output = pypandoc.convert_file(file_name, 'rst')
    return output


def extract_text_from_txt(file_path):
    """Extract text from a txt file"""
    with open(file_path, 'r') as txt_file:
        txt_text = txt_file.read()
    return txt_text

def extract_text_from_csv(file_path):
    """Extract text from a CSV file."""
    try:
        df = pd.read_csv(file_path)
        text = df.to_string(index=False)
        return text
    except Exception as e:
        print(f"Error extracting text from CSV: {e}")
        return ""

def extract_text_from_xlsx(file_path):
    """Extract text from an XLSX file."""
    try:
        df = pd.read_excel(file_path)
        text = df.to_string(index=False)
        return text
    except Exception as e:
        print(f"Error extracting text from XLSX: {e}")
        return ""

def extract_text_from_html(file_path):
    """Extract text from an HTML file."""
    try:
        with open(file_path, 'r') as html_file:
            soup = BeautifulSoup(html_file, 'html.parser')
            text = soup.get_text()
        return text
    except Exception as e:
        print(f"Error extracting text from HTML: {e}")
        return ""



def extract_pdf_text(file_path):
    """Extract text from a PDF file."""
    doc = fitz.open(file_path)
    full_text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        full_text += text
    return full_text

def convert_pdf_to_images(file_path):
    """Convert PDF pages to images."""
    doc = fitz.open(file_path)
    image_paths = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap()
        image_path = f"page_{page_num + 1}.png"
        pix.save(image_path)
        image_paths.append(image_path)
    return image_paths

def extract_text_from_image(image_path):
    """Extract text from an image file using Azure Vision OCR."""
    try:
        with open(image_path, "rb") as image_stream:
            ocr_result = computervision_client.read_in_stream(image_stream, raw=True)
        
        operation_location = ocr_result.headers["Operation-Location"]
        operation_id = operation_location.split("/")[-1]

        while True:
            result = computervision_client.get_read_result(operation_id)
            if result.status not in ['notStarted', 'running']:
                break
            time.sleep(1)

        if result.status == OperationStatusCodes.succeeded:
            text = ""
            for read_result in result.analyze_result.read_results:
                for line in read_result.lines:
                    text += line.text + " "
            return text
        else:
            print("Sorry, the image quality is not sufficient for text extraction. Please try again with a clearer image.")
            return ""
    except Exception as e:
        print("Image is invalid for text extraction.")
        return ""

def is_text_based_pdf(file_path):
    """Check if a PDF file is text-based or scanned."""
    doc = fitz.open(file_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        if text.strip():
            return True
    return False

def process_pdf(file_path):
    """Process the PDF file to extract text."""
    if not file_path.lower().endswith('.pdf'):
        raise ValueError("The provided file is not a PDF.")

    if is_text_based_pdf(file_path):
        print("The PDF is text-based. Extracting text...")
        return extract_pdf_text(file_path)
    else:
        print("The PDF contains scanned images. Performing OCR...")
        try:
            image_paths = convert_pdf_to_images(file_path)
            full_text = ""
            for image_path in image_paths:
                text = extract_text_from_image(image_path)
                full_text += text
                os.remove(image_path)
            return full_text
        except Exception as e:
            print("Sorry, the image quality is not sufficient for text extraction. Please try again with a clearer image.")
            return ""
        
        

def extract_text_from_msg(file_path):
    """Extract text content from an MSG file."""
    try:
        msg = extract_msg.Message(file_path)
        return {
            "Subject": msg.subject,
            "From": msg.sender,
            "To": msg.to,
            "Date": msg.date,
            "Body": msg.body,
            "Attachments": [attachment.longFilename if attachment.longFilename else attachment.shortFilename for attachment in msg.attachments]
        }
    except Exception as e:
        print(f"Error extracting details from MSG: {e}")
        return {"error": "Invalid attachment or MSG file."}
