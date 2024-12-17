import streamlit as st
import pdfplumber
import pandas as pd
import openai
import json
import os
from dotenv import load_dotenv

# Load .env file
load_dotenv()

# Fetch API key
openai.api_key = os.getenv("OPENAI_API_KEY")




# Function to extract text from PDF
def extract_text_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        return "".join(page.extract_text() for page in pdf.pages)

# Function to query GPT-3.5 Turbo to extract key fields
def extract_invoice_fields(invoice_text):
    """
    Query GPT-3.5 Turbo to extract key invoice fields and return cleaned JSON.
    """
    prompt = f"""
    Extract the following key fields from the given invoice text:
    - Invoice Number
    - Invoice Date
    - Due Date
    - Items (Description, Quantity, Price)
    - Tax
    - Total Amount

    Invoice Text:
    {invoice_text}

    Output in JSON format:
    {{
        "Invoice Number": "INV-XXXX",
        "Invoice Date": "YYYY-MM-DD",
        "Due Date": "YYYY-MM-DD",
        "Items": [
            {{"Description": "Item 1", "Quantity": 1.0, "Price": 100.0}}
        ],
        "Tax": 10.0,
        "Total Amount": 110.0
    }}
    """
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "You are a helpful assistant for invoice data extraction."},
            {"role": "user", "content": prompt}
        ]
    )
    result = response['choices'][0]['message']['content']
    
    # Clean and validate JSON
    try:
        result = result.strip("```json").strip("```").strip()
        parsed_result = json.loads(result)
        return parsed_result  # Return as a valid dictionary
    except json.JSONDecodeError as e:
        print("Error decoding JSON:", e)
        return None

# Function to save extracted data to Excel
def save_to_excel(data, output_file):
    """
    Save structured invoice data to an Excel file.
    """
    try:
        # Extract line items and append other fields
        line_items = pd.DataFrame(data["Items"])
        line_items["Invoice Number"] = data.get("Invoice Number", "N/A")
        line_items["Invoice Date"] = data.get("Invoice Date", "N/A")
        line_items["Due Date"] = data.get("Due Date", "N/A")
        line_items["Tax"] = data.get("Tax", 0)
        line_items["Total"] = data.get("Total", data.get("Total Amount", 0))  # Check for both keys

        # Save to Excel
        line_items.to_excel(output_file, index=False)
        return True, output_file
    except Exception as e:
        return False, str(e)

# Streamlit UI
st.title("Invoice Extraction System")

# File uploader
uploaded_file = st.file_uploader("Upload Invoice PDF", type="pdf")

if uploaded_file:
    st.info("Extracting text from PDF...")
    try:
        # Step 1: Extract text from PDF
        pdf_text = extract_text_from_pdf(uploaded_file)
        st.success("Text extracted successfully!")
        st.text_area("Extracted Text", pdf_text, height=300)

        # Step 2: Extract fields using GPT-3.5 Turbo
        st.info("Extracting key fields using GPT-3.5 Turbo...")
        extracted_data = extract_invoice_fields(pdf_text)

        if extracted_data:
            st.success("Fields extracted successfully!")
            st.json(extracted_data)
            
            # Step 3: Save data to Excel
            output_file = "extracted_invoice_data.xlsx"
            success, message = save_to_excel(extracted_data, output_file)
            if success:
                st.success("Data saved successfully!")
                with open(output_file, "rb") as f:
                    st.download_button("Download Extracted Data as Excel", f, file_name="invoice_data.xlsx")
            else:
                st.error(f"Error saving data: {message}")
        else:
            st.error("Failed to extract data. Please try again.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
