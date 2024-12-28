import streamlit as st
import pytesseract
from PIL import Image
import pandas as pd
import openai
import os
import json
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")

# Configure Tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

st.title("Fine-Tuned Invoice Extraction System")
st.write("Upload an invoice image and extract structured data dynamically, including accurate line-item details.")

# Function to extract text from image using Tesseract
def extract_text_from_image(image):
    return pytesseract.image_to_string(image, lang="eng")

# Function to parse invoice data dynamically with fine-tuning
def parse_invoice_data(invoice_text):
    """
    Use GPT-3.5 Turbo to extract invoice data dynamically and provide fine-tuned, accurate line-item details.
    """
    prompt = f"""
    Analyze the following invoice text and extract structured data dynamically. Ensure accurate line-item extraction with dynamically inferred column headers and values. Focus on:

    1. Invoice metadata: Invoice Number, Invoice Date, Due Date.
    2. Vendor details: Company Name, Address.
    3. Client details: Company Name, Address.
    4. Line items: Identify all relevant column headers (e.g., 'Description', 'Quantity', 'Rate', 'Subtotal', 'Unit Price', etc.) dynamically based on the invoice text, and map the associated values for each line.
    5. Summary: Extract key fields like 'Subtotal', 'Tax', 'Total' dynamically from the text.

    Ensure that:
    - All column headers for line items are dynamically inferred based on the text.
    - The output is well-structured JSON with clear, dynamically inferred keys for all sections.

    Invoice Text:
    {invoice_text}

    Example Output:
    {{
        "Invoice Metadata": {{
            "Invoice Number": "12345",
            "Invoice Date": "YYYY-MM-DD",
            "Due Date": "YYYY-MM-DD"
        }},
        "Vendor Details": {{
            "Company Name": "Vendor Company",
            "Address": "Vendor Address"
        }},
        "Client Details": {{
            "Company Name": "Client Company",
            "Address": "Client Address"
        }},
        "Line Items": [
            {{
                "Column Header 1": "Value 1",
                "Column Header 2": "Value 2",
                ...
            }}
        ],
        "Summary": {{
            "Field 1": "Value 1",
            "Field 2": "Value 2",
            ...
        }}
    }}
    """
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "You are an advanced invoice data extraction assistant, focused on accuracy."},
            {"role": "user", "content": prompt}
        ]
    )
    try:
        extracted_data = response['choices'][0]['message']['content']
        cleaned_data = extracted_data.strip("```json").strip("```").strip()
        return json.loads(cleaned_data)
    except Exception as e:
        st.error(f"Error parsing invoice data: {e}")
        return None

# Function to save fully dynamic data to Excel with fine-tuned line items
def save_to_dynamic_excel(data, output_file):
    """
    Save extracted invoice data to an Excel file dynamically, including fine-tuned and accurate line-item details.
    """
    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Save Invoice Metadata
            if "Invoice Metadata" in data:
                metadata = pd.DataFrame.from_dict(data["Invoice Metadata"], orient="index", columns=["Value"])
                metadata.reset_index(inplace=True)
                metadata.columns = ["Field", "Value"]
                metadata.to_excel(writer, sheet_name="Invoice_Metadata", index=False)

            # Save Vendor and Client Details
            details = []
            if "Vendor Details" in data:
                for key, value in data["Vendor Details"].items():
                    details.append({"Category": "Vendor", "Field": key, "Value": value})
            if "Client Details" in data:
                for key, value in data["Client Details"].items():
                    details.append({"Category": "Client", "Field": key, "Value": value})
            if details:
                details_df = pd.DataFrame(details)
                details_df.to_excel(writer, sheet_name="Vendor_Client_Details", index=False)

            # Save Line Items
            if "Line Items" in data:
                line_items_df = pd.DataFrame(data["Line Items"])
                line_items_df.to_excel(writer, sheet_name="Line_Items", index=False)

            # Save Summary
            if "Summary" in data:
                summary_df = pd.DataFrame.from_dict(data["Summary"], orient="index", columns=["Value"])
                summary_df.reset_index(inplace=True)
                summary_df.columns = ["Field", "Value"]
                summary_df.to_excel(writer, sheet_name="Summary", index=False)

        return True, output_file
    except Exception as e:
        return False, str(e)

# Upload image
uploaded_file = st.file_uploader("Upload Invoice Image", type=["png", "jpg", "jpeg"])
if uploaded_file:
    try:
        # Display uploaded image
        st.image(uploaded_file, caption="Uploaded Invoice", use_column_width=True)

        # Extract text from image
        st.info("Extracting text from image...")
        image = Image.open(uploaded_file)
        invoice_text = extract_text_from_image(image)
        st.text_area("Extracted Text", invoice_text, height=300)

        # Parse invoice data dynamically
        st.info("Parsing invoice data dynamically...")
        parsed_data = parse_invoice_data(invoice_text)

        if parsed_data:
            st.json(parsed_data)
            st.success("Invoice data parsed successfully!")

            # Save parsed data to Excel
            output_file = "fine_tuned_invoice_data.xlsx"
            success, message = save_to_dynamic_excel(parsed_data, output_file)
            if success:
                with open(output_file, "rb") as f:
                    st.download_button(
                        label="Download Extracted Data as Excel",
                        data=f,
                        file_name="invoice_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error(f"Error saving to Excel: {message}")
        else:
            st.error("Failed to parse invoice data.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
