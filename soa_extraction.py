import os
import sys
import json
import tempfile
import extract_msg
import pandas as pd
from docx import Document
from pydantic import BaseModel, Field
from typing import List, Optional
from dotenv import load_dotenv
from google import genai
from google.genai import types


def get_gemini_api_key():
    return os.getenv("GEMINI_API_KEY", "").strip()

def _load_environment():
    # Attempt to load .env from current dir, exe dir, or bundled dir
    dirs_to_check = [os.getcwd()]
    if getattr(sys, 'frozen', False):
        dirs_to_check.append(os.path.dirname(sys.executable))
        if hasattr(sys, '_MEIPASS'):
            dirs_to_check.append(sys._MEIPASS)
    dirs_to_check.append(os.path.dirname(os.path.abspath(__file__)))
    
    for d in dirs_to_check:
        env_file = os.path.join(d, ".env")
        if os.path.exists(env_file):
            # Never let a local file silently replace a securely injected
            # process-level secret.
            load_dotenv(env_file, override=False)
            break

_load_environment()

# Pydantic Schemas for Structured Output
class SOALineItem(BaseModel):
    invoice_number: str = Field(description="Exactly from SOA. NOTE: If the SOA uses 'Document Number' or 'Reference Number' instead of 'Invoice Number', extract that as the invoice_number.")
    invoice_date: Optional[str] = Field(None, description="in DD/MM/YYYY format")
    invoice_value: Optional[float] = Field(None, description="The 'Amount (Original)' exactly from SOA with 2 decimal points. If both Original Amount and Open Amount are present and different, always use the Original Amount here.")
    open_amount: Optional[float] = Field(None, description="Open Amount (if available)")
    source_document: str = Field(description="The exact name of the file (e.g. 'invoice.pdf', 'table.xlsx') or 'Email Body' where this line item was found.")

class ExtractedSOA(BaseModel):
    line_items: List[SOALineItem] = Field(description="List of all transactions extracted from the SOA")

DIRECT_MIME_TYPES = {
    '.html': 'text/html',
    '.htm': 'text/html',
    '.css': 'text/css',
    '.txt': 'text/plain',
    '.log': 'text/plain',
    '.xml': 'text/xml',
    '.csv': 'text/csv',
    '.rtf': 'text/rtf',
    '.js': 'text/javascript',
    '.json': 'application/json',
    '.pdf': 'application/pdf',
    '.bmp': 'image/bmp',
    '.jpeg': 'image/jpeg',
    '.jpg': 'image/jpeg',
    '.png': 'image/png',
    '.webp': 'image/webp'
}

def extract_msg_contents(msg_path):
    print(f"Reading MSG File: {msg_path}")
    msg = extract_msg.Message(msg_path)
    
    # 1. Extract Email Body & Metadata
    email_sender = msg.sender
    email_subject = msg.subject
    email_date = msg.date
    
    # Prefer HTML Body over Plain Text to preserve tables, styling, and structure
    email_body_html = msg.htmlBody.decode('utf-8', errors='ignore') if msg.htmlBody else (msg.body or "")
    
    print(f"Subject: {email_subject}")
    print(f"From: {email_sender}")
    print(f"Date: {email_date}")
    print(f"Body (HTML) length: {len(email_body_html)} chars")
    
    # Create temp directory for attachments
    attachments_extracted = []
    
    # 2. Extract Attachments
    for att in msg.attachments:
        att_name = att.longFilename or att.shortFilename
        if not att_name:
            continue
        print(f"Found attachment: {att_name}")
        
        # Determine file type
        _, ext = os.path.splitext(att_name.lower())
        
        # Read raw content
        content_bytes = att.data
        
        attachment_info = {
            "filename": att_name,
            "extension": ext,
            "bytes_length": len(content_bytes)
        }
        
        if ext in ['.xlsx', '.xls', '.xlsm']:
            # Convert Excel directly to CSV bytes to send natively to Gemini as text/csv
            with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as temp_file:
                temp_file.write(content_bytes)
                temp_path = temp_file.name
            
            try:
                excel_file = pd.ExcelFile(temp_path)
                csv_content = ""
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(temp_path, sheet_name=sheet_name)
                    df_cleaned = df.copy()
                    for col in df_cleaned.columns:
                        if pd.api.types.is_datetime64_any_dtype(df_cleaned[col]):
                            df_cleaned[col] = df_cleaned[col].dt.strftime('%Y-%m-%d %H:%M:%S')
                    csv_content += f"--- SHEET: {sheet_name} ---\n"
                    csv_content += df_cleaned.to_csv(index=False) + "\n\n"
                
                attachment_info["mime_type"] = "text/csv"
                attachment_info["raw_bytes"] = csv_content.encode('utf-8')
                attachment_info["send_via_bytes"] = True
                excel_file.close() # explicitly close to free file handle on Windows
                print(f"   Successfully converted Excel '{att_name}' to CSV and routing directly to Gemini natively.")
            except Exception as e:
                print(f"   Error parsing Excel file {att_name}: {e}")
            finally:
                if os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except Exception as ex:
                        print(f"   Warning: Could not remove temp file {temp_path}: {ex}")
                    
        elif ext in DIRECT_MIME_TYPES:
            # Handle native Gemini filetypes (Text, Application, Images) directly via bytes
            attachment_info["mime_type"] = DIRECT_MIME_TYPES[ext]
            attachment_info["raw_bytes"] = content_bytes
            attachment_info["send_via_bytes"] = True
            print(f"   Native attachment '{att_name}' ({DIRECT_MIME_TYPES[ext]}) ready for direct Part.from_bytes LLM processing.")
            
        elif ext == '.docx':
            # Handle Word documents (.docx) by converting them into text
            with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as temp_file:
                temp_file.write(content_bytes)
                temp_path = temp_file.name
                
            try:
                doc = Document(temp_path)
                fullText = []
                for para in doc.paragraphs:
                    fullText.append(para.text)
                # Handle tables inside the word document
                for table in doc.tables:
                    for row in table.rows:
                        row_text = [cell.text for cell in row.cells]
                        fullText.append(" | ".join(row_text))
                
                attachment_info["word_text"] = "\n".join(fullText)
                print(f"   Successfully extracted text from Word Document '{att_name}'.")
            except Exception as e:
                print(f"   Error parsing Word Document {att_name}: {e}")
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                    
        else:
            # Other attachments
            print(f"   Generic attachment '{att_name}' extracted.")
            
        attachments_extracted.append(attachment_info)
        
    return {
        "subject": email_subject,
        "sender": email_sender,
        "date": email_date,
        "body_html": email_body_html, # Export HTML instead of plain text body
        "attachments": attachments_extracted
    }

def extract_standalone_file_contents(file_path):
    """
    Parses a standalone document/attachment (PDF, Excel, Word, Image, CSV, etc.)
    without requiring an email wrapper.
    """
    filename = os.path.basename(file_path)
    _, ext = os.path.splitext(filename.lower())
    print(f"Reading Standalone File: {filename} (Type: {ext})")
    
    with open(file_path, 'rb') as f:
        content_bytes = f.read()
        
    attachment_info = {
        "filename": filename,
        "extension": ext,
        "bytes_length": len(content_bytes)
    }
    
    if ext in ['.xlsx', '.xls', '.xlsm']:
        try:
            excel_file = pd.ExcelFile(file_path)
            csv_content = ""
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                df_cleaned = df.copy()
                for col in df_cleaned.columns:
                    if pd.api.types.is_datetime64_any_dtype(df_cleaned[col]):
                        df_cleaned[col] = df_cleaned[col].dt.strftime('%Y-%m-%d %H:%M:%S')
                csv_content += f"--- SHEET: {sheet_name} ---\n"
                csv_content += df_cleaned.to_csv(index=False) + "\n\n"
            attachment_info["mime_type"] = "text/csv"
            attachment_info["raw_bytes"] = csv_content.encode('utf-8')
            attachment_info["send_via_bytes"] = True
            excel_file.close()
        except Exception as e:
            print(f"   Error parsing Excel file {filename}: {e}")
            
    elif ext in DIRECT_MIME_TYPES:
        attachment_info["mime_type"] = DIRECT_MIME_TYPES[ext]
        attachment_info["raw_bytes"] = content_bytes
        attachment_info["send_via_bytes"] = True
        
    elif ext == '.docx':
        try:
            doc = Document(file_path)
            fullText = []
            for para in doc.paragraphs:
                fullText.append(para.text)
            for table in doc.tables:
                for row in table.rows:
                    row_text = [cell.text for cell in row.cells]
                    fullText.append(" | ".join(row_text))
            attachment_info["word_text"] = "\n".join(fullText)
        except Exception as e:
            print(f"   Error parsing Word Document {filename}: {e}")
            
    return {
        "subject": f"Direct Upload: {filename}",
        "sender": "Direct Upload",
        "date": "",
        "body_html": f"<p>Directly uploaded statement of account attachment: {filename}</p>",
        "attachments": [attachment_info]
    }

def call_llm_for_extraction(parsed_email: dict) -> tuple[ExtractedSOA, dict[str, int]]:
    api_key = get_gemini_api_key()
    client = genai.Client(vertexai=True, api_key=api_key)
    
    system_prompt = """
    You are an AI financial auditor. Your task is to process statement of account (SOA) requests from emails and attachments.
    Extract the following information for EVERY line item / invoice:
    A) Invoice Numbers: exactly from SOA. NOTE: If the SOA uses 'Document Number' or 'Reference Number' instead of 'Invoice Number', extract that as the Invoice Number.
    B) Invoice Date: in DD/MM/YYYY format
    C) Invoice Value/ Amount (Original): exactly from SOA with 2 decimal points. CRITICAL RULE: If both Amount (Original) and Open Amount are available and they differ (due to partial payments, etc.), ALWAYS use the value from Amount (Original) for this field.
    D) Open Amount: if available
    E) Source Document: The exact name of the attachment (e.g. 'invoice.pdf', 'table.xlsx', 'image001.png') or 'Email Body' where you found this row.
    """
    
    user_prompt_text = f"### EMAIL CONTEXT:\nSender: {parsed_email['sender']}\nSubject: {parsed_email['subject']}\n\n### EMAIL BODY (HTML):\n{parsed_email['body_html']}\n"
    
    # 1. Add Word Text context (Excel is now sent natively as text/csv)
    for att in parsed_email['attachments']:
        if "word_text" in att:
            user_prompt_text += f"\n### WORD ATTACHMENT ({att['filename']}):\n"
            user_prompt_text += att['word_text']
            
    contents = [user_prompt_text]
    
    # 2. Automatically stream all native files directly into Gemini context
    for att in parsed_email['attachments']:
        if att.get("send_via_bytes"):
            # Check the actual content to see if it is a valid image or just a tiny pixel wrapper
            raw_data = att["raw_bytes"]
            ext_type = att["mime_type"]
            
            # For images, verify they have basic signatures or skip them (often tracking pixels are 43 bytes)
            if ext_type.startswith("image/"):
                # 1. Skip social media logos and tracking pixels (usually < 5KB)
                if len(raw_data) < 5000:
                    print(f"   Skipping logo/icon {att.get('filename')} (Size: {len(raw_data)} bytes) to ignore signatures and tracking pixels.")
                    continue
                
                # 2. Mathematically verify the image isn't corrupt (this is what caused the Gemini 400 error!)
                try:
                    import io
                    from PIL import Image
                    img = Image.open(io.BytesIO(raw_data))
                    img.verify() # Validates the image headers
                except Exception as e:
                    print(f"   Skipping Corrupt Image {att.get('filename')} (Size: {len(raw_data)} bytes). Error: {e}")
                    continue
                    
            size_bytes = len(raw_data)
            is_pdf = (ext_type == 'application/pdf')
            # Thresholds based on Gemini docs: 50MB for PDF, 100MB for everything else
            # (We set it slightly below to be safe)
            threshold = 48 * 1024 * 1024 if is_pdf else 98 * 1024 * 1024
            
            if size_bytes > threshold:
                print(f"   File {att.get('filename')} is very large ({size_bytes/1024/1024:.2f} MB). Uploading via Gemini File API...")
                _, ext = os.path.splitext(att.get('filename', ''))
                with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                    tmp.write(raw_data)
                    tmp_path = tmp.name
                    
                try:
                    uploaded_file = client.files.upload(file=tmp_path)
                    contents.append(uploaded_file)
                    print(f"   Successfully uploaded to File API: {uploaded_file.name}")
                except Exception as e:
                    print(f"   Failed to upload {att.get('filename')} via File API: {e}")
                finally:
                    if os.path.exists(tmp_path):
                        try:
                            os.remove(tmp_path)
                        except Exception:
                            pass
            else:
                contents.append(
                    types.Part.from_bytes(
                        data=raw_data,
                        mime_type=ext_type,
                    )
                )
            
    print("Calling Gemini GenAI API...")
    response = client.models.generate_content(
        model=os.getenv("GEMINI_MODEL", "gemini-3.7-flash"),
        contents=contents,
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            response_mime_type="application/json",
            response_schema=ExtractedSOA,
            temperature=0.0
        )
    )
    
    usage = response.usage_metadata
    usage_info = {
        "prompt_tokens": usage.prompt_token_count if usage else 0,
        "output_tokens": usage.candidates_token_count if usage else 0,
        "total_tokens": usage.total_token_count if usage else 0
    }
    
    return ExtractedSOA.model_validate_json(response.text), usage_info

def export_to_excel(extracted_soa: ExtractedSOA, output_filename: str):
    records = []
    for item in extracted_soa.line_items:
        records.append({
            "Date": item.invoice_date,
            "Invoice Number": item.invoice_number,
            "Amount( Original )": item.invoice_value,
            "Open amount": item.open_amount,
            "Invoice Status": "",
            "Pending with": ""
        })
        
    df = pd.DataFrame(records, columns=[
        "Date", "Invoice Number", "Amount( Original )", "Open amount", "Invoice Status", "Pending with"
    ])
    
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='SOA_Status', index=False)
        
    print(f"Successfully exported {len(records)} line items to {output_filename} with sheet 'SOA_Status'.")

if __name__ == "__main__":
    msg_file = r"C:\Users\SohaibNM\OneDrive - Otis Global Tenant\Documents\Projects\VSR_UK\Data\FW A1 Stairlifts Limited - Request for statement of accounts - July 2026.msg"
    
    # 1. Parse and extract all contents natively
    data = extract_msg_contents(msg_file)
    print("\n--- Extraction complete! Output preview: ---")
    print(f"Total attachments extracted: {len(data['attachments'])}")
    for att in data['attachments']:
        print(f"  - Name: {att['filename']} (Type: {att['extension']})")
        if "excel_json" in att:
            print(f"    Excel Sheets Extracted: {list(att['excel_json'].keys())}")
            
    # 2. Call LLM for Structured Extraction
    print("\n--- Starting LLM Extraction ---")
    extracted_soa_data, usage = call_llm_for_extraction(data)
    
    # 3. Export exactly as required
    print("\n--- Exporting to Excel ---")
    output_excel_path = "Output_SOA_Status.xlsx"
    export_to_excel(extracted_soa_data, output_excel_path)
