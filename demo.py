#!/usr/bin/env python3
"""
Document to Excel extractor for distribution companies.

Features:
- Supports Word (.docx), PDF (.pdf), and images (.jpg, .jpeg, .png)
- Extracts text (with OCR for images and scanned PDFs)
- Uses OpenAI to parse building materials documents into structured JSON
- Appends rows to a master Excel file

Usage:
  python doc_to_excel.py --input /path/to/file_or_folder --output data.xlsx
"""

import os
import sys
import json
import logging
import argparse
from pathlib import Path
from typing import List, Dict, Optional

import pandas as pd
from dotenv import load_dotenv

import pdfplumber
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
from docx import Document

from openai import OpenAI

LLAVE_DE_MI_VIDA = "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"

ALLOWED_COLUMNS = [
    "document_type",
    "document_number",
    "document_date",
    "supplier_name",
    "customer_name",
    "shipment_location",
    "payment_terms",
    "line_number",
    "item_code",
    "item_description",
    "category",
    "quantity",
    "unit_of_measure",
    "unit_price",
    "extended_price"
]

# ------------- Config and logging -------------

SUPPORTED_SUFFIXES = {".docx", ".pdf", ".jpg", ".jpeg", ".png", ".csv", ".xlsx"}
MAX_CHARS_FOR_LLM = 15000  # simple safeguard


def setup_logging(level: str = "INFO") -> None:
    numeric = getattr(logging, level.upper(), logging.INFO)
    logging.basicConfig(
        level=numeric,
        format="[%(asctime)s] %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def load_config() -> Dict[str, str]:
    load_dotenv()
    api_key = os.getenv("OPENAI_API_KEY") #LLAVE_DE_MI_VIDA
    if not api_key:
        logging.error("OPENAI_API_KEY is not set in environment or .env")
        sys.exit(1)
    model = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
    return {"api_key": api_key, "model": model}


def get_openai_client(api_key: str) -> OpenAI:
    return OpenAI(api_key=api_key)


# ------------- Text extraction -------------

def extract_text_from_docx(path: str) -> str:
    try:
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except Exception as e:
        logging.exception("Failed to extract text from DOCX %s: %s", path, e)
        raise


def extract_text_from_pdf_text_layer(path: str) -> str:
    text_chunks: List[str] = []
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                txt = page.extract_text()
                if txt:
                    text_chunks.append(txt)
    except Exception as e:
        logging.exception("Failed to read PDF %s with pdfplumber: %s", path, e)
        raise
    return "\n".join(text_chunks)


def extract_text_from_pdf_ocr(path: str) -> str:
    """
    Fallback OCR for scanned PDFs using pdf2image + pytesseract.
    """
    try:
        images = convert_from_path(path)
    except Exception as e:
        logging.exception("Failed to convert PDF %s to images for OCR: %s", path, e)
        raise

    all_text: List[str] = []
    for idx, img in enumerate(images, start=1):
        try:
            txt = pytesseract.image_to_string(img)
            if txt.strip():
                all_text.append(txt)
        except Exception as e:
            logging.warning("OCR failed on page %d of %s: %s", idx, path, e)
    return "\n".join(all_text)


def extract_text_from_pdf(path: str) -> str:
    # Try text layer first
    text = extract_text_from_pdf_text_layer(path)
    if text.strip():
        logging.info("Extracted text from PDF text layer for %s", path)
        return text

    # Fallback to OCR
    logging.info("No text layer detected in %s, running OCR on PDF pages", path)
    return extract_text_from_pdf_ocr(path)


def extract_text_from_image(path: str) -> str:
    try:
        img = Image.open(path)
        txt = pytesseract.image_to_string(img)
        return txt
    except Exception as e:
        logging.exception("Failed to extract text from image %s: %s", path, e)
        raise


def extract_text(path: str) -> str:
    suffix = Path(path).suffix.lower()
    logging.info("Extracting text from %s (type %s)", path, suffix)

    if suffix == ".docx":
        return extract_text_from_docx(path)
    if suffix == ".pdf":
        return extract_text_from_pdf(path)
    if suffix in {".jpg", ".jpeg", ".png"}:
        return extract_text_from_image(path)

    raise ValueError(f"Unsupported file type: {suffix}")

def read_tabular_file_to_rows(path: str) -> list[dict]:
    suffix = Path(path).suffix.lower()

    if suffix == ".csv":
        df = pd.read_csv(path)
    elif suffix == ".xlsx":
        df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    else:
        raise ValueError(f"Not a tabular file: {suffix}")

    # Normalize to your allowed schema
    for col in ALLOWED_COLUMNS:
        if col not in df.columns:
            df[col] = "Not Defined"

    df = df[ALLOWED_COLUMNS].fillna("Not Defined")

    return df.to_dict(orient="records")

 

def read_tabular(path: str) -> pd.DataFrame:
    suf = Path(path).suffix.lower()
    if suf == ".csv":
        return pd.read_csv(path)
    if suf == ".xlsx":
        return pd.read_excel(path, sheet_name=0, engine="openpyxl")
    raise ValueError(f"Unsupported tabular file: {suf}")

def df_to_compact_table_text(df: pd.DataFrame, max_rows: int = 60, max_cols: int = 20) -> str:
    # Keep it bounded
    df2 = df.copy()
    if df2.shape[1] > max_cols:
        df2 = df2.iloc[:, :max_cols]
    if len(df2) > max_rows:
        df2 = df2.head(max_rows)

    # Make sure column names are strings
    df2.columns = [str(c) for c in df2.columns]

    # Convert to a plain-text table the LLM can read
    # (markdown table style)
    return df2.to_markdown(index=False)

# ------------- LLM call -------------

SYSTEM_PROMPT = """
You are an assistant that extracts structured data from invoices, quotes, and packing slips
for building materials distributors.

Given the text of a document, return a single JSON object with this structure:

{
  "document_type": "invoice | quote | packing_slip | purchase_order | other",
  "document_number": "string or null",
  "document_date": "YYYY-MM-DD or null",
  "supplier_name": "string or null",
  "customer_name": "string or null",
  "shipment_location": "string or null",
  "payment_terms": "string or null",
  "lines": [
    {
      "line_number": int or null,
      "item_code": "string or null",
      "item_description": "string or null",
      "category": "lumber | drywall | roofing | fasteners | insulation | other",
      "quantity": float or null,
      "unit_of_measure": "string or null",
      "unit_price": float or null,
      "extended_price": float or null
    }
  ]
}

Rules:
- If a field is missing in the document, set it to null.
- Use ISO date format for document_date.
- If no line items exist, return an empty list for "lines".
- Only output valid JSON. No additional text.
""".strip()


def truncate_text(raw_text: str, max_chars: int = MAX_CHARS_FOR_LLM) -> str:
    if len(raw_text) <= max_chars:
        return raw_text
    logging.warning("Document text is long (%d chars). Truncating to %d chars.", len(raw_text), max_chars)
    return raw_text[:max_chars]

def fill_empty(obj, replacement="N/A"):
    if isinstance(obj, dict):
        return {k: fill_empty(v, replacement) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [fill_empty(v, replacement) for v in obj]
    elif obj is None or obj == "":
        return replacement
    else:
        return obj


def extract_structured_data_from_tabular_with_llm(client, model: str, df: pd.DataFrame) -> dict:
    table_text = df_to_compact_table_text(df)

    prompt = f"""
    You are given a spreadsheet excerpt. It may be an invoice/quote/packing slip or a tabular document.

    Spreadsheet excerpt (markdown table):
    {table_text}

    Return ONLY valid JSON in this exact schema:
    {SYSTEM_PROMPT.split("Given the text of a document, return a single JSON object with this structure:")[1].strip()}
    """.strip()

    # Reuse whatever OpenAI call you already do inside extract_structured_data_with_llm,
    # but pass `prompt` as the user content.
    return extract_structured_data_with_llm(client, model, prompt)

def extract_structured_data_with_llm(
    client: OpenAI,
    model: str,
    raw_text: str,
) -> Dict:
    safe_text = truncate_text(raw_text)

    user_prompt = f"Document text:\n```text\n{safe_text}\n```"

    try:
        # Use chat completions with JSON mode
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            response_format={"type": "json_object"},
        )
    except Exception as e:
        logging.exception("OpenAI API call failed: %s", e)
        raise

    try:
        content = response.choices[0].message.content.strip()
    except Exception as e:
        logging.exception("Unexpected response format from OpenAI: %s", e)
        raise

    # At this point content should be pure JSON
    try:
        jsonThing = json.loads(content)
        filled_json = fill_empty(jsonThing, replacement="Unknown")
        return filled_json
       
    except json.JSONDecodeError as e:
        logging.error("Failed to parse JSON from OpenAI: %s", e)
        logging.debug("Raw response content: %s", content)
        raise


# ------------- JSON to Excel rows -------------

def json_to_rows(data: Dict) -> List[Dict]:
    header = {
        "document_type": data.get("document_type"),
        "document_number": data.get("document_number"),
        "document_date": data.get("document_date"),
        "supplier_name": data.get("supplier_name"),
        "customer_name": data.get("customer_name"),
        "shipment_location": data.get("shipment_location"),
        "payment_terms": data.get("payment_terms"),
    }

    lines = data.get("lines") or []
    rows: List[Dict] = []

    if not isinstance(lines, list):
        logging.warning("Expected 'lines' to be a list. Got %s", type(lines))
        lines = []

    if not lines:
        # Still create one row with header only
        rows.append(header)
        return rows

    for line in lines:
        if not isinstance(line, dict):
            logging.warning("Skipping non dict line item: %s", line)
            continue
        row = {**header, **line}
        rows.append(row)

    return rows

def append_rows_to_excel(rows: List[Dict], excel_path: str) -> None:
    if not rows:
        logging.info("No rows to append to Excel.")
        return

    # New rows as DataFrame, replacing missing values with "Not Defined"
    # Convert rows to DataFrame and keep only allowed columns
# Convert rows to DataFrame and keep only allowed columns
    new_df = pd.DataFrame(rows)

    # Insert missing columns with "Not Defined"
    for col in ALLOWED_COLUMNS:
        if col not in new_df.columns:
            new_df[col] = "Not Defined"

    # Drop any unexpected columns that should not be in the CSV
    new_df = new_df[ALLOWED_COLUMNS]

    # Replace blanks/NaN with "Not Defined"
    new_df = new_df.fillna("Not Defined")


    new_df = new_df.fillna("Not Defined")

    path = Path(excel_path)

    if path.exists():
        try:
            existing = pd.read_excel(path)
            # Also clean existing data to keep everything consistent
            existing = existing.fillna("Not Defined")
            combined = pd.concat([existing, new_df], ignore_index=True)
        except Exception as e:
            logging.exception("Failed to read existing Excel file %s: %s", excel_path, e)
            raise
    else:
        combined = new_df

    # Final safety: ensure there are no NaNs before writing
    combined = combined.fillna("Not Defined")

    try:
        combined.to_excel(path, index=False)
        logging.info("Wrote %d rows to %s", len(new_df), excel_path)
    except Exception as e:
        logging.exception("Failed to write Excel file %s: %s", excel_path, e)
        raise

# ------------- Processing logic -------------

def process_file(
    path: str,
    excel_path: str,
    client: OpenAI,
    model: str,
) -> Optional[int]:
    logging.info("Processing file: %s", path)

    try:
        raw_text = extract_text(path)
        if not raw_text.strip():
            logging.warning("No text extracted from %s. Skipping.", path)
            return None

        structured = extract_structured_data_with_llm(client, model, raw_text)
        rows = json_to_rows(structured)
        append_rows_to_excel(rows, excel_path)
        return len(rows)

    except Exception as e:
        logging.error("Failed to process %s: %s", path, e)
        return None


def process_path(
    input_path: str,
    excel_path: str,
    client: OpenAI,
    model: str,
) -> None:
    p = Path(input_path)

    if p.is_file():
        if p.suffix.lower() not in SUPPORTED_SUFFIXES:
            logging.error("File %s has unsupported extension.", p)
            return
        count = process_file(str(p), excel_path, client, model)
        if count:
            logging.info("Finished %s with %d rows.", p.name, count)
        return

    if p.is_dir():
        total_files = 0
        total_rows = 0
        for file in sorted(p.iterdir()):
            if not file.is_file():
                continue
            if file.suffix.lower() not in SUPPORTED_SUFFIXES:
                continue
            total_files += 1
            count = process_file(str(file), excel_path, client, model)
            if count:
                total_rows += count
        logging.info("Processed %d files and wrote %d rows in total.", total_files, total_rows)
        return

    logging.error("Input path %s is neither a file nor a directory.", input_path)


# ------------- CLI entry point -------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract building materials document data into Excel using OpenAI."
    )
    parser.add_argument(
        "--input",
        "-i",
        required=True,
        help="Path to input file or folder (Word, PDF, images).",
    )
    parser.add_argument(
        "--output",
        "-o",
        required=True,
        help="Path to output Excel file (will be created or updated).",
    )
    parser.add_argument(
        "--log-level",
        "-l",
        default="INFO",
        help="Logging level: DEBUG, INFO, WARNING, ERROR.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    setup_logging(args.log_level)
    config = load_config()
    client = get_openai_client(config["api_key"])

    logging.info("Using OpenAI model: %s", config["model"])
    logging.info("Input: %s", args.input)
    logging.info("Output Excel: %s", args.output)

    process_path(args.input, args.output, client, config["model"])


##if __name__ == "__main__":
  #  main()
