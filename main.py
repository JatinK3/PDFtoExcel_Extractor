import os, sys, json, time
from pathlib import Path
import pdfplumber
import pandas as pd
from tqdm import tqdm
from dotenv import load_dotenv
from google import genai

load_dotenv()
dotenv_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=dotenv_path, override=False)

MAX_CHUNK_CHARS = 1800
DEFAULT_MODEL = "gemini-2.5-flash"
CHUNK_SLEEP = 0.2

def extract_pdf_text(path: Path):
    pages = []
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            txt = page.extract_text() or ""
            pages.append({"page": i, "text": txt})
    return pages

def chunk_text_block(text: str, max_chars: int = MAX_CHUNK_CHARS):
    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
    chunks, cur = [], ""
    for p in paragraphs:
        if len(cur) + len(p) + 2 <= max_chars:
            cur = (cur + "\n\n" + p).strip()
        else:
            if cur: chunks.append(cur)
            if len(p) <= max_chars:
                cur = p
            else:
                for i in range(0, len(p), max_chars):
                    chunks.append(p[i:i+max_chars])
                cur = ""
    if cur: chunks.append(cur)
    return chunks

def safe_parse_json(maybe_json_text: str):
    try:
        parsed = json.loads(maybe_json_text)
        if isinstance(parsed, list):
            return parsed
        if isinstance(parsed, dict) and "items" in parsed:
            return parsed["items"]
    except Exception:
        pass

    start = maybe_json_text.find('[')
    end = maybe_json_text.rfind(']')
    if start != -1 and end != -1 and end > start:
        return json.loads(maybe_json_text[start:end+1])

    raise ValueError(f"Invalid JSON output:\n{maybe_json_text[:500]}")

PROMPT_TEMPLATE = """
You are a reliable extractor. Input is an English text chunk from a PDF. Your task:
1) Identify all implicit or explicit key:value pairs present in the text.
2) Preserve original wording.
3) Return each as: "key", "value", "comments".
4) If content cannot form key:value pairs, create key="__UNSTRUCTURED__" and value=full chunk.
Output JSON only.
Chunk:{chunk}
"""

genai_client = genai.Client()

def extract_text_from_genai_response(resp):
    if hasattr(resp, "text") and resp.text:
        return resp.text
    if hasattr(resp, "candidates") and resp.candidates:
        parts = []
        for c in resp.candidates:
            content = getattr(c, "content", None) or getattr(c, "text", None)
            if content: parts.append(content)
        return "\n".join(parts)
    return str(resp)

def call_llm_model_with_genai(prompt: str, model: str = DEFAULT_MODEL):
    try:
        resp = genai_client.models.generate_content(model=model, contents=prompt)
        return extract_text_from_genai_response(resp)
    except Exception as e:
        raise RuntimeError(f"GenAI call failed: {e}")

def process_pdf_to_rows(pdf_path: Path, model: str = DEFAULT_MODEL):
    pages = extract_pdf_text(pdf_path)
    rows, full_text_backup = [], []

    for page in tqdm(pages, desc="Processing pages"):
        page_num = page["page"]
        txt = page["text"] or ""
        full_text_backup.append({"page": page_num, "text": txt})
        chunks = chunk_text_block(txt)

        if not chunks:
            rows.append({"Key": "__UNSTRUCTURED__", "Value": "", "Comments": "Empty page", "Source_Page": page_num})
            continue

        for chunk in chunks:
            prompt = PROMPT_TEMPLATE.replace("{chunk}", chunk)

            try:
                raw_out = call_llm_model_with_genai(prompt, model=model)
            except Exception as e:
                rows.append({
                    "Key": "__UNSTRUCTURED__",
                    "Value": chunk,
                    "Comments": f"Model error: {str(e)}",
                    "Source_Page": page_num
                })
                time.sleep(CHUNK_SLEEP)
                continue

            try:
                items = safe_parse_json(raw_out)
            except Exception:
                rows.append({
                    "Key": "__UNSTRUCTURED__",
                    "Value": chunk,
                    "Comments": f"Bad JSON output.\n{raw_out[:800]}",
                    "Source_Page": page_num
                })
                time.sleep(CHUNK_SLEEP)
                continue

            for it in items:
                if not isinstance(it, dict):
                    rows.append({"Key": "__UNSTRUCTURED__", "Value": str(it), "Comments": "Invalid item", "Source_Page": page_num})
                    continue

                rows.append({
                    "Key": it.get("key", ""),
                    "Value": it.get("value", ""),
                    "Comments": it.get("comments", ""),
                    "Source_Page": page_num
                })

            time.sleep(CHUNK_SLEEP)

    return rows, full_text_backup

def write_to_excel(rows, full_text_backup, out_path: Path):
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name="Structured", index=False)
        pd.DataFrame(full_text_backup).to_excel(writer, sheet_name="Raw_Pages", index=False)
        pd.DataFrame({
            "num_structured_rows": [len(rows)],
            "num_pages": [len(full_text_backup)]
        }).to_excel(writer, sheet_name="Metrics", index=False)

def main():
    inp = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("Data Input.pdf")
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("Structured Output (GENERATED).xlsx")

    if not inp.exists():
        print(f"Input file not found: {inp}")
        sys.exit(1)

    print(f"Running extraction...")
    rows, backup = process_pdf_to_rows(inp, model=DEFAULT_MODEL)
    write_to_excel(rows, backup, out)
    print(f"Done. Output written to {out}")

if __name__ == "__main__":
    main()

