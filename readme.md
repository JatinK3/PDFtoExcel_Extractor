# AI-Powered PDF → Structured Excel Extractor  
Uses **Google Gemini 2.5 Flash** to convert unstructured PDF documents into a structured Excel format with zero predefined schema.

## Features
- Extracts **all text** from any readable PDF (no OCR needed)
- Automatically identifies **key:value pairs** using an LLM
- Preserves **original text** without paraphrasing
- Captures additional context into a **Comments** column
- Maintains complete data coverage with **Raw_Pages** backup sheet
- Outputs a clean, structured **Excel (.xlsx)** file

---

## 📦 Requirements

### **1. Python 3.10+**
Ensure Python is installed.

### **2. Install required Python packages**
Run:

```bash
pip install pdfplumber pandas openpyxl tqdm python-dotenv google-genai
Or use the included requirements file (if provided):
pip install -r requirements.txt


🔑 Environment Setup
Create a .env file in the project folder:
GEMINI_API_KEY=YOUR_API_KEY_HERE
You can get a free API key from:
🔗 https://aistudio.google.com/

📁 Project Structure 

📂 YourProject/
 ├── main.py
 ├── Data Input.pdf
 ├── .env
 ├── README.md
 └── requirements.txt

▶️ How to Run
1.
Default (uses Data Input.pdf and outputs Structured Output (Generated).xlsx)
python main.py

2.
Custom file input/output
python main.py "path/to/input.pdf" "path/to/output.xlsx"

Example:
python main.py "./samples/resume.pdf" "./out/resume_extracted.xlsx"

📤 Output Files
Structured Output (Generated).xlsx

🧠 How It Works (Summary)
PDF text is extracted with pdfplumber
Text is split into safe LLM chunks
Each chunk is sent to Gemini 2.5 Flash with a structured-extraction prompt
The model returns JSON containing keys/values/comments
JSON is parsed and written to Excel
Any errors or malformed model output are stored as __UNSTRUCTURED__ rows (no data loss)


❗Important Notes
The script does not predefine keys — key inference is fully LLM-driven
It does not summarize or remove text
It works on text-layer PDFs (no OCR)
Only the Google google-genai client is used (not REST API)


🏆 Ideal For
AI/Data engineering tasks
Document automation
Resume/document parsing
Enterprise data extraction workflows