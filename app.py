from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
import base64, logging, re
from datetime import datetime
from io import BytesIO
import PyPDF2, docx, openpyxl
import torch
from transformers import pipeline, AutoTokenizer, AutoModelForSeq2SeqLM
from sentence_transformers import SentenceTransformer

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize FastAPI app
app = FastAPI(
    title="Email Summarization API",
    description="API for summarizing emails and analyzing sentiment using Hugging Face models.",
    version="1.0.0"
)

# CORS Middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --------- Request & Response Models ----------
class Attachment(BaseModel):
    name: str
    content: str
    contentType: Optional[str] = None

class EmailRequest(BaseModel):
    subject: Optional[str] = ""
    body: Optional[str] = ""
    attachmentNames: Optional[List[str]] = []
    attachmentContents: Optional[List[Attachment]] = []

class SummaryResponse(BaseModel):
    summary: str
    sentiment: str
    has_attachments: bool
    attachment_count: int
    word_count: int
    urgency_score: int
    importance_score: int

class HealthResponse(BaseModel):
    status: str
    model_loaded: bool
    timestamp: str

# --------- Summarizer Class (Same as Flask Version) ----------
class EmailSummarizer:
    def __init__(self):
        try:
            model_name = "facebook/bart-large-cnn"
            self.tokenizer = AutoTokenizer.from_pretrained(model_name)
            self.model = AutoModelForSeq2SeqLM.from_pretrained(model_name)
            self.summarizer = pipeline("summarization",
                                       model=self.model,
                                       tokenizer=self.tokenizer,
                                       device=0 if torch.cuda.is_available() else -1)
            self.sentence_model = SentenceTransformer('all-MiniLM-L6-v2')
            logger.info("Models loaded successfully")
        except Exception as e:
            logger.error(f"Error loading models: {e}")
            self.summarizer = None
            self.sentence_model = None

    def clean_text(self, text):
        if not text:
            return ""
        text = re.sub(r'<[^>]+>', '', text)
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'(?i)sent from my \w+', '', text)
        text = re.sub(r'(?i)get outlook for \w+', '', text)
        text = re.sub(r'\n{3,}', '\n\n', text)
        return text.strip()

    def extract_key_information(self, email_data: EmailRequest):
        body = self.clean_text(email_data.body or '')
        subject = email_data.subject or ''
        attachments = email_data.attachmentNames or []
        attachment_contents = email_data.attachmentContents or []
        attachment_text = ""

        for att in attachment_contents:
            try:
                content = self.extract_attachment_text(att.dict())
                if content:
                    attachment_text += f"\n\nAttachment '{att.name}':\n{content}"
            except Exception as e:
                logger.warning(f"Failed to process attachment {att.name}: {e}")

        full_text = f"Subject: {subject}\n\n{body}"
        if attachment_text:
            full_text += f"\n\nAttachments:{attachment_text}"

        return {
            'full_text': full_text,
            'body': body,
            'subject': subject,
            'has_attachments': len(attachments) > 0,
            'attachment_names': attachments
        }

    def extract_attachment_text(self, attachment):
        try:
            content_b64 = attachment['content']
            content_bytes = base64.b64decode(content_b64)
            filename = attachment['name']
            content_type = attachment.get('contentType', '')

            if filename.lower().endswith('.pdf') or 'pdf' in content_type.lower():
                return self.extract_pdf_text(content_bytes)
            elif filename.lower().endswith(('.doc', '.docx')) or 'word' in content_type.lower():
                return self.extract_word_text(content_bytes)
            elif filename.lower().endswith(('.xls', '.xlsx')) or 'excel' in content_type.lower():
                return self.extract_excel_text(content_bytes)
            elif filename.lower().endswith('.txt') or 'text' in content_type.lower():
                return content_bytes.decode('utf-8', errors='ignore')
            else:
                return f"[{filename} - content type not supported for text extraction]"

        except Exception as e:
            logger.error(f"Error extracting text from attachment: {e}")
            return f"[Error reading {attachment['name']}]"

    def extract_pdf_text(self, pdf_bytes):
        try:
            pdf_file = BytesIO(pdf_bytes)
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
            return text.strip()
        except Exception as e:
            logger.error(f"Error extracting PDF text: {e}")
            return "[PDF text extraction failed]"

    def extract_word_text(self, doc_bytes):
        try:
            doc_file = BytesIO(doc_bytes)
            doc = docx.Document(doc_file)
            text = ""
            for paragraph in doc.paragraphs:
                text += paragraph.text + "\n"
            return text.strip()
        except Exception as e:
            logger.error(f"Error extracting Word text: {e}")
            return "[Word document text extraction failed]"

    def extract_excel_text(self, excel_bytes):
        try:
            excel_file = BytesIO(excel_bytes)
            workbook = openpyxl.load_workbook(excel_file)
            text = ""
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                text += f"Sheet: {sheet_name}\n"
                for row in sheet.iter_rows(values_only=True):
                    row_text = " | ".join([str(cell) if cell is not None else "" for cell in row])
                    if row_text.strip():
                        text += row_text + "\n"
                text += "\n"
            return text.strip()
        except Exception as e:
            logger.error(f"Error extracting Excel text: {e}")
            return "[Excel text extraction failed]"

    def generate_summary_local(self, text, max_length=150):
        if not self.summarizer:
            return "Summarization service not available"
        try:
            summary = self.summarizer(text[:1000], max_length=80, min_length=30, do_sample=False)[0]["summary_text"]
            return summary
        except Exception as e:
            logger.error(f"Error generating summary: {e}")
            return f"Summary generation failed: {str(e)}"

    def analyze_email_sentiment(self, text):
        try:
            urgent_keywords = ['urgent', 'asap', 'immediately', 'emergency', 'critical', 'deadline']
            importance_keywords = ['important', 'priority', 'required', 'mandatory', 'essential']

            text_lower = text.lower()
            urgency_score = sum(1 for keyword in urgent_keywords if keyword in text_lower)
            importance_score = sum(1 for keyword in importance_keywords if keyword in text_lower)

            sentiment = "neutral"
            if urgency_score > 0:
                sentiment = "urgent"
            elif importance_score > 0:
                sentiment = "important"

            return {
                'sentiment': sentiment,
                'urgency_score': urgency_score,
                'importance_score': importance_score
            }
        except Exception as e:
            logger.error(f"Error analyzing sentiment: {e}")
            return {'sentiment': 'neutral', 'urgency_score': 0, 'importance_score': 0}

# Initialize summarizer
email_summarizer = EmailSummarizer()

# --------- API Endpoints ----------
@app.post("/api/summarize", response_model=SummaryResponse)
def summarize_email(email_data: EmailRequest):
    try:
        extracted_info = email_summarizer.extract_key_information(email_data)
        summary = email_summarizer.generate_summary_local(extracted_info['full_text'])
        sentiment_analysis = email_summarizer.analyze_email_sentiment(extracted_info['full_text'])

        return SummaryResponse(
            summary=summary,
            sentiment=sentiment_analysis['sentiment'],
            has_attachments=extracted_info['has_attachments'],
            attachment_count=len(extracted_info['attachment_names']),
            word_count=len(extracted_info['body'].split()),
            urgency_score=sentiment_analysis['urgency_score'],
            importance_score=sentiment_analysis['importance_score']
        )
    except Exception as e:
        logger.error(f"Error processing email: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/health", response_model=HealthResponse)
def health_check():
    return HealthResponse(
        status="healthy",
        model_loaded=email_summarizer.summarizer is not None,
        timestamp=datetime.now().isoformat()
    )

# Run using: uvicorn app:app --reload --host 0.0.0.0 --port 5000
