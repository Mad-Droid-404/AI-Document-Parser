import os
from flask import Flask, request, jsonify
from flask_cors import CORS
from transformers import pipeline, AutoTokenizer
import torch
import base64
import PyPDF2
import docx
import io
import uuid
import re
import random
from datetime import datetime

app = Flask(__name__)
CORS(app)

print("Loading BART model...")
try:
    model_name = "facebook/bart-large-cnn"
    device = 0 if torch.cuda.is_available() else -1

    summarizer = pipeline("summarization", model=model_name, device=device, framework="pt")
    tokenizer = AutoTokenizer.from_pretrained(model_name)
    print(f"✅ BART model loaded on {'GPU' if device == 0 else 'CPU'}")

except Exception as e:
    try:
        model_name = "facebook/bart-base"
        summarizer = pipeline("summarization", model=model_name, device=-1)
        tokenizer = AutoTokenizer.from_pretrained(model_name)
        print("✅ BART base model loaded")
    except Exception as e2:
        print(f"❌ Model loading failed: {e2}")
        summarizer = None
        tokenizer = None


def extract_text_from_attachment(file_data, filename):
    try:
        file_bytes = base64.b64decode(file_data)

        if filename.lower().endswith('.pdf'):
            pdf_reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
            return "".join(page.extract_text() for page in pdf_reader.pages)

        elif filename.lower().endswith('.docx'):
            doc = docx.Document(io.BytesIO(file_bytes))
            return "\n".join(paragraph.text for paragraph in doc.paragraphs)

        elif filename.lower().endswith('.txt'):
            return file_bytes.decode('utf-8')

        return f"[{filename} - Unsupported format]"

    except Exception as e:
        return f"[Error: {filename} - {str(e)}]"


def preprocess_text(text):
    if not text:
        return ""
    text = re.sub(r'\s+', ' ', text.strip())
    lines = [line.strip() for line in text.split('\n') if len(line.strip()) > 10]
    return ' '.join(lines)


def chunk_text(text, max_tokens=900):
    if not tokenizer:
        max_chars = max_tokens * 4
        return [text[i:i + max_chars] for i in range(0, len(text), max_chars)]

    tokens = tokenizer.encode(text, add_special_tokens=False)
    if len(tokens) <= max_tokens:
        return [text]

    chunks = []
    for i in range(0, len(tokens), max_tokens):
        chunk_tokens = tokens[i:i + max_tokens]
        chunks.append(tokenizer.decode(chunk_tokens, skip_special_tokens=True))

    return chunks


def get_style_config(style):
    configs = {
        "short": {"temp": 0.7, "top_p": 0.85, "top_k": 40, "beams": 2, "rep_penalty": 1.3, "len_penalty": 1.2},
        "long": {"temp": 0.8, "top_p": 0.92, "top_k": 50, "beams": 1, "rep_penalty": 1.1, "len_penalty": 0.8},
        "bullets": {"temp": 0.6, "top_p": 0.85, "top_k": 35, "beams": 2, "rep_penalty": 1.4, "len_penalty": 1.0},
        "action": {"temp": 0.7, "top_p": 0.88, "top_k": 45, "beams": 2, "rep_penalty": 1.2, "len_penalty": 1.0},
        "highlights": {"temp": 0.75, "top_p": 0.9, "top_k": 45, "beams": 1, "rep_penalty": 1.25, "len_penalty": 0.9},
        "executive": {"temp": 0.65, "top_p": 0.85, "top_k": 40, "beams": 2, "rep_penalty": 1.15, "len_penalty": 0.95}
    }
    return configs.get(style, configs["short"])


def format_output(text, style):
    if style == "bullets":
        sentences = [s.strip().rstrip('.') for s in text.replace('.', '.\n').split('\n') if s.strip()]
        return '\n'.join([f"• {s}" for s in sentences]) if len(sentences) > 1 else f"• {text}"

    elif style == "action":
        sentences = [s.strip().rstrip('.') for s in text.replace('.', '.\n').split('\n') if s.strip()]
        action_words = ['should', 'must', 'need', 'require', 'action', 'task', 'follow up', 'schedule', 'contact',
                        'prepare', 'review', 'complete', 'submit']
        formatted = []
        for sentence in sentences:
            if sentence:
                has_action = any(word in sentence.lower() for word in action_words)
                icon = "🔹" if has_action else "•"
                formatted.append(f"{icon} {sentence}")
        return '\n'.join(formatted) if formatted else f"🔹 {text}"

    elif style == "highlights":
        sentences = [s.strip().rstrip('.') for s in text.replace('.', '.\n').split('\n') if
                     s.strip() and len(s.strip()) > 10]
        return '\n'.join([f"⭐ {s}" for s in sentences]) if len(sentences) > 1 else f"⭐ {text}"

    return text.strip()


def generate_summary(content, style, max_length, min_length):
    if not summarizer:
        return "Error: AI model not loaded"

    if len(content.strip()) < 50:
        return "Text too short to summarize"

    cleaned_text = preprocess_text(content)
    chunks = chunk_text(cleaned_text)
    config = get_style_config(style)

    if len(chunks) == 1:
        result = summarizer(
            cleaned_text,
            max_length=max_length,
            min_length=min_length,
            do_sample=True,
            temperature=config['temp'],
            top_p=config['top_p'],
            top_k=config['top_k'],
            num_beams=config['beams'],
            repetition_penalty=config['rep_penalty'],
            length_penalty=config['len_penalty'],
            no_repeat_ngram_size=3
        )
        return format_output(result[0]['summary_text'], style)

    else:
        chunk_summaries = []
        for chunk in chunks:
            try:
                chunk_temp = config['temp'] + (random.random() * 0.2 - 0.1)
                chunk_temp = max(0.1, min(1.0, chunk_temp))

                result = summarizer(
                    chunk,
                    max_length=min(120, max_length // len(chunks) + 50),
                    min_length=max(20, min_length // len(chunks)),
                    do_sample=True,
                    temperature=chunk_temp,
                    top_p=config['top_p'],
                    top_k=config['top_k'],
                    num_beams=config['beams'],
                    repetition_penalty=config['rep_penalty']
                )
                chunk_summaries.append(result[0]['summary_text'])
            except Exception as e:
                chunk_summaries.append(f"[Processing error: {str(e)}]")

        combined = " ".join(chunk_summaries)

        if len(combined.split()) > max_length:
            try:
                final_result = summarizer(
                    combined,
                    max_length=max_length,
                    min_length=min_length,
                    do_sample=True,
                    temperature=config['temp'],
                    top_p=config['top_p'],
                    repetition_penalty=config['rep_penalty']
                )
                return format_output(final_result[0]['summary_text'], style)
            except:
                return format_output(combined, style)

        return format_output(combined, style)


def process_attachments(attachments, mode):
    processed = []
    for att in attachments:
        filename = att.get('name', 'unknown')
        file_data = att.get('content', '')
        if file_data:
            content = extract_text_from_attachment(file_data, filename)
            processed.append({'filename': filename, 'content': content})
    return processed


def create_sections(email_body, attachments, summary_type, attachment_mode):
    sections = []

    if email_body and email_body.strip():
        email_summary = generate_summary(email_body, summary_type, 200, 50)
        sections.append({'title': '📧 Email Summary', 'content': email_summary})

    if attachments:
        if attachment_mode == 'separate':
            for att in attachments:
                if att['content'] and att['content'].strip():
                    att_summary = generate_summary(att['content'], summary_type, 150, 30)
                    sections.append({'title': f"📎 {att['filename']}", 'content': att_summary})
        else:
            combined_text = "\n\n".join([f"--- {att['filename']} ---\n{att['content']}"
                                         for att in attachments if att['content'] and att['content'].strip()])
            if combined_text.strip():
                att_summary = generate_summary(combined_text, summary_type, 200, 50)
                sections.append({'title': f"📎 Attachments ({len(attachments)} files)", 'content': att_summary})

    return sections


@app.route('/', methods=['GET'])
def home():
    """Root endpoint - this was missing!"""
    return jsonify({
        "message": "✅ Email Summarization API is running!",
        "status": "active",
        "endpoints": {
            "/health": "Health check",
            "/summarize": "POST - Summarize emails and attachments",
            "/test": "GET - Simple test endpoint"
        },
        "timestamp": datetime.now().isoformat()
    })


@app.route('/test', methods=['GET'])
def test():
    """Test endpoint"""
    return jsonify({
        "message": "🎉 Test endpoint working!",
        "model_status": "loaded" if summarizer else "not loaded",
        "gpu_available": torch.cuda.is_available()
    })


@app.route('/summarize', methods=['POST'])
def summarize_email():
    try:
        if not summarizer:
            return jsonify({'success': False, 'error': 'AI model not available'}), 503

        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'error': 'No data provided'}), 400

        email_body = data.get('emailBody', '')
        attachments = data.get('attachments', [])
        summary_type = data.get('summaryType', 'short')
        parsing_options = data.get('parsingOptions', {})

        include_email = parsing_options.get('includeEmail', True)
        include_attachments = parsing_options.get('includeAttachments', True)
        attachment_mode = parsing_options.get('attachmentMode', 'combined')
        output_mode = parsing_options.get('outputMode', 'combined')

        if not include_email and not include_attachments:
            return jsonify({'success': False, 'error': 'No content selected'}), 400

        request_id = str(uuid.uuid4())
        processed_attachments = process_attachments(attachments, attachment_mode) if include_attachments else []

        length_map = {'short': (100, 30), 'long': (300, 100), 'bullets': (200, 80), 'action': (180, 60),
                      'highlights': (150, 50), 'executive': (250, 80)}
        max_length, min_length = length_map.get(summary_type, (150, 50))

        if output_mode == 'sections':
            sections = create_sections(
                email_body if include_email else '',
                processed_attachments,
                summary_type,
                attachment_mode
            )
            result = {'summaryData': {'sections': sections}}
        else:
            content_parts = []
            if include_email and email_body.strip():
                content_parts.append(email_body.strip())

            if include_attachments and processed_attachments:
                if attachment_mode == 'separate':
                    att_summaries = []
                    for att in processed_attachments:
                        if att['content'].strip():
                            att_sum = generate_summary(att['content'], summary_type, 100, 20)
                            att_summaries.append(f"{att['filename']}: {att_sum}")
                    if att_summaries:
                        content_parts.append(" | ".join(att_summaries))
                else:
                    combined_att = "\n\n".join([f"--- {att['filename']} ---\n{att['content']}"
                                                for att in processed_attachments if att['content'].strip()])
                    if combined_att.strip():
                        content_parts.append(combined_att)

            if not content_parts:
                return jsonify({'success': False, 'error': 'No content to summarize'}), 400

            full_content = "\n\n".join(content_parts)
            summary = generate_summary(full_content, summary_type, max_length, min_length)
            result = {'summary': summary}

        return jsonify({
            'success': True,
            'summary': result.get('summary', ''),
            'summaryData': result.get('summaryData'),
            'summaryType': summary_type,
            'parsingOptions': parsing_options,
            'requestId': request_id,
            'timestamp': datetime.now().isoformat(),
            'attachmentsProcessed': len(processed_attachments)
        })

    except Exception as e:
        return jsonify({
            'success': False,
            'error': 'Processing failed',
            'timestamp': datetime.now().isoformat()
        }), 500


@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({
        'status': 'healthy' if summarizer else 'degraded',
        'model_loaded': summarizer is not None,
        'timestamp': datetime.now().isoformat()
    })


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug_mode = os.environ.get('DEBUG', 'False').lower() == 'true'

    print(f"🌐 Starting Flask server on port {port}")
    print(f"🔧 Debug mode: {debug_mode}")

    app.run(
        host='0.0.0.0',
        port=port,
        debug=debug_mode,
        threaded=True
    )