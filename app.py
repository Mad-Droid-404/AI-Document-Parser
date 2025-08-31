import os
import re
import io
import uuid
import base64
import random
from datetime import datetime

import torch
import PyPDF2
import docx
from flask import Flask, request, jsonify
from flask_cors import CORS
from transformers import pipeline, AutoTokenizer

MAX_TEXT_LENGTH = 50000
MAX_TOKENS_PER_CHUNK = 800
MAX_TOKEN_LIMIT = 1000
CHARS_PER_TOKEN_ESTIMATE = 3

SUMMARY_CONFIGS = {
    "short": {"temp": 0.7, "top_p": 0.85, "top_k": 40, "beams": 2, "rep_penalty": 1.3, "len_penalty": 1.2},
    "long": {"temp": 0.8, "top_p": 0.92, "top_k": 50, "beams": 1, "rep_penalty": 1.1, "len_penalty": 0.8},
    "bullets": {"temp": 0.6, "top_p": 0.85, "top_k": 35, "beams": 2, "rep_penalty": 1.4, "len_penalty": 1.0},
    "action": {"temp": 0.7, "top_p": 0.88, "top_k": 45, "beams": 2, "rep_penalty": 1.2, "len_penalty": 1.0},
    "highlights": {"temp": 0.75, "top_p": 0.9, "top_k": 45, "beams": 1, "rep_penalty": 1.25, "len_penalty": 0.9},
    "executive": {"temp": 0.65, "top_p": 0.85, "top_k": 40, "beams": 2, "rep_penalty": 1.15, "len_penalty": 0.95}
}

LENGTH_MAP = {
    'short': (100, 30), 'long': (300, 100), 'bullets': (200, 80),
    'action': (180, 60), 'highlights': (150, 50), 'executive': (250, 80)
}

ACTION_KEYWORDS = [
    'should', 'must', 'need', 'require', 'action', 'task', 'follow up',
    'schedule', 'contact', 'prepare', 'review', 'complete', 'submit'
]

app = Flask(__name__)
CORS(app)

summarizer = None
tokenizer = None
model_loaded = False


def initialize_model():
    """Initialize BART model with fallback options."""
    global summarizer, tokenizer, model_loaded

    model_configs = [
        ("facebook/bart-large-cnn", "BART large model"),
        ("facebook/bart-base", "BART base model")
    ]

    for model_name, description in model_configs:
        try:
            device = 0 if torch.cuda.is_available() else -1
            summarizer = pipeline("summarization", model=model_name, device=device, framework="pt")
            tokenizer = AutoTokenizer.from_pretrained(model_name)
            model_loaded = True
            print(f"Model loaded: {description} on {'GPU' if device == 0 else 'CPU'}")
            return
        except Exception as e:
            print(f"Failed to load {description}: {e}")

    print("Model loading failed completely")
    model_loaded = False


def extract_text_from_file(file_data, filename):
    """Extract text from various file formats."""
    try:
        file_bytes = base64.b64decode(file_data)
        extension = filename.lower().split('.')[-1]

        extractors = {
            'pdf': lambda: "".join(page.extract_text() for page in PyPDF2.PdfReader(io.BytesIO(file_bytes)).pages),
            'docx': lambda: "\n".join(p.text for p in docx.Document(io.BytesIO(file_bytes)).paragraphs),
            'txt': lambda: file_bytes.decode('utf-8')
        }

        if extension in extractors:
            return extractors[extension]()

        return f"[{filename} - Unsupported format]"

    except Exception as e:
        return f"[Error: {filename} - {str(e)}]"


def preprocess_text(text):
    """Clean and preprocess text for summarization."""
    if not text:
        return ""

    text = re.sub(r'\s+', ' ', text.strip())
    meaningful_lines = [line.strip() for line in text.split('\n') if len(line.strip()) > 10]
    processed = ' '.join(meaningful_lines)

    if len(processed) > MAX_TEXT_LENGTH:
        processed = processed[:MAX_TEXT_LENGTH] + "... [truncated]"

    return processed


def create_text_chunks(text, max_tokens=MAX_TOKENS_PER_CHUNK):
    """Safely chunk text into manageable pieces."""
    if not text:
        return []

    if not tokenizer:
        return _character_based_chunking(text, max_tokens)

    try:
        tokens = tokenizer.encode(text, add_special_tokens=False, truncation=False)

        if len(tokens) <= max_tokens:
            return [text]

        chunks = []
        for i in range(0, len(tokens), max_tokens):
            chunk_tokens = tokens[i:i + max_tokens]
            try:
                chunk_text = tokenizer.decode(chunk_tokens, skip_special_tokens=True)
                chunks.append(chunk_text)
            except Exception:
                # Fallback to character estimation
                start_char = i * CHARS_PER_TOKEN_ESTIMATE
                end_char = (i + max_tokens) * CHARS_PER_TOKEN_ESTIMATE
                chunks.append(text[start_char:end_char])

        return chunks

    except Exception:
        return _character_based_chunking(text, max_tokens)


def _character_based_chunking(text, max_tokens):
    """Fallback chunking method based on character count."""
    max_chars = max_tokens * CHARS_PER_TOKEN_ESTIMATE

    if len(text) <= max_chars:
        return [text]

    chunks = []
    remaining_text = text

    while remaining_text:
        if len(remaining_text) <= max_chars:
            chunks.append(remaining_text)
            break

        chunk = remaining_text[:max_chars]
        last_period = chunk.rfind('. ')
        if last_period > max_chars * 0.7:
            chunks.append(chunk[:last_period + 1])
            remaining_text = remaining_text[last_period + 1:].lstrip()
        else:
            chunks.append(chunk)
            remaining_text = remaining_text[max_chars:]

    return chunks


def validate_token_length(text):
    """Ensure text doesn't exceed model limits."""
    if not tokenizer:
        return text

    try:
        tokens = tokenizer.encode(text, add_special_tokens=False)
        if len(tokens) > MAX_TOKEN_LIMIT:
            truncated_tokens = tokens[:MAX_TOKEN_LIMIT]
            return tokenizer.decode(truncated_tokens, skip_special_tokens=True)
    except Exception:
        pass

    return text


def summarize_single_chunk(text, max_length, min_length, config):
    """Summarize a single text chunk with error handling."""
    try:
        validated_text = validate_token_length(text)

        result = summarizer(
            validated_text,
            max_length=max_length,
            min_length=min_length,
            do_sample=True,
            truncation=True,
            **config
        )

        return result[0]['summary_text']

    except Exception:
        sentences = text.split('. ')[:3]
        return '. '.join(sentences) + '.' if sentences else "Summary unavailable."


def format_summary_by_style(text, style):
    """Format summary text according to specified style."""
    if not text or style not in ['bullets', 'action', 'highlights']:
        return text.strip()

    sentences = [s.strip().rstrip('.') for s in text.replace('.', '.\n').split('\n') if s.strip()]

    if style == "bullets":
        return '\n'.join(f"• {s}" for s in sentences) if len(sentences) > 1 else f"• {text}"

    elif style == "action":
        formatted = []
        for sentence in sentences:
            if sentence:
                has_action = any(keyword in sentence.lower() for keyword in ACTION_KEYWORDS)
                icon = "🔹" if has_action else "•"
                formatted.append(f"{icon} {sentence}")
        return '\n'.join(formatted) if formatted else f"🔹 {text}"

    elif style == "highlights":
        long_sentences = [s for s in sentences if len(s.strip()) > 10]
        return '\n'.join(f"⭐ {s}" for s in long_sentences) if len(long_sentences) > 1 else f"⭐ {text}"

    return text.strip()


def generate_summary(content, style, max_length, min_length):
    """Main summary generation function."""
    if not model_loaded or not summarizer:
        return "Error: AI model not available"

    if len(content.strip()) < 50:
        return "Text too short to summarize"

    try:
        cleaned_text = preprocess_text(content)
        chunks = create_text_chunks(cleaned_text)
        config = SUMMARY_CONFIGS.get(style, SUMMARY_CONFIGS["short"])

        if len(chunks) == 1:
            summary = summarize_single_chunk(chunks[0], max_length, min_length, config)
            return format_summary_by_style(summary, style)

        chunk_summaries = []
        chunk_max_length = min(100, max_length // len(chunks) + 30)
        chunk_min_length = max(15, min_length // len(chunks))

        for chunk in chunks:
            chunk_config = config.copy()
            chunk_config['temperature'] = max(0.1, min(1.0,
                                                       config['temp'] + random.uniform(-0.1, 0.1)))

            try:
                summary = summarize_single_chunk(chunk, chunk_max_length, chunk_min_length, chunk_config)
                chunk_summaries.append(summary)
            except Exception:
                # Extract fallback summary
                sentences = chunk.split('. ')[:2]
                chunk_summaries.append('. '.join(sentences) + '.' if sentences else "")

        combined_summary = " ".join(filter(None, chunk_summaries))

        if tokenizer:
            try:
                combined_tokens = tokenizer.encode(combined_summary, add_special_tokens=False)
                if len(combined_tokens) > max_length * 2:
                    # Final summarization round
                    final_chunks = create_text_chunks(combined_summary)
                    if len(final_chunks) > 1:
                        # Keep most important parts (first and last)
                        important_parts = [final_chunks[0]]
                        if len(final_chunks) > 1:
                            important_parts.append(final_chunks[-1])
                        combined_summary = " ".join(important_parts)

                    final_summary = summarize_single_chunk(combined_summary, max_length, min_length, config)
                    return format_summary_by_style(final_summary, style)
            except Exception:
                pass

        return format_summary_by_style(combined_summary, style)

    except Exception as e:
        return f"Error generating summary: Unable to process content. {str(e)[:100]}"


def process_attachments(attachments):
    """Process and extract text from attachments."""
    processed = []
    for attachment in attachments:
        filename = attachment.get('name', 'unknown')
        file_data = attachment.get('content', '')

        if file_data:
            content = extract_text_from_file(file_data, filename)
            processed.append({'filename': filename, 'content': content})

    return processed


def create_summary_sections(email_body, attachments, summary_type, attachment_mode):
    """Create structured summary sections."""
    sections = []

    if email_body and email_body.strip():
        email_summary = generate_summary(email_body, summary_type, 200, 50)
        sections.append({'title': 'Email Summary', 'content': email_summary})

    if attachments:
        if attachment_mode == 'separate':
            for att in attachments:
                if att['content'] and att['content'].strip() and not att['content'].startswith('[Error'):
                    att_summary = generate_summary(att['content'], summary_type, 150, 30)
                    sections.append({'title': f"Attachment: {att['filename']}", 'content': att_summary})
        else:
            valid_attachments = [att for att in attachments
                                 if
                                 att['content'] and att['content'].strip() and not att['content'].startswith('[Error')]

            if valid_attachments:
                combined_content = "\n\n".join([f"--- {att['filename']} ---\n{att['content']}"
                                                for att in valid_attachments])
                att_summary = generate_summary(combined_content, summary_type, 200, 50)
                sections.append({'title': f"Attachments ({len(valid_attachments)} files)", 'content': att_summary})

    return sections


initialize_model()


@app.route('/', methods=['GET'])
def home():
    return jsonify({
        "message": "Email Summarization API is running",
        "status": "active",
        "endpoints": {"/health": "Health check", "/summarize": "POST - Summarize content",
                      "/test": "GET - Test endpoint"},
        "timestamp": datetime.now().isoformat()
    })


@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({
        'status': 'healthy' if model_loaded else 'degraded',
        'model_loaded': model_loaded,
        'timestamp': datetime.now().isoformat()
    })


@app.route('/test', methods=['GET'])
def test_endpoint():
    test_text = "This is a test email to verify the summarization functionality is working correctly."

    try:
        test_summary = generate_summary(test_text, "short", 50, 20) if model_loaded else "Model not available"
    except Exception as e:
        test_summary = f"Test failed: {str(e)}"

    return jsonify({
        "message": "Test endpoint working",
        "model_status": "loaded" if model_loaded else "not loaded",
        "gpu_available": torch.cuda.is_available(),
        "test_summary": test_summary
    })


@app.route('/summarize', methods=['POST'])
def summarize_email():
    try:
        if not model_loaded:
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
        processed_attachments = process_attachments(attachments) if include_attachments else []
        max_length, min_length = LENGTH_MAP.get(summary_type, (150, 50))

        if output_mode == 'sections':
            sections = create_summary_sections(
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
                valid_attachments = [att for att in processed_attachments
                                     if att['content'].strip() and not att['content'].startswith('[Error')]

                if attachment_mode == 'separate':
                    att_summaries = []
                    for att in valid_attachments:
                        att_sum = generate_summary(att['content'], summary_type, 100, 20)
                        att_summaries.append(f"{att['filename']}: {att_sum}")
                    if att_summaries:
                        content_parts.append(" | ".join(att_summaries))
                else:
                    combined_att = "\n\n".join([f"--- {att['filename']} ---\n{att['content']}"
                                                for att in valid_attachments])
                    if combined_att.strip():
                        content_parts.append(combined_att)

            if not content_parts:
                return jsonify({'success': False, 'error': 'No valid content to summarize'}), 400

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
            'error': 'Processing failed. Content may be too long or contain unsupported characters.',
            'timestamp': datetime.now().isoformat()
        }), 500

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