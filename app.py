from flask import Flask, request, jsonify
from flask_cors import CORS
import base64
import PyPDF2
import docx
import io
import uuid
import re
import random
from datetime import datetime
from transformers import pipeline, AutoTokenizer, AutoModelForSeq2SeqLM
import torch

app = Flask(__name__)
CORS(app)  # Enable CORS for Outlook add-in

# Initialize BART model for summarization
print("Loading BART model for summarization...")
try:
    # Use the CNN/DailyMail version of BART, specifically fine-tuned for summarization
    model_name = "facebook/bart-large-cnn"

    # Check if CUDA is available
    device = 0 if torch.cuda.is_available() else -1
    print(f"Using device: {'GPU' if device == 0 else 'CPU'}")

    # Initialize the summarization pipeline
    summarizer = pipeline(
        "summarization",
        model=model_name,
        tokenizer=model_name,
        device=device,
        framework="pt"
    )

    # Also load tokenizer for text length checking
    tokenizer = AutoTokenizer.from_pretrained(model_name)

    print("BART model loaded successfully!")

except Exception as e:
    print(f"Error loading BART model: {e}")
    print("Falling back to a smaller model...")
    # Fallback to a smaller model if the large one fails
    try:
        model_name = "facebook/bart-base"
        summarizer = pipeline("summarization", model=model_name, device=device)
        tokenizer = AutoTokenizer.from_pretrained(model_name)
        print("BART base model loaded successfully!")
    except Exception as e2:
        print(f"Failed to load any BART model: {e2}")
        summarizer = None
        tokenizer = None


def extract_text_from_attachment(file_data, filename):
    """Extract text from various attachment types"""
    try:
        file_bytes = base64.b64decode(file_data)

        if filename.lower().endswith('.pdf'):
            pdf_reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()
            return text

        elif filename.lower().endswith('.docx'):
            doc = docx.Document(io.BytesIO(file_bytes))
            text = ""
            for paragraph in doc.paragraphs:
                text += paragraph.text + "\n"
            return text

        elif filename.lower().endswith('.txt'):
            return file_bytes.decode('utf-8')

        else:
            return f"[Attachment: {filename} - Content type not supported for text extraction]"

    except Exception as e:
        return f"[Error processing {filename}: {str(e)}]"


def clean_and_preprocess_text(text):
    """Clean and preprocess text for better summarization"""
    # Remove extra whitespace and normalize
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()

    # Remove very short lines that might be formatting artifacts
    lines = text.split('\n')
    meaningful_lines = [line.strip() for line in lines if len(line.strip()) > 10]

    return ' '.join(meaningful_lines)


def chunk_text(text, max_tokens=900):
    """Split text into chunks that fit within BART's token limit"""
    if not tokenizer:
        # Fallback: split by characters if tokenizer not available
        max_chars = max_tokens * 4  # Rough estimate
        chunks = [text[i:i + max_chars] for i in range(0, len(text), max_chars)]
        return chunks

    # Tokenize the text
    tokens = tokenizer.encode(text, add_special_tokens=False)

    if len(tokens) <= max_tokens:
        return [text]

    # Split into chunks
    chunks = []
    current_pos = 0

    while current_pos < len(tokens):
        # Get chunk tokens
        chunk_tokens = tokens[current_pos:current_pos + max_tokens]

        # Decode back to text
        chunk_text = tokenizer.decode(chunk_tokens, skip_special_tokens=True)
        chunks.append(chunk_text)

        current_pos += max_tokens

    return chunks


def summarize_with_bart(text, max_length=150, min_length=50, style="general"):
    """Generate summary using BART model with style-specific parameters"""
    if not summarizer:
        return "Error: BART model not loaded. Please check the server logs."

    try:
        # Clean and preprocess the text
        cleaned_text = clean_and_preprocess_text(text)

        if len(cleaned_text.strip()) < 50:
            return "Text too short to summarize meaningfully."

        # Add style-specific prefixes to guide the model
        style_prompts = {
            "concise": "Summarize briefly: ",
            "detailed": "Provide a comprehensive summary: ",
            "bullets": "List the main points: ",
            "actions": "Identify action items and tasks: ",
            "highlights": "Extract key highlights: ",
            "executive": "Executive summary with key decisions: ",
            "general": ""
        }

        # Prepend style prompt if available
        prompt = style_prompts.get(style, "")
        if prompt:
            cleaned_text = prompt + cleaned_text

        # Adjust parameters based on style
        style_params = get_style_parameters(style)

        # Check text length and chunk if necessary
        chunks = chunk_text(cleaned_text)

        if len(chunks) == 1:
            # Single chunk - direct summarization with style-specific settings
            summary_result = summarizer(
                cleaned_text,
                max_length=max_length,
                min_length=min_length,
                do_sample=True,
                temperature=style_params['temperature'],
                top_p=style_params['top_p'],
                top_k=style_params['top_k'],
                num_beams=style_params['num_beams'],
                repetition_penalty=style_params['repetition_penalty'],
                length_penalty=style_params['length_penalty'],
                early_stopping=False,
                no_repeat_ngram_size=3
            )
            return post_process_by_style(summary_result[0]['summary_text'], style)

        else:
            # Multiple chunks - summarize each with style-specific settings
            chunk_summaries = []

            for i, chunk in enumerate(chunks):
                try:
                    # Use slightly varied temperature for each chunk
                    chunk_temp = style_params['temperature'] + (random.random() * 0.2 - 0.1)
                    chunk_temp = max(0.1, min(1.0, chunk_temp))  # Keep within bounds

                    chunk_summary = summarizer(
                        chunk,
                        max_length=min(120, max_length // len(chunks) + 50),
                        min_length=max(20, min_length // len(chunks)),
                        do_sample=True,
                        temperature=chunk_temp,
                        top_p=style_params['top_p'],
                        top_k=style_params['top_k'],
                        num_beams=style_params['num_beams'],
                        repetition_penalty=style_params['repetition_penalty'],
                        no_repeat_ngram_size=2
                    )
                    chunk_summaries.append(chunk_summary[0]['summary_text'])
                except Exception as e:
                    chunk_summaries.append(f"[Chunk {i + 1} processing error: {str(e)}]")

            # Combine chunk summaries
            combined_summary = " ".join(chunk_summaries)

            # If combined summary is still too long, summarize it again
            if len(combined_summary.split()) > max_length:
                try:
                    final_summary = summarizer(
                        combined_summary,
                        max_length=max_length,
                        min_length=min_length,
                        do_sample=True,
                        temperature=style_params['temperature'],
                        top_p=style_params['top_p'],
                        top_k=style_params['top_k'],
                        num_beams=style_params['num_beams'],
                        repetition_penalty=style_params['repetition_penalty'],
                        no_repeat_ngram_size=3
                    )
                    return post_process_by_style(final_summary[0]['summary_text'], style)
                except:
                    return post_process_by_style(combined_summary, style)

            return post_process_by_style(combined_summary, style)

    except Exception as e:
        return f"Error generating {style} summary with BART: {str(e)}"


def get_style_parameters(style):
    """Get BART parameters optimized for different summary styles"""
    params = {
        "concise": {
            "temperature": 0.7,
            "top_p": 0.85,
            "top_k": 40,
            "num_beams": 2,
            "repetition_penalty": 1.3,
            "length_penalty": 1.2  # Favor shorter content
        },
        "detailed": {
            "temperature": 0.8,
            "top_p": 0.92,
            "top_k": 50,
            "num_beams": 1,
            "repetition_penalty": 1.1,
            "length_penalty": 0.8  # Allow longer content
        },
        "bullets": {
            "temperature": 0.6,
            "top_p": 0.85,
            "top_k": 35,
            "num_beams": 2,
            "repetition_penalty": 1.4,
            "length_penalty": 1.0
        },
        "actions": {
            "temperature": 0.7,
            "top_p": 0.88,
            "top_k": 45,
            "num_beams": 2,
            "repetition_penalty": 1.2,
            "length_penalty": 1.0
        },
        "highlights": {
            "temperature": 0.75,
            "top_p": 0.9,
            "top_k": 45,
            "num_beams": 1,
            "repetition_penalty": 1.25,
            "length_penalty": 0.9
        },
        "executive": {
            "temperature": 0.65,
            "top_p": 0.85,
            "top_k": 40,
            "num_beams": 2,
            "repetition_penalty": 1.15,
            "length_penalty": 0.95
        }
    }

    # Default parameters for any unrecognized style
    default = {
        "temperature": 0.8,
        "top_p": 0.9,
        "top_k": 50,
        "num_beams": 1,
        "repetition_penalty": 1.2,
        "length_penalty": 0.9
    }

    return params.get(style, default)


def post_process_by_style(summary, style):
    """Apply style-specific post-processing to the summary"""
    if style == "bullets":
        return format_as_bullets(summary)
    elif style == "actions":
        return format_action_items(summary)
    elif style == "highlights":
        return format_highlights(summary)
    elif style == "executive":
        return format_executive(summary)
    else:
        return summary.strip()


def format_action_items(text):
    """Format text to emphasize action items"""
    # Look for action words and format accordingly
    sentences = [s.strip() for s in text.replace('.', '.\n').split('\n') if s.strip()]

    formatted_actions = []
    action_words = ['should', 'must', 'need', 'require', 'action', 'task', 'follow up', 'schedule', 'contact',
                    'prepare', 'review', 'complete', 'submit']

    for sentence in sentences:
        sentence = sentence.strip().rstrip('.')
        if sentence:
            # Check if sentence contains action words
            has_action = any(action_word in sentence.lower() for action_word in action_words)
            if has_action:
                formatted_actions.append(f"🔹 {sentence}")
            else:
                formatted_actions.append(f"• {sentence}")

    return '\n'.join(formatted_actions) if formatted_actions else f"🔹 {text}"


def format_highlights(text):
    """Format text to emphasize key highlights"""
    sentences = [s.strip() for s in text.replace('.', '.\n').split('\n') if s.strip() and len(s.strip()) > 10]

    if len(sentences) <= 1:
        return f"⭐ {text}"

    highlights = []
    for sentence in sentences:
        sentence = sentence.strip().rstrip('.')
        if sentence:
            highlights.append(f"⭐ {sentence}")

    return '\n'.join(highlights) if highlights else f"⭐ {text}"


def format_executive(text):
    """Format text in executive summary style"""
    # Keep it clean and professional, just ensure proper formatting
    return text.strip()
    """Extract text from various attachment types"""
    try:
        file_bytes = base64.b64decode(file_data)

        if filename.lower().endswith('.pdf'):
            pdf_reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()
            return text

        elif filename.lower().endswith('.docx'):
            doc = docx.Document(io.BytesIO(file_bytes))
            text = ""
            for paragraph in doc.paragraphs:
                text += paragraph.text + "\n"
            return text

        elif filename.lower().endswith('.txt'):
            return file_bytes.decode('utf-8')

        else:
            return f"[Attachment: {filename} - Content type not supported for text extraction]"

    except Exception as e:
        return f"[Error processing {filename}: {str(e)}]"


def generate_flexible_summary(email_body, attachments, summary_type, attachment_mode, output_mode, request_id):
    """Generate summaries with flexible parsing options"""
    try:
        result = {'summary': '', 'summaryData': None}

        # Validate inputs
        has_email = email_body and email_body.strip()
        has_attachments = attachments and len(attachments) > 0

        if not has_email and not has_attachments:
            result['summary'] = "No content available to summarize."
            return result

        # Generate summaries based on output mode
        if output_mode == 'sections':
            # Generate separate sections
            sections = []

            # Email section
            if has_email:
                email_summary = generate_summary_by_type(email_body, '', summary_type, request_id)
                sections.append({
                    'title': '📧 Email Summary',
                    'content': email_summary
                })

            # Attachment sections
            if has_attachments:
                if attachment_mode == 'separate':
                    # Individual attachment summaries
                    for i, attachment in enumerate(attachments):
                        if attachment['content'] and attachment['content'].strip():
                            att_summary = generate_summary_by_type('', attachment['content'], summary_type, request_id)
                            sections.append({
                                'title': f"📎 {attachment['filename']}",
                                'content': att_summary
                            })
                else:
                    # Combined attachments summary
                    combined_att_text = "\n\n".join([
                        f"--- {att['filename']} ---\n{att['content']}"
                        for att in attachments if att['content'] and att['content'].strip()
                    ])
                    if combined_att_text.strip():
                        att_summary = generate_summary_by_type('', combined_att_text, summary_type, request_id)
                        sections.append({
                            'title': f"📎 Attachments ({len(attachments)} files)",
                            'content': att_summary
                        })

            result['summaryData'] = {'sections': sections}

        else:
            # Combined output mode
            content_parts = []

            if has_email:
                content_parts.append(f"Email Content: {email_body.strip()}")

            if has_attachments:
                if attachment_mode == 'separate':
                    # Process each attachment separately but combine in final output
                    attachment_summaries = []
                    for attachment in attachments:
                        if attachment['content'] and attachment['content'].strip():
                            att_summary = generate_summary_by_type('', attachment['content'], summary_type, request_id)
                            attachment_summaries.append(f"{attachment['filename']}: {att_summary}")

                    if attachment_summaries:
                        content_parts.append(f"Individual Attachments: {' | '.join(attachment_summaries)}")

                else:
                    # Combined attachments
                    combined_att_text = "\n\n".join([
                        f"--- {att['filename']} ---\n{att['content']}"
                        for att in attachments if att['content'] and att['content'].strip()
                    ])
                    if combined_att_text.strip():
                        content_parts.append(f"Attachments: {combined_att_text}")

            # Generate single combined summary
            full_content = "\n\n".join(content_parts)
            result['summary'] = generate_summary_by_type(full_content, '', summary_type, request_id)

        return result

    except Exception as e:
        print(f"Error in generate_flexible_summary: {str(e)}")
        return {
            'summary': f"Error generating flexible summary: {str(e)}",
            'summaryData': None
        }


def generate_summary_by_type(email_body, attachments_text, summary_type, request_id):
    """Generate different types of summaries based on user selection"""
    try:
        # Combine email body and attachment content
        content_parts = []

        if email_body and email_body.strip():
            content_parts.append(email_body.strip())

        if attachments_text and attachments_text.strip():
            content_parts.append(attachments_text.strip())

        if not content_parts:
            return "No content found to summarize."

        full_content = "\n\n".join(content_parts)

        # Generate summary based on type
        if summary_type == 'short':
            return generate_short_summary(full_content)
        elif summary_type == 'long':
            return generate_long_summary(full_content)
        elif summary_type == 'bullets':
            return generate_bullet_summary(full_content)
        elif summary_type == 'action':
            return generate_action_summary(full_content)
        elif summary_type == 'highlights':
            return generate_highlights_summary(full_content)
        elif summary_type == 'executive':
            return generate_executive_summary(full_content)
        else:
            # Default to short summary
            return generate_short_summary(full_content)

    except Exception as e:
        return f"Error generating {summary_type} summary: {str(e)}"


def generate_short_summary(content):
    """Generate a brief, concise summary"""
    return summarize_with_bart(
        content,
        max_length=100,
        min_length=30,
        style="concise"
    )


def generate_long_summary(content):
    """Generate a detailed, comprehensive summary"""
    return summarize_with_bart(
        content,
        max_length=300,
        min_length=100,
        style="detailed"
    )


def generate_bullet_summary(content):
    """Generate summary in bullet point format"""
    summary = summarize_with_bart(
        content,
        max_length=200,
        min_length=80,
        style="bullets"
    )

    # Post-process to ensure bullet format
    return format_as_bullets(summary)


def generate_action_summary(content):
    """Generate summary focused on action items and tasks"""
    return summarize_with_bart(
        content,
        max_length=180,
        min_length=60,
        style="actions"
    )


def generate_highlights_summary(content):
    """Generate summary focused on key highlights and important points"""
    return summarize_with_bart(
        content,
        max_length=150,
        min_length=50,
        style="highlights"
    )


def generate_executive_summary(content):
    """Generate executive-style summary with key decisions and outcomes"""
    return summarize_with_bart(
        content,
        max_length=250,
        min_length=80,
        style="executive"
    )


def format_as_bullets(text):
    """Convert text to bullet point format"""
    # Split into sentences and format as bullets
    sentences = [s.strip() for s in text.replace('.', '.\n').split('\n') if s.strip() and len(s.strip()) > 10]

    # Format as bullet points
    if len(sentences) <= 1:
        return f"• {text}"

    bullets = []
    for sentence in sentences:
        # Clean up the sentence
        sentence = sentence.strip().rstrip('.')
        if sentence:
            bullets.append(f"• {sentence}")

    return '\n'.join(bullets) if bullets else f"• {text}"
    """Generate AI summary of email and attachments using BART with fresh content each time"""
    try:
        # Combine email body and attachment content
        content_parts = []

        if email_body and email_body.strip():
            content_parts.append(f"Email Content: {email_body.strip()}")

        if attachments_text and attachments_text.strip():
            content_parts.append(f"Attachments: {attachments_text.strip()}")

        if not content_parts:
            return "No content found to summarize."

        full_content = "\n\n".join(content_parts)

        # Generate summary using BART with high variability for fresh content
        summary_text = summarize_with_bart(full_content, max_length=200, min_length=50)

        # Return clean summary without metadata
        return summary_text

    except Exception as e:
        return f"Error generating summary: {str(e)}"


@app.route('/summarize', methods=['POST'])
def summarize_email():
    """Main endpoint to summarize email and attachments with flexible parsing options"""
    try:
        data = request.json
        email_body = data.get('emailBody', '')
        attachments = data.get('attachments', [])
        summary_type = data.get('summaryType', 'short')
        parsing_options = data.get('parsingOptions', {})

        # Extract parsing preferences
        include_email = parsing_options.get('includeEmail', True)
        include_attachments = parsing_options.get('includeAttachments', True)
        attachment_mode = parsing_options.get('attachmentMode', 'combined')  # 'combined' or 'separate'
        output_mode = parsing_options.get('outputMode', 'combined')  # 'combined' or 'sections'

        print(f"Processing request with options: {parsing_options}")

        # Generate unique request ID for fresh responses
        request_id = str(uuid.uuid4())

        # Process attachments
        processed_attachments = []
        for attachment in attachments:
            filename = attachment.get('name', 'unknown')
            file_data = attachment.get('content', '')

            if file_data:
                attachment_text = extract_text_from_attachment(file_data, filename)
                processed_attachments.append({
                    'filename': filename,
                    'content': attachment_text
                })

        # Generate summaries based on parsing options
        result = generate_flexible_summary(
            email_body=email_body if include_email else '',
            attachments=processed_attachments if include_attachments else [],
            summary_type=summary_type,
            attachment_mode=attachment_mode,
            output_mode=output_mode,
            request_id=request_id
        )

        return jsonify({
            'success': True,
            'summary': result.get('summary', ''),
            'summaryData': result.get('summaryData', None),
            'summaryType': summary_type,
            'parsingOptions': parsing_options,
            'requestId': request_id,
            'timestamp': datetime.now().isoformat(),
            'attachmentsProcessed': len(processed_attachments)
        })

    except Exception as e:
        print(f"Error in summarize_email: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e),
            'timestamp': datetime.now().isoformat()
        }), 500


@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat()
    })


if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1', port=5000)