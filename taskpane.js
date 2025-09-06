const BACKEND_URL = 'http://127.0.0.1:5000';

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', init);
        } else {
            init();
        }
    }
});

function init() {
    const summarizeBtn = document.getElementById('summarize-btn');
    const copyBtn = document.getElementById('copy-btn');
    
    if (summarizeBtn) summarizeBtn.onclick = generateSummary;
    if (copyBtn) copyBtn.onclick = copySummary;
    
    initParsingOptions();
}

function initParsingOptions() {
    const includeAttachments = document.getElementById('include-attachments');
    const includeEmail = document.getElementById('include-email');
    const attachmentOptions = document.getElementById('attachment-options');
    const sectionOptions = document.getElementById('section-options');
    
    if (includeAttachments && attachmentOptions) {
        includeAttachments.addEventListener('change', function() {
            attachmentOptions.classList.toggle('disabled', !this.checked);
        });
    }
    
    if (includeEmail && includeAttachments && sectionOptions) {
        function updateSectionOptions() {
            const bothSelected = includeEmail.checked && includeAttachments.checked;
            sectionOptions.classList.toggle('disabled', !bothSelected);
            
            if (!bothSelected) {
                document.querySelector('input[name="output-mode"][value="combined"]').checked = true;
            }
        }
        
        includeEmail.addEventListener('change', updateSectionOptions);
        includeAttachments.addEventListener('change', updateSectionOptions);
        updateSectionOptions();
    }
}

function showElement(id) {
    document.getElementById(id).style.display = 'block';
}

function hideElement(id) {
    document.getElementById(id).style.display = 'none';
}

function showMessage(type, message) {
    ['info-bar', 'error-message', 'success-message'].forEach(hideElement);
    
    const elementId = type === 'info' ? 'info-bar' : type + '-message';
    const element = document.getElementById(elementId);
    
    if (element) {
        element.textContent = message;
        showElement(elementId);
        
        if (type === 'success') {
            setTimeout(() => hideElement(elementId), 3000);
        }
    }
}

function getEmailContent() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                getAttachments()
                    .then(attachments => resolve({ emailBody: result.value, attachments }))
                    .catch(reject);
            } else {
                reject(new Error('Failed to get email body: ' + result.error.message));
            }
        });
    });
}

function getAttachments() {
    return new Promise((resolve) => {
        const attachments = Office.context.mailbox.item.attachments;
        
        if (!attachments || attachments.length === 0) {
            resolve([]);
            return;
        }
        
        const promises = attachments.map(attachment => {
            return new Promise((resolveAtt) => {
                if (attachment.attachmentType === Office.MailboxEnums.AttachmentType.File) {
                    Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, (result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            resolveAtt({
                                name: attachment.name,
                                content: result.value.content,
                                contentType: result.value.format
                            });
                        } else {
                            resolveAtt({
                                name: attachment.name,
                                content: null,
                                error: result.error.message
                            });
                        }
                    });
                } else {
                    resolveAtt({
                        name: attachment.name,
                        content: null,
                        type: 'item'
                    });
                }
            });
        });
        
        Promise.all(promises).then(resolve);
    });
}

async function sendToBackend(emailData) {
    const summaryType = document.getElementById('summary-type').value;
    const includeEmail = document.getElementById('include-email').checked;
    const includeAttachments = document.getElementById('include-attachments').checked;
    const attachmentMode = document.querySelector('input[name="attachment-mode"]:checked')?.value || 'combined';
    const outputMode = document.querySelector('input[name="output-mode"]:checked')?.value || 'combined';
    
    const requestData = {
        ...emailData,
        summaryType,
        parsingOptions: { includeEmail, includeAttachments, attachmentMode, outputMode }
    };
    
    const response = await fetch(`${BACKEND_URL}/summarize`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestData)
    });

    const result = await response.json();

    if (!response.ok) {
        throw new Error(result.error);
    }

    if (!result.success) {
        throw new Error(result.error || 'Unknown error occurred');
    }
    
    return result;
}

async function generateSummary() {
    const summarizeBtn = document.getElementById('summarize-btn');
    const summaryContainer = document.getElementById('summary-container');
    const summaryContent = document.getElementById('summary-content');
    const summaryTypeDropdown = document.getElementById('summary-type');
    
    if (!summarizeBtn || !summaryContent || !summaryTypeDropdown) {
        showMessage('error', 'UI elements not found. Please refresh the add-in.');
        return;
    }
    
    const includeEmail = document.getElementById('include-email')?.checked || false;
    const includeAttachments = document.getElementById('include-attachments')?.checked || false;
    
    if (!includeEmail && !includeAttachments) {
        showMessage('error', 'Please select at least one content type to summarize.');
        return;
    }
    
    try {
        const selectedOption = summaryTypeDropdown.options[summaryTypeDropdown.selectedIndex];
        const summaryTypeName = selectedOption.text;
        const outputMode = document.querySelector('input[name="output-mode"]:checked')?.value || 'combined';
        const attachmentMode = document.querySelector('input[name="attachment-mode"]:checked')?.value || 'combined';
        
        summarizeBtn.disabled = true;
        summarizeBtn.textContent = 'Processing...';
        summaryTypeDropdown.disabled = true;
        toggleParsingControls(true);
        
        showElement('loading');
        ['summary-container', 'info-bar', 'error-message', 'success-message'].forEach(hideElement);
        
        showMessage('info', 'Extracting content...');
        const emailData = await getEmailContent();
        
        const contentDesc = [];
        if (includeEmail) contentDesc.push('email body');
        if (includeAttachments) {
            const attachCount = emailData.attachments.length;
            if (attachCount > 0) {
                const mode = attachmentMode === 'separate' ? 'individually' : 'combined';
                contentDesc.push(`${attachCount} attachments (${mode})`);
            } else {
                contentDesc.push('attachments (none found)');
            }
        }
        
        const processingMsg = `Generating ${summaryTypeName} for: ${contentDesc.join(' + ')}${outputMode === 'sections' ? ' (separate sections)' : ''}...`;
        showMessage('info', processingMsg);
        
        const result = await sendToBackend(emailData);
        
        if (result.summaryData?.sections) {
            displaySections(result.summaryData.sections);
        } else {
            summaryContent.textContent = result.summary || 'No summary generated.';
        }
        
        if (summaryContainer) showElement('summary-container');
        
        let successMsg = `${summaryTypeName} generated successfully`;
        if (result.attachmentsProcessed > 0) {
            successMsg += ` (${result.attachmentsProcessed} attachments processed)`;
        }
        showMessage('success', successMsg);
        
    } catch (error) {
        showMessage('error', error.message);
    } finally {
        summarizeBtn.disabled = false;
        summarizeBtn.textContent = 'Generate Summary';
        summaryTypeDropdown.disabled = false;
        toggleParsingControls(false);
        hideElement('loading');
    }
}

function toggleParsingControls(disabled) {
    const controls = [
        'include-email',
        'include-attachments',
        ...document.querySelectorAll('input[name="attachment-mode"]'),
        ...document.querySelectorAll('input[name="output-mode"]')
    ];
    
    controls.forEach(ctrl => {
        const element = typeof ctrl === 'string' ? document.getElementById(ctrl) : ctrl;
        if (element) element.disabled = disabled;
    });
}

function displaySections(sections) {
    const summaryContent = document.getElementById('summary-content');
    if (!summaryContent) return;
    
    const output = sections.map((section, index) => {
        let text = '';
        if (section.title) {
            text += `\n${'='.repeat(50)}\n${section.title.toUpperCase()}\n${'='.repeat(50)}\n\n`;
        }
        text += section.content;
        return text;
    }).join('\n\n');
    
    summaryContent.textContent = output.trim();
}

async function copySummary() {
    const summaryContent = document.getElementById('summary-content');
    const copyBtn = document.getElementById('copy-btn');
    
    if (!summaryContent || !copyBtn) {
        showMessage('error', 'Copy elements not found');
        return;
    }
    
    const text = summaryContent.textContent;
    if (!text?.trim()) {
        showMessage('error', 'No summary content to copy');
        return;
    }
    
    try {
        if (navigator.clipboard?.writeText) {
            await navigator.clipboard.writeText(text);
            showCopySuccess(copyBtn);
            return;
        }
        
        const textarea = document.createElement('textarea');
        Object.assign(textarea.style, {
            position: 'fixed',
            left: '-9999px',
            opacity: '0'
        });
        textarea.value = text;
        
        document.body.appendChild(textarea);
        textarea.select();
        
        const success = document.execCommand('copy');
        document.body.removeChild(textarea);
        
        if (success) {
            showCopySuccess(copyBtn);
        } else {
            selectTextForManualCopy(summaryContent, copyBtn);
        }
        
    } catch (error) {
        showCopyModal(text);
    }
}

function showCopySuccess(btn) {
    const originalText = btn.textContent;
    const originalColor = btn.style.backgroundColor;
    
    btn.textContent = 'Copied!';
    btn.style.backgroundColor = '#28a745';
    
    setTimeout(() => {
        btn.textContent = originalText;
        btn.style.backgroundColor = originalColor || '#28a745';
    }, 2000);
    
    showMessage('success', 'Summary copied to clipboard!');
}

function selectTextForManualCopy(content, btn) {
    const range = document.createRange();
    range.selectNodeContents(content);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    
    const originalText = btn.textContent;
    btn.textContent = 'Selected!';
    btn.style.backgroundColor = '#17a2b8';
    
    setTimeout(() => {
        btn.textContent = originalText;
        btn.style.backgroundColor = '#28a745';
        if (window.getSelection) window.getSelection().removeAllRanges();
    }, 3000);
    
    showMessage('success', 'Text selected! Press Ctrl+C to copy.');
}

function showCopyModal(text) {
    const existingModal = document.getElementById('copy-modal');
    if (existingModal) existingModal.remove();
    
    const modal = document.createElement('div');
    modal.id = 'copy-modal';
    modal.style.cssText = `
        position: fixed; top: 0; left: 0; right: 0; bottom: 0;
        background: rgba(0,0,0,0.7); display: flex; align-items: center;
        justify-content: center; z-index: 10000; padding: 20px;
    `;
    
    modal.innerHTML = `
        <div style="background: white; border-radius: 8px; padding: 20px; max-width: 90%; max-height: 80%; overflow: auto; box-shadow: 0 4px 20px rgba(0,0,0,0.3);">
            <h3 style="margin-top: 0; color: #2c3e50;">Copy Summary</h3>
            <p style="color: #666; margin-bottom: 15px;">Select all text below and copy manually:</p>
            <textarea readonly style="width: 100%; height: 200px; padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-family: inherit; resize: vertical; box-sizing: border-box;">${text}</textarea>
            <div style="text-align: right; margin-top: 15px;">
                <button onclick="this.closest('#copy-modal').remove()" style="background: #0078d4; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer;">Close</button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    setTimeout(() => {
        const textarea = modal.querySelector('textarea');
        if (textarea) {
            textarea.select();
            textarea.focus();
        }
    }, 100);
    
    modal.addEventListener('click', (e) => {
        if (e.target === modal) modal.remove();
    });
}

function displaySections(sections) {
    const summaryContent = document.getElementById('summary-content');
    if (!summaryContent) return;
    
    const output = sections.map(section => {
        let text = '';
        if (section.title) {
            text += `\n${'='.repeat(50)}\n${section.title.toUpperCase()}\n${'='.repeat(50)}\n\n`;
        }
        text += section.content;
        return text;
    }).join('\n\n');
    
    summaryContent.textContent = output.trim();
}