// Python backend URL
const BACKEND_URL = 'http://127.0.0.1:5000';

// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Wait for DOM to be fully loaded
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', initializeButtons);
        } else {
            initializeButtons();
        }
        console.log('Email Summarizer add-in loaded successfully');
    } else {
        console.error('This add-in only works in Outlook');
    }
});

function initializeButtons() {
    const summarizeBtn = document.getElementById('summarize-btn');
    const copyBtn = document.getElementById('copy-btn');
    
    if (summarizeBtn) {
        summarizeBtn.onclick = generateSummary;
    } else {
        console.error('Summarize button not found');
    }
    
    if (copyBtn) {
        copyBtn.onclick = copySummary;
    } else {
        console.error('Copy button not found');
    }
    
    // Initialize parsing option event listeners
    initializeParsingOptions();
}

function initializeParsingOptions() {
    // Handle attachment checkbox toggle
    const includeAttachments = document.getElementById('include-attachments');
    const attachmentOptions = document.getElementById('attachment-options');
    
    if (includeAttachments && attachmentOptions) {
        includeAttachments.addEventListener('change', function() {
            if (this.checked) {
                attachmentOptions.classList.remove('disabled');
            } else {
                attachmentOptions.classList.add('disabled');
            }
        });
    }
    
    // Handle content selection logic
    const includeEmail = document.getElementById('include-email');
    const sectionOptions = document.getElementById('section-options');
    
    if (includeEmail && includeAttachments && sectionOptions) {
        function updateSectionOptions() {
            const emailChecked = includeEmail.checked;
            const attachmentsChecked = includeAttachments.checked;
            
            // Show section options only when both email and attachments are selected
            if (emailChecked && attachmentsChecked) {
                sectionOptions.classList.remove('disabled');
            } else {
                sectionOptions.classList.add('disabled');
                // Auto-select combined when only one content type is selected
                document.querySelector('input[name="output-mode"][value="combined"]').checked = true;
            }
        }
        
        includeEmail.addEventListener('change', updateSectionOptions);
        includeAttachments.addEventListener('change', updateSectionOptions);
        
        // Initial state
        updateSectionOptions();
    }
}

// Show/hide UI elements
function showElement(elementId) {
    document.getElementById(elementId).style.display = 'block';
}

function hideElement(elementId) {
    document.getElementById(elementId).style.display = 'none';
}

function showMessage(type, message) {
    // Hide all message types first
    hideElement('info-bar');
    hideElement('error-message');
    hideElement('success-message');
    
    // Show the appropriate message with correct element ID mapping
    let elementId;
    if (type === 'info') {
        elementId = 'info-bar';
    } else {
        elementId = type + '-message';
    }
    
    const messageElement = document.getElementById(elementId);
    if (messageElement) {
        messageElement.textContent = message;
        showElement(elementId);
        
        // Auto-hide success messages after 3 seconds
        if (type === 'success') {
            setTimeout(() => hideElement(elementId), 3000);
        }
    } else {
        console.error('Message element not found:', elementId);
    }
}

// Get email content from Outlook
function getEmailContent() {
    return new Promise((resolve, reject) => {
        // Get email body
        Office.context.mailbox.item.body.getAsync(
            Office.CoercionType.Text,
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const emailBody = result.value;
                    
                    // Get attachments
                    getAttachments()
                        .then(attachments => {
                            resolve({
                                emailBody: emailBody,
                                attachments: attachments
                            });
                        })
                        .catch(reject);
                } else {
                    reject(new Error('Failed to get email body: ' + result.error.message));
                }
            }
        );
    });
}

// Get email attachments
function getAttachments() {
    return new Promise((resolve, reject) => {
        const attachments = Office.context.mailbox.item.attachments;
        
        if (!attachments || attachments.length === 0) {
            resolve([]);
            return;
        }
        
        const attachmentPromises = attachments.map(attachment => {
            return new Promise((resolveAttachment, rejectAttachment) => {
                if (attachment.attachmentType === Office.MailboxEnums.AttachmentType.File) {
                    // Get attachment content
                    Office.context.mailbox.item.getAttachmentContentAsync(
                        attachment.id,
                        (result) => {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                resolveAttachment({
                                    name: attachment.name,
                                    content: result.value.content,
                                    contentType: result.value.format
                                });
                            } else {
                                console.warn(`Failed to get attachment ${attachment.name}: ${result.error.message}`);
                                resolveAttachment({
                                    name: attachment.name,
                                    content: null,
                                    error: result.error.message
                                });
                            }
                        }
                    );
                } else {
                    // Item attachment (not supported for content extraction)
                    resolveAttachment({
                        name: attachment.name,
                        content: null,
                        type: 'item'
                    });
                }
            });
        });
        
        Promise.all(attachmentPromises)
            .then(resolve)
            .catch(reject);
    });
}

// Send data to Python backend for summarization
async function sendToBackend(emailData) {
    try {
        // Get selected summary type
        const summaryType = document.getElementById('summary-type').value;
        
        // Get parsing preferences
        const includeEmail = document.getElementById('include-email').checked;
        const includeAttachments = document.getElementById('include-attachments').checked;
        
        const attachmentMode = document.querySelector('input[name="attachment-mode"]:checked')?.value || 'combined';
        const outputMode = document.querySelector('input[name="output-mode"]:checked')?.value || 'combined';
        
        // Add parsing preferences to the data being sent
        const requestData = {
            ...emailData,
            summaryType: summaryType,
            parsingOptions: {
                includeEmail: includeEmail,
                includeAttachments: includeAttachments,
                attachmentMode: attachmentMode, // 'combined' or 'separate'
                outputMode: outputMode // 'combined' or 'sections'
            }
        };
        
        console.log('Sending parsing options:', requestData.parsingOptions);
        
        const response = await fetch(`${BACKEND_URL}/summarize`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(requestData)
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const result = await response.json();
        
        if (!result.success) {
            throw new Error(result.error || 'Unknown error occurred');
        }
        
        return result;
    } catch (error) {
        if (error.name === 'TypeError' && error.message.includes('fetch')) {
            throw new Error('Cannot connect to backend server. Please make sure your Python Flask server is running.');
        }
        throw error;
    }
}

// Main function to generate summary
async function generateSummary() {
    const summarizeBtn = document.getElementById('summarize-btn');
    const summaryContainer = document.getElementById('summary-container');
    const summaryContent = document.getElementById('summary-content');
    const summaryTypeDropdown = document.getElementById('summary-type');
    
    // Check if required elements exist
    if (!summarizeBtn) {
        console.error('Summarize button not found');
        return;
    }
    
    if (!summaryContent) {
        console.error('Summary content element not found');
        showMessage('error', 'UI elements not found. Please refresh the add-in.');
        return;
    }
    
    if (!summaryTypeDropdown) {
        console.error('Summary type dropdown not found');
        showMessage('error', 'Summary type selector not found. Please refresh the add-in.');
        return;
    }
    
    // Validate parsing options
    const includeEmail = document.getElementById('include-email')?.checked || false;
    const includeAttachments = document.getElementById('include-attachments')?.checked || false;
    
    if (!includeEmail && !includeAttachments) {
        showMessage('error', 'Please select at least one content type to summarize (Email Body or Attachments).');
        return;
    }
    
    try {
        // Get selected summary type for display
        const selectedOption = summaryTypeDropdown.options[summaryTypeDropdown.selectedIndex];
        const summaryTypeName = selectedOption.text;
        
        // Get parsing options for display
        const outputMode = document.querySelector('input[name="output-mode"]:checked')?.value || 'combined';
        const attachmentMode = document.querySelector('input[name="attachment-mode"]:checked')?.value || 'combined';
        
        // Disable controls during processing
        summarizeBtn.disabled = true;
        summarizeBtn.textContent = 'Processing...';
        summaryTypeDropdown.disabled = true;
        disableParsingControls(true);
        
        showElement('loading');
        hideElement('summary-container');
        hideElement('info-bar');
        hideElement('error-message');
        hideElement('success-message');
        
        // Get email content
        showMessage('info', 'Extracting content...');
        const emailData = await getEmailContent();
        
        // Build progress message based on selections
        let contentDescription = [];
        if (includeEmail) contentDescription.push('email body');
        if (includeAttachments) {
            if (emailData.attachments.length > 0) {
                if (attachmentMode === 'separate') {
                    contentDescription.push(`${emailData.attachments.length} attachments (individually)`);
                } else {
                    contentDescription.push(`${emailData.attachments.length} attachments (combined)`);
                }
            } else {
                contentDescription.push('attachments (none found)');
            }
        }
        
        const processingMsg = `Generating ${summaryTypeName} for: ${contentDescription.join(' + ')}${outputMode === 'sections' ? ' (separate sections)' : ''}...`;
        showMessage('info', processingMsg);
        
        // Send to backend for processing
        const result = await sendToBackend(emailData);
        
        // Display the summary/summaries
        if (result.summaryData && result.summaryData.sections) {
            // Handle sectioned output
            displaySectionedSummaries(result.summaryData.sections);
        } else {
            // Handle single summary
            summaryContent.textContent = result.summary || 'No summary generated.';
        }
        
        if (summaryContainer) {
            showElement('summary-container');
        }
        
        // Show success message with details
        let successMessage = `${summaryTypeName} generated successfully`;
        if (result.attachmentsProcessed > 0) {
            successMessage += ` (${result.attachmentsProcessed} attachments processed)`;
        }
        showMessage('success', successMessage);
        
    } catch (error) {
        console.error('Error generating summary:', error);
        showMessage('error', error.message);
    } finally {
        // Reset controls
        summarizeBtn.disabled = false;
        summarizeBtn.textContent = 'Generate Summary';
        summaryTypeDropdown.disabled = false;
        disableParsingControls(false);
        hideElement('loading');
    }
}

function disableParsingControls(disabled) {
    // Disable/enable all parsing option controls
    const controls = [
        'include-email',
        'include-attachments',
        ...document.querySelectorAll('input[name="attachment-mode"]'),
        ...document.querySelectorAll('input[name="output-mode"]')
    ];
    
    controls.forEach(controlId => {
        const element = typeof controlId === 'string' ? document.getElementById(controlId) : controlId;
        if (element) {
            element.disabled = disabled;
        }
    });
}

function displaySectionedSummaries(sections) {
    const summaryContent = document.getElementById('summary-content');
    if (!summaryContent) return;
    
    let formattedOutput = '';
    
    sections.forEach((section, index) => {
        if (section.title) {
            formattedOutput += `\n${'='.repeat(50)}\n`;
            formattedOutput += `${section.title.toUpperCase()}\n`;
            formattedOutput += `${'='.repeat(50)}\n\n`;
        }
        
        formattedOutput += section.content;
        
        // Add spacing between sections (except for the last one)
        if (index < sections.length - 1) {
            formattedOutput += '\n\n';
        }
    });
    
    summaryContent.textContent = formattedOutput.trim();
}

// Copy summary to clipboard
async function copySummary() {
    const summaryContent = document.getElementById('summary-content');
    const copyBtn = document.getElementById('copy-btn');
    
    if (!summaryContent) {
        console.error('Summary content element not found');
        showMessage('error', 'Cannot find summary to copy');
        return;
    }
    
    if (!copyBtn) {
        console.error('Copy button not found');
        return;
    }
    
    const summaryText = summaryContent.textContent;
    
    if (!summaryText || summaryText.trim() === '') {
        showMessage('error', 'No summary content to copy');
        return;
    }
    
    try {
        let copySuccess = false;
        
        // Method 1: Modern Clipboard API (preferred for HTTPS)
        if (navigator.clipboard && navigator.clipboard.writeText) {
            try {
                await navigator.clipboard.writeText(summaryText);
                copySuccess = true;
                console.log('Copy successful via Clipboard API');
            } catch (clipboardError) {
                console.log('Clipboard API failed:', clipboardError);
                // Continue to fallback methods
            }
        }
        
        // Method 2: execCommand fallback (works in more environments)
        if (!copySuccess) {
            try {
                const textArea = document.createElement('textarea');
                textArea.value = summaryText;
                
                // Position off-screen
                textArea.style.position = 'fixed';
                textArea.style.left = '-9999px';
                textArea.style.top = '-9999px';
                textArea.style.opacity = '0';
                textArea.style.pointerEvents = 'none';
                
                // Add to DOM
                document.body.appendChild(textArea);
                
                // Select and focus
                textArea.focus();
                textArea.select();
                textArea.setSelectionRange(0, summaryText.length);
                
                // Try copy command
                const successful = document.execCommand('copy');
                
                // Remove from DOM
                document.body.removeChild(textArea);
                
                if (successful) {
                    copySuccess = true;
                    console.log('Copy successful via execCommand');
                } else {
                    console.log('execCommand copy failed');
                }
            } catch (execError) {
                console.log('execCommand method failed:', execError);
            }
        }
        
        // Method 3: Office.js clipboard (Outlook-specific)
        if (!copySuccess && typeof Office !== 'undefined' && Office.context) {
            try {
                // For Outlook add-ins, try Office clipboard
                if (Office.context.ui && Office.context.ui.displayDialogAsync) {
                    // This is a more complex method that would require a dialog
                    // For now, we'll skip this and go to manual method
                    console.log('Office.js clipboard not implemented in this version');
                } else {
                    console.log('Office.js clipboard API not available');
                }
            } catch (officeError) {
                console.log('Office.js clipboard failed:', officeError);
            }
        }
        
        // Method 4: Create a selection for manual copy
        if (!copySuccess) {
            try {
                // Select the summary text for manual copying
                const range = document.createRange();
                range.selectNodeContents(summaryContent);
                const selection = window.getSelection();
                selection.removeAllRanges();
                selection.addRange(range);
                
                copySuccess = true; // Consider this successful as text is selected
                console.log('Text selected for manual copy');
                
                // Show instructions
                showMessage('success', 'Summary text selected! Press Ctrl+C (Windows) or Cmd+C (Mac) to copy.');
                
                // Visual feedback for selection method
                const originalText = copyBtn.textContent;
                copyBtn.textContent = '📋 Selected!';
                copyBtn.style.backgroundColor = '#17a2b8'; // Info color
                
                setTimeout(() => {
                    copyBtn.textContent = originalText;
                    copyBtn.style.backgroundColor = '#28a745';
                    // Clear selection after a moment
                    if (window.getSelection) {
                        window.getSelection().removeAllRanges();
                    }
                }, 3000);
                
                return; // Exit early for selection method
            } catch (selectionError) {
                console.log('Selection method failed:', selectionError);
            }
        }
        
        if (copySuccess) {
            // Visual feedback for successful copy
            const originalText = copyBtn.textContent;
            const originalColor = copyBtn.style.backgroundColor;
            
            copyBtn.textContent = '✅ Copied!';
            copyBtn.style.backgroundColor = '#28a745';
            
            setTimeout(() => {
                copyBtn.textContent = originalText;
                copyBtn.style.backgroundColor = originalColor || '#28a745';
            }, 2000);
            
            showMessage('success', 'Summary copied to clipboard!');
        } else {
            // All methods failed
            throw new Error('All copy methods failed');
        }
        
    } catch (error) {
        console.error('All copy methods failed:', error);
        
        // Final fallback: Show a modal with the text
        showCopyFallbackModal(summaryText);
    }
}

// Fallback modal for copying when all else fails
function showCopyFallbackModal(text) {
    // Remove any existing modal
    const existingModal = document.getElementById('copy-modal');
    if (existingModal) {
        existingModal.remove();
    }
    
    // Create modal
    const modal = document.createElement('div');
    modal.id = 'copy-modal';
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0,0,0,0.7);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 10000;
        padding: 20px;
    `;
    
    const modalContent = document.createElement('div');
    modalContent.style.cssText = `
        background: white;
        border-radius: 8px;
        padding: 20px;
        max-width: 90%;
        max-height: 80%;
        overflow: auto;
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    `;
    
    modalContent.innerHTML = `
        <h3 style="margin-top: 0; color: #2c3e50;">📋 Copy Summary</h3>
        <p style="color: #666; margin-bottom: 15px;">Select all text below and copy it manually:</p>
        <textarea readonly style="
            width: 100%;
            height: 200px;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-family: inherit;
            font-size: 14px;
            resize: vertical;
        " id="copy-modal-text">${text}</textarea>
        <div style="text-align: right; margin-top: 15px;">
            <button onclick="document.getElementById('copy-modal').remove()" style="
                background: #0078d4;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                cursor: pointer;
            ">Close</button>
        </div>
    `;
    
    modal.appendChild(modalContent);
    document.body.appendChild(modal);
    
    // Auto-select the text
    setTimeout(() => {
        const textarea = document.getElementById('copy-modal-text');
        if (textarea) {
            textarea.select();
            textarea.focus();
        }
    }, 100);
    
    // Close on background click
    modal.addEventListener('click', (e) => {
        if (e.target === modal) {
            modal.remove();
        }
    });
    
    showMessage('info', 'Manual copy dialog opened. Select the text and copy with Ctrl+C.');
}

// Utility function for debugging backend connection
function checkBackendConnection() {
    fetch(`${BACKEND_URL}/health`)
        .then(response => response.json())
        .then(data => {
            console.log('Backend connection successful:', data);
            showMessage('success', 'Backend server is running');
        })
        .catch(error => {
            console.error('Backend connection failed:', error);
            showMessage('error', 'Cannot connect to backend server');
        });
}

// Add keyboard shortcuts
document.addEventListener('keydown', (event) => {
    // Ctrl/Cmd + Enter to generate summary
    if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
        generateSummary();
    }
    
    // Ctrl/Cmd + C to copy summary (when summary is visible)
    if ((event.ctrlKey || event.metaKey) && event.key === 'c') {
        const summaryContainer = document.getElementById('summary-container');
        if (summaryContainer && summaryContainer.style.display !== 'none') {
            // Only trigger our copy if the summary is visible and nothing else is selected
            const selection = window.getSelection();
            if (!selection || selection.toString().trim() === '') {
                event.preventDefault();
                copySummary();
            }
        }
    }
    
    // Ctrl/Cmd + 1-6 to select summary types quickly
    if ((event.ctrlKey || event.metaKey) && event.key >= '1' && event.key <= '6') {
        const dropdown = document.getElementById('summary-type');
        if (dropdown) {
            const optionIndex = parseInt(event.key) - 1;
            if (optionIndex < dropdown.options.length) {
                dropdown.selectedIndex = optionIndex;
                event.preventDefault();
            }
        }
    }
    
    // Escape to close modal
    if (event.key === 'Escape') {
        const modal = document.getElementById('copy-modal');
        if (modal) {
            modal.remove();
        }
    }
});

console.log('TaskPane JavaScript loaded');
console.log('Backend URL:', BACKEND_URL);
console.log('Keyboard shortcuts:');
console.log('  Ctrl+Enter: Generate summary');
console.log('  Ctrl+C: Copy summary (when visible)');
console.log('  Ctrl+1-6: Quick style selection');
console.log('  Escape: Close copy modal');