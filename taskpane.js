Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("summarize-btn").onclick = summarizeEmail;
    }
});

async function summarizeEmail() {
    const loadingDiv = document.getElementById("loading");
    const errorDiv = document.getElementById("error");
    const summaryContainer = document.getElementById("summary-container");

    // Show loading state
    loadingDiv.style.display = "block";
    errorDiv.style.display = "none";
    summaryContainer.style.display = "none";

    try {
        // Get email data using Office.js
        const emailData = await getEmailData();

        // Send to Python backend for AI processing
        const summary = await generateSummary(emailData);

        // Display results
        displaySummary(emailData, summary);

        loadingDiv.style.display = "none";
        summaryContainer.style.display = "block";

    } catch (error) {
        console.error("Error summarizing email:", error);
        loadingDiv.style.display = "none";
        errorDiv.textContent = "Failed to summarize email: " + error.message;
        errorDiv.style.display = "block";
    }
}

function getEmailData() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync("text", async (bodyResult) => {
            if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error("Failed to get email body"));
                return;
            }

            try {
                // Get attachments
                const attachments = Office.context.mailbox.item.attachments || [];
                const attachmentNames = attachments.map(att => att.name);

                // Get attachment content for supported types
                const attachmentContents = [];
                for (const attachment of attachments) {
                    if (attachment.attachmentType === Office.MailboxEnums.AttachmentType.File) {
                        try {
                            const content = await getAttachmentContent(attachment.id);
                            attachmentContents.push({
                                name: attachment.name,
                                content: content,
                                contentType: attachment.contentType
                            });
                        } catch (err) {
                            console.warn("Failed to get content for attachment ${attachment.name}:", err);
                        }
                    }
                }

                const emailData = {
                    subject: Office.context.mailbox.item.subject,
                    body: bodyResult.value,
                    from: Office.context.mailbox.item.from.displayName + " <" + Office.context.mailbox.item.from.emailAddress + ">",
                    attachmentNames: attachmentNames,
                    attachmentContents: attachmentContents,
                    dateTimeReceived: Office.context.mailbox.item.dateTimeCreated
                };

                resolve(emailData);

            } catch (error) {
                reject(error);
            }
        });
    });
}

function getAttachmentContent(attachmentId) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value.content);
            } else {
                reject(new Error("Failed to get attachment content"));
            }
        });
    });
}

async function generateSummary(emailData) {
    const response = await fetch('http://localhost:5000/api/summarize', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(emailData)
    });

    if (!response.ok) {
        throw new Error("HTTP error! status: ${response.status}");
    }

    const result = await response.json();
    return result.summary;
}

function displaySummary(emailData, summary) {
    document.getElementById("summary-text").textContent = summary;
    document.getElementById("email-subject").textContent = emailData.subject;
    document.getElementById("email-from").textContent = emailData.from;

    const attachmentText = emailData.attachmentNames.length > 0
        ? emailData.attachmentNames.join(", ")
        : "None";
    document.getElementById("email-attachments").textContent = attachmentText;
}