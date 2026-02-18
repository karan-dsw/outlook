// API URL Configuration
const CLAIMS_API_URL = 'https://metamathematical-mariano-interresponsible.ngrok-free.dev';
// const CLAIMS_API_URL = 'https://demo.datasciencewizards.ai:5006';
// const UNDERWRITING_API_URL = 'https://demo.datasciencewizards.ai:5004';
const UNDERWRITING_API_URL = 'https://corinne-unstudded-uneugenically.ngrok-free.dev';

let mailboxItem = null;
let filename = '';
let processingType = ''; // 'claims' or 'underwriting'
let extractedData = null; // Store extracted data globally


Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        mailboxItem = Office.context.mailbox.item;
        console.log('Office.js initialized successfully');

        // Automatically trigger the flow when taskpane opens
        triggerFlowAndLoadForm();
    }
});

function detectProcessingTypeFromAttachments(attachments) {
    // Check attachment names to determine processing type
    console.log('Checking attachments for processing type detection:');
    for (const att of attachments) {
        const attName = att.name || '';
        console.log(`  - Attachment: "${attName}"`);

        // Check if filename starts with 'C' followed by a number (C1, C2, C3, etc.)
        if (/^c\d/i.test(attName)) {
            console.log(`  ✓ Detected CLAIMS workflow (starts with C + number)`);
            return 'claims';
        } else if (attName.toLowerCase().startsWith('acord')) {
            console.log(`  ✓ Detected UNDERWRITING workflow (starts with acord)`);
            return 'underwriting';
        }
    }
    console.log('  → Defaulting to UNDERWRITING workflow');
    return 'underwriting'; // default
}

// Function to extract form fields directly from email data
function extractFormFieldsFromEmail(emailData) {
    console.log('Extracting form fields from email data...');

    const formFields = {
        policy_number: '',
        document_name: '',
        subject: '',
        comments: '',
        timestamp: ''
    };

    // Extract policy number from subject or body using regex
    const policyRegex = /(?:policy\s*(?:no|number|#)?[:\s]*)?([A-Z]?\d{6,12})/i;
    const subjectMatch = emailData.subject?.match(policyRegex);
    if (subjectMatch) {
        formFields.policy_number = subjectMatch[1];
    }

    // Extract document name from first attachment
    if (emailData.attachments && emailData.attachments.length > 0) {
        // formFields.document_name = emailData.attachments[0].name || '';
        // filename = emailData.attachments[0].name || '';
        let acordFile = emailData.attachments.find(att =>
            att.name && att.name.toLowerCase().startsWith('acord_')
        );
        let selectedAttachment = acordFile || emailData.attachments[0];
        formFields.document_name = selectedAttachment.name || '';
        filename = selectedAttachment.name;
        console.log(`Selected attachment for processing: ${filename}`);
        if (acordFile) {
            console.log('  ✓ ACORD file detected');
        } else {
            console.log('  ⚠ No ACORD file found, using first attachment');
        }
    }

    // Use email subject as subject name
    formFields.subject = emailData.subject || '';

    // Leave comments empty - user will add manually
    formFields.comments = '';

    // Format timestamp from receivedDateTime
    if (emailData.receivedDateTime) {
        const date = new Date(emailData.receivedDateTime);
        formFields.timestamp = date.toLocaleString('en-US', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
        });
    } else {
        const now = new Date();
        formFields.timestamp = now.toLocaleString('en-US', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
        });
    }

    console.log('Extracted form fields:', formFields);
    return formFields;
}

async function triggerFlowAndLoadForm() {
    const loadingContainer = document.getElementById('loadingContainer');
    const formContainer = document.getElementById('formContainer');
    const loadingText = document.querySelector('.loading-text');
    const loadingSubtext = document.querySelector('.loading-subtext');

    try {
        // Show progress notification in email view
        Office.context.mailbox.item.notificationMessages.addAsync(
            "progress",
            {
                type: "progressIndicator",
                message: "Analyzing email and sending to workflow..."
            }
        );

        // Step 1: Get email data and detect processing type from attachments
        loadingText.textContent = 'Collecting email data...';
        const emailData = await getEmailData();
        console.log('Email data collected:', emailData);

        // Detect processing type from attachments
        if (mailboxItem && mailboxItem.attachments && mailboxItem.attachments.length > 0) {
            processingType = detectProcessingTypeFromAttachments(mailboxItem.attachments);
            console.log(`Processing type detected from attachments: ${processingType.toUpperCase()}`);
        } else {
            processingType = 'underwriting'; // default
            console.log('No attachments found, defaulting to: UNDERWRITING');
        }

        // Step 2: Extract form fields from email data immediately
        loadingText.textContent = 'Extracting form data...';
        loadingSubtext.textContent = 'Parsing email content';

        const formFields = extractFormFieldsFromEmail(emailData);

        // Prepare extracted data object
        extractedData = {
            filename: formFields.document_name,
            email_fields: formFields,
            detected_at: emailData.receivedDateTime || new Date().toISOString()
        };

        // Find ACORD attachment (and collect extras) before showing form
        let primaryAttachment = null;
        for (const att of emailData.attachments) {
            if (att.name && att.name.toLowerCase().startsWith('acord_')) {
                primaryAttachment = att;
                break;
            }
        }
        if (!primaryAttachment && emailData.attachments.length > 0) {
            primaryAttachment = emailData.attachments[0];
        }

        // Step 3: Show form IMMEDIATELY (don't wait for backend)
        loadingText.textContent = 'Loading form...';
        loadingSubtext.textContent = 'Preparing your insurance policy information';

        loadingContainer.style.display = 'none';
        formContainer.style.display = 'block';

        await new Promise(resolve => setTimeout(resolve, 100));
        populateForm(extractedData);

        // Step 4: Fire-and-forget backend extraction (runs in background)
        if (primaryAttachment) {
            (async () => {
                try {
                    const apiPrefix = processingType === 'claims' ? '/claims-api' : '/api';
                    const apiBaseURL = processingType === 'claims' ? CLAIMS_API_URL : UNDERWRITING_API_URL;

                    // Build FormData with ACORD PDF + extra attachments + email_metadata
                    const formDataPayload = new FormData();

                    // Decode base64 ACORD PDF and attach as binary
                    const acordB64 = primaryAttachment.contentBytes.includes(',')
                        ? primaryAttachment.contentBytes.split(',')[1]
                        : primaryAttachment.contentBytes;
                    const acordBytes = Uint8Array.from(atob(acordB64), c => c.charCodeAt(0));
                    formDataPayload.append('file', new Blob([acordBytes], { type: 'application/pdf' }), primaryAttachment.name);

                    // Attach extra attachments (loss run DOCX etc.) - skip images
                    const IMAGE_EXTS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp', '.ico'];
                    for (const att of emailData.attachments) {
                        if (att === primaryAttachment) continue;
                        if (!att.contentBytes) continue;
                        const attExt = (att.name || '').toLowerCase().split('.').pop();
                        if (IMAGE_EXTS.includes('.' + attExt)) {
                            console.log('Skipping image attachment:', att.name);
                            continue;
                        }
                        try {
                            const extraB64 = att.contentBytes.includes(',')
                                ? att.contentBytes.split(',')[1]
                                : att.contentBytes;
                            const extraBytes = Uint8Array.from(atob(extraB64), c => c.charCodeAt(0));
                            formDataPayload.append('extra_attachments', new Blob([extraBytes]), att.name);
                        } catch (e) {
                            console.warn('Could not attach extra file:', att.name, e);
                        }
                    }

                    // Attach full email_metadata (including itemId + internetMessageId for EML download)
                    formDataPayload.append('email_metadata', JSON.stringify({
                        subject: emailData.subject,
                        from: emailData.from,
                        receivedDateTime: emailData.receivedDateTime,
                        body: emailData.body,
                        userEmail: emailData.userEmail,
                        triggeredAt: emailData.triggeredAt,
                        id: emailData.itemId,
                        itemId: emailData.itemId,
                        internetMessageId: emailData.internetMessageId,
                        conversationId: emailData.conversationId
                    }));

                    const extractResponse = await fetch(`${apiBaseURL}${apiPrefix}/extract`, {
                        method: 'POST',
                        headers: { 'ngrok-skip-browser-warning': 'true' },
                        body: formDataPayload
                    });

                    if (extractResponse.ok) {
                        const result = await extractResponse.json();
                        console.log('✓ Background extraction complete:', result);
                        extractedData.session_id = result.session_id;
                        extractedData.processing_status = 'extracted';
                    } else {
                        console.warn('Background extraction failed:', extractResponse.status);
                    }
                } catch (apiError) {
                    console.error('Background extraction error:', apiError);
                }
            })();
        }

    } catch (error) {
        console.error('Error:', error);

        // Remove any pending notifications
        Office.context.mailbox.item.notificationMessages.removeAsync("progress");
        Office.context.mailbox.item.notificationMessages.removeAsync("formProcessing");

        // Show error notification in email view
        Office.context.mailbox.item.notificationMessages.addAsync(
            "error",
            {
                type: "errorMessage",
                message: "Saving failed: " + error.message
            }
        );

        loadingText.textContent = 'Error occurred';
        loadingSubtext.innerHTML = `<div class="error-message">${error.message}</div>`;
    }
}


function populateForm(extractedData) {
    // Store filename for later submission
    filename = extractedData.filename || '';
    console.log(`Populating form for ${processingType.toUpperCase()} workflow with file: ${filename}`);

    // Get email_fields from the extracted data
    const data = extractedData.email_fields || extractedData.extracted_data || {};

    console.log("Populating form with data:", data);

    // Populate form fields - only the 5 required fields
    const policyNumberEl = document.getElementById('policyNumber');
    const documentNameEl = document.getElementById('documentName');
    // const subjectNameEl = document.getElementById('subjectName');
    const commentsEl = document.getElementById('comments');

    // Debug logging
    console.log('Form elements found:', {
        policyNumber: !!policyNumberEl,
        documentName: !!documentNameEl,
        // subjectName: !!subjectNameEl,
        comments: !!commentsEl
    });

    if (policyNumberEl && data.policy_number) policyNumberEl.value = data.policy_number;
    // Populate Document Name field with email subject (primary) or attachment name (fallback)
    if (documentNameEl) {
        documentNameEl.value = data.subject;
    }
    // if (subjectNameEl && data.subject) subjectNameEl.value = data.subject;
    if (commentsEl && data.comments) commentsEl.value = data.comments;

    // Commented out old fields
    // if (data.broker_email) document.getElementById('brokerEmail').value = data.broker_email;
    // if (data.broker_name) document.getElementById('brokerName').value = data.broker_name;
    // if (data.underwriter_email) document.getElementById('underwriterEmail').value = data.underwriter_email;
    // if (data.underwriter_name) document.getElementById('underwriterName').value = data.underwriter_name;
    // if (data.broker_agency_name) document.getElementById('agencyName').value = data.broker_agency_name;
    // if (data.broker_agency_id || data.agency_id) document.getElementById('agencyId').value = data.broker_agency_id || data.agency_id;
    // if (data.email_summary) document.getElementById('emailSummary').value = data.email_summary;

    // Set timestamp
    const timestampField = document.getElementById('timestamp');
    if (timestampField) {
        if (data.timestamp) {
            // Use timestamp from backend if available
            timestampField.value = data.timestamp;
        } else if (extractedData.detected_at) {
            // Format detected_at timestamp (remove seconds)
            const detectedDate = new Date(extractedData.detected_at);
            timestampField.value = detectedDate.toLocaleString('en-US', {
                year: 'numeric',
                month: '2-digit',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                hour12: true
            });
        } else {
            const now = new Date();
            timestampField.value = now.toLocaleString('en-US', {
                year: 'numeric',
                month: '2-digit',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                hour12: true
            });
        }
    }
}

// Collapse form permanently after successful submission
function collapseForm() {
    const formContainer = document.getElementById('formContainer');
    const formBody = document.querySelector('.form-body');
    const submitSection = document.querySelector('.submit-section');
    const header = document.querySelector('.header');
    const successMessage = document.getElementById('successMessage');

    console.log('Collapsing form...');

    // Hide form body, submit section, and success message
    if (formBody) {
        formBody.style.display = 'none';
        console.log('Form body hidden');
    }
    if (submitSection) {
        submitSection.style.display = 'none';
        console.log('Submit section hidden');
    }
    if (successMessage) {
        successMessage.classList.remove('show');
        successMessage.style.display = 'none';
        console.log('Success message hidden');
    }

    // Add collapsed state styling
    if (formContainer) {
        formContainer.style.transition = 'all 0.3s ease';
    }

    // Create or update success summary
    let successSummary = document.getElementById('successSummary');
    if (!successSummary) {
        successSummary = document.createElement('div');
        successSummary.id = 'successSummary';
        successSummary.className = 'success-summary';
        successSummary.innerHTML = `
            <div class="success-summary-content">
                <div class="success-icon">✓</div>
                <div class="success-text">
                    <h3>Form Submitted Successfully</h3>
                    <p>Your insurance policy information has been processed and the report has been generated.</p>
                </div>
            </div>
        `;

        // Insert after header
        if (header && header.parentNode) {
            header.parentNode.insertBefore(successSummary, header.nextSibling);
        }
    }

    successSummary.style.display = 'block';
    console.log('Success summary displayed');
}

async function getEmailData() {
    const item = mailboxItem;
    if (!item) throw new Error("No email context");

    const subject = await getSubject(item);
    const body = await getBody(item);
    const attachments = await getAttachmentContents(item);

    const itemId = item.itemId || "";
    const conversationId = item.conversationId || "";

    console.log("Item ID:", itemId);
    console.log("Conversation ID:", conversationId);

    return {
        triggeredAt: new Date().toISOString(),
        userEmail: Office.context.mailbox.userProfile.emailAddress || "",
        subject: subject,
        body: body,
        from: getSender(item),
        receivedDateTime: item.dateTimeCreated || item.dateTimeModified || new Date().toISOString(),
        internetMessageId: item.internetMessageId || "",
        itemId: itemId,
        conversationId: conversationId,
        hasAttachments: attachments.length > 0,
        attachmentCount: attachments.length,
        attachments: attachments
    };
}

function getSubject(item) {
    return new Promise(resolve => {
        if (typeof item.subject === "string") {
            resolve(item.subject);
        } else {
            item.subject.getAsync(result =>
                resolve(result.status === Office.AsyncResultStatus.Succeeded
                    ? result.value
                    : "Subject unavailable")
            );
        }
    });
}

function getBody(item) {
    return new Promise(resolve => {
        item.body.getAsync(Office.CoercionType.Text, result =>
            resolve(result.status === Office.AsyncResultStatus.Succeeded
                ? result.value
                : "Body unavailable")
        );
    });
}

function getSender(item) {
    if (!item.from) return "Unknown";
    if (typeof item.from === "string") return item.from;
    if (item.from.emailAddress) return item.from.emailAddress;
    if (item.from.displayName) return item.from.displayName;
    return "Unknown";
}

async function getAttachmentContents(item) {
    if (!item.attachments || item.attachments.length === 0) {
        return [];
    }

    const results = [];

    for (const att of item.attachments) {
        console.log(`Retrieving content for attachment: ${att.name} (ID: ${att.id})`);

        try {
            await new Promise(resolve => {
                // Explicitly request Base64 format for file attachments
                item.getAttachmentContentAsync(att.id, { format: Office.MailboxEnums.AttachmentContentFormat.Base64 }, res => {
                    if (res.status === Office.AsyncResultStatus.Succeeded) {
                        console.log(`  ✓ Successfully retrieved ${att.name}`);
                        results.push({
                            id: att.id,
                            name: att.name,
                            contentType: att.contentType,
                            size: att.size,
                            content: res.value.content, // legacy field name
                            contentBytes: res.value.content, // standard name for Power Automate
                            format: res.value.format
                        });
                    } else {
                        console.error(`  ✗ Failed to retrieve ${att.name}: ${res.error ? res.error.message : 'Unknown error'}`);
                        results.push({
                            id: att.id,
                            name: att.name,
                            error: res.error ? res.error.message : 'Unknown error'
                        });
                    }
                    resolve();
                });
            });
        } catch (error) {
            console.error(`Exception retrieving ${att.name}:`, error);
            results.push({
                id: att.id,
                name: att.name,
                error: error.message
            });
        }
    }

    return results;
}

// Generate PDF from form data
async function generateFormPDF(formData) {
    return new Promise((resolve) => {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();

        // Add header
        doc.setFillColor(249, 168, 37);
        doc.rect(0, 0, 210, 40, 'F');

        doc.setTextColor(255, 255, 255);
        doc.setFontSize(22);
        doc.setFont(undefined, 'bold');
        doc.text('Insurance Policy Information', 105, 20, { align: 'center' });

        doc.setFontSize(12);
        doc.setFont(undefined, 'normal');
        doc.text('Submitted Form Response', 105, 30, { align: 'center' });

        // Reset text color for body
        doc.setTextColor(0, 0, 0);
        let yPos = 50;

        // Add form fields - only the 5 required fields
        const fields = [
            { label: "Policy Number", value: formData.policy_number },
            { label: "Document Name", value: formData.subject },
            // { label: "Document Name", value: formData.document_name },
            // { label: "Subject Name", value: formData.subject },
            { label: "Comments", value: formData.comments, multiline: false },
            { label: "Timestamp", value: formData.timestamp }
            // Commented out old fields
            // { label: "Sender's Email", value: formData.broker_email },
            // { label: "Sender's Name", value: formData.broker_name },
            // { label: "Receiver's Email", value: formData.underwriter_email },
            // { label: "Receiver's Name", value: formData.underwriter_name },
            // { label: "Agency Name", value: formData.broker_agency_name },
            // { label: "Agency/Broker ID", value: formData.broker_agency_id },
            // { label: "Email Summary", value: formData.email_summary, multiline: true },
        ];

        fields.forEach(field => {
            if (field.value) {
                doc.setFontSize(11);
                doc.setFont(undefined, 'bold');
                doc.text(field.label + ':', 20, yPos);

                doc.setFont(undefined, 'normal');
                doc.setFontSize(10);

                if (field.multiline && field.value.length > 60) {
                    // Handle long text with wrapping
                    const lines = doc.splitTextToSize(field.value, 170);
                    yPos += 6;
                    lines.forEach(line => {
                        if (yPos > 270) {
                            doc.addPage();
                            yPos = 20;
                        }
                        doc.text(line, 20, yPos);
                        yPos += 5;
                    });
                    yPos += 3;
                } else {
                    yPos += 6;
                    doc.text(field.value, 20, yPos);
                    yPos += 8;
                }

                if (yPos > 270) {
                    doc.addPage();
                    yPos = 20;
                }
            }
        });

        // Add footer
        const pageCount = doc.internal.getNumberOfPages();
        for (let i = 1; i <= pageCount; i++) {
            doc.setPage(i);
            doc.setFontSize(9);
            doc.setTextColor(128, 128, 128);
            doc.text(`Page ${i} of ${pageCount}`, 105, 285, { align: 'center' });

        }

        // Convert to base64
        const pdfBase64 = doc.output('dataurlstring').split(',')[1];
        resolve(pdfBase64);
    });
}

// Handle form submission
document.addEventListener('DOMContentLoaded', function () {
    const insuranceForm = document.getElementById('insuranceForm');
    if (insuranceForm) {
        insuranceForm.addEventListener('submit', handleFormSubmit);
    }
});

async function handleFormSubmit(e) {
    e.preventDefault();

    const formData = {
        policy_number: document.getElementById('policyNumber').value,
        document_name: document.getElementById('documentName').value,
        // subject: document.getElementById('subjectName').value,
        comments: document.getElementById('comments').value,
        timestamp: document.getElementById('timestamp').value
        // Commented out old fields
        // broker_email: document.getElementById('brokerEmail').value,
        // broker_name: document.getElementById('brokerName').value,
        // underwriter_email: document.getElementById('underwriterEmail').value,
        // underwriter_name: document.getElementById('underwriterName').value,
        // broker_agency_name: document.getElementById('agencyName').value,
        // broker_agency_id: document.getElementById('agencyId').value,
        // email_summary: document.getElementById('emailSummary').value,
    };

    const submitButton = document.querySelector('.submit-button');
    const successMessage = document.getElementById('successMessage');

    try {
        submitButton.disabled = true;
        submitButton.textContent = 'Generating PDF...';

        // Generate PDF from form data
        const pdfBase64 = await generateFormPDF(formData);
        console.log('PDF generated successfully');

        submitButton.textContent = 'Submitting...';

        // Combined request: Send email data + attachment + form fields directly to backend
        console.log('Sending direct submission to backend...');
        console.log('Filename:', filename);
        console.log('Form data:', formData);
        const apiPrefix = processingType === 'claims' ? '/claims-api' : '/api';
        const apiBaseURL = processingType === 'claims' ? CLAIMS_API_URL : UNDERWRITING_API_URL;

        // Get email data with attachment
        const emailDataWithAttachment = await getEmailData();

        // Find the appropriate attachment (ACORD file or first attachment)
        let primaryAttachment = null;
        for (const att of emailDataWithAttachment.attachments) {
            if (att.name && att.name.toLowerCase().startsWith('acord_')) {
                primaryAttachment = att;
                break;
            }
        }

        // Fallback to first attachment if no ACORD file found
        if (!primaryAttachment && emailDataWithAttachment.attachments.length > 0) {
            primaryAttachment = emailDataWithAttachment.attachments[0];
        }

        if (!primaryAttachment) {
            throw new Error('No attachment found in email');
        }

        // Use /api/process if we have a session_id from the trigger step, else fallback to /api/submit
        const sessionId = extractedData && extractedData.session_id;
        let submitResponse;

        if (sessionId && processingType !== 'claims') {
            // Part 2: process using existing session (no re-extraction)
            console.log('Using /api/process with session_id:', sessionId);
            submitResponse = await fetch(`${apiBaseURL}${apiPrefix}/process`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'ngrok-skip-browser-warning': 'true'
                },
                body: JSON.stringify({
                    session_id: sessionId,
                    email_fields: formData,
                    form_pdf: pdfBase64
                })
            });
        } else {
            // Fallback: full submit (session_id not ready or claims flow)
            console.log('Falling back to /api/submit (no session_id or claims flow)');
            submitResponse = await fetch(`${apiBaseURL}${apiPrefix}/submit`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'ngrok-skip-browser-warning': 'true'
                },
                body: JSON.stringify({
                    filename: primaryAttachment.name,
                    attachment_base64: primaryAttachment.contentBytes,
                    email_metadata: {
                        subject: emailDataWithAttachment.subject,
                        from: emailDataWithAttachment.from,
                        receivedDateTime: emailDataWithAttachment.receivedDateTime,
                        body: emailDataWithAttachment.body,
                        userEmail: emailDataWithAttachment.userEmail,
                        triggeredAt: emailDataWithAttachment.triggeredAt,
                        id: emailDataWithAttachment.itemId,
                        itemId: emailDataWithAttachment.itemId,
                        internetMessageId: emailDataWithAttachment.internetMessageId,
                        conversationId: emailDataWithAttachment.conversationId
                    },
                    email_fields: formData,
                    extra_attachments: emailDataWithAttachment.attachments
                        .filter(a => a !== primaryAttachment && a.contentBytes)
                        .map(a => ({ name: a.name, contentBytes: a.contentBytes }))
                })
            });
        }

        // Handle async 202 (background processing) or sync 200 (skipped/done)
        if (submitResponse.status === 202 || submitResponse.ok) {
            const result = await submitResponse.json();
            console.log('Submit response:', result);

            if (result.status === 'skipped') {
                console.log('File was skipped - not an ACORD form:', result.filename);
                successMessage.textContent = 'Email analysis complete';
                successMessage.classList.add('show');
                submitButton.textContent = 'Complete';
                submitButton.disabled = true;
                setTimeout(() => {
                    const underwritingUrl = `${UNDERWRITING_API_URL}/policy-center`;
                    successMessage.innerHTML = `<strong>Email saved to Policy Center successfully</strong><br><br>` +
                        `<a href="${underwritingUrl}" target="_blank">Open Policy Center</a>`;
                }, 1000);
                return;
            }

            // 202: processing started in background — show immediate feedback
            if (submitResponse.status === 202 || result.status === 'processing') {
                successMessage.innerHTML = `Sending email data to the Policy Center&hellip;`;
                successMessage.classList.add('show');
                submitButton.textContent = 'Processing…';
                submitButton.disabled = true;

                // Show Outlook notification immediately
                Office.context.mailbox.item.notificationMessages.removeAsync("progress");
                Office.context.mailbox.item.notificationMessages.addAsync("processComplete", {
                    type: "informationalMessage",
                    message: "This email's data has been saved in Policy Center.",
                    icon: "Icon.80x80",
                    persistent: true
                });

                // Use SSE (Server-Sent Events) — server pushes result instantly, no polling
                const sseSessionId = result.session_id || (extractedData && extractedData.session_id);
                if (sseSessionId) {
                    const evtSource = new EventSource(
                        `${UNDERWRITING_API_URL}/api/stream/${sseSessionId}`
                    );

                    evtSource.onmessage = (event) => {
                        evtSource.close(); // single-event stream, close immediately
                        try {
                            const statusData = JSON.parse(event.data);
                            console.log('SSE result:', statusData);

                            if (statusData.status === 'done') {
                                submitButton.textContent = 'Complete';
                                const underwritingUrl = `${UNDERWRITING_API_URL}/policy-detail/${formData.policy_number}`;
                                let successHtml = `<strong>Email saved to Policy Center successfully</strong><br><br>`;
                                if (processingType === 'claims') {
                                    successHtml += `<a href="${CLAIMS_API_URL}/claims" target="_blank" style="color:#0078d4;text-decoration:none;font-weight:600;display:block;padding:8px 12px;background:#f3f9fc;border-radius:4px;border-left:3px solid #0078d4;">Open Claims Center</a>`;
                                } else {
                                    successHtml += `<a href="${underwritingUrl}" target="_blank" style="color:#0078d4;text-decoration:none;font-weight:600;display:block;padding:8px 12px;background:#f3f9fc;border-radius:4px;border-left:3px solid #0078d4;">Open Policy Center</a>`;
                                }
                                successMessage.innerHTML = successHtml;

                            } else if (statusData.status === 'error') {
                                submitButton.textContent = 'Submit';
                                submitButton.disabled = false;
                                successMessage.innerHTML = `<strong>Processing failed:</strong> ${statusData.error || 'Unknown error'}`;
                                successMessage.style.background = '#fde7e9';
                                successMessage.style.color = '#a80000';
                                successMessage.style.borderLeft = '4px solid #a80000';
                            }
                        } catch (parseErr) {
                            console.warn('SSE parse error:', parseErr);
                        }
                    };

                    evtSource.onerror = (err) => {
                        console.warn('SSE connection error:', err);
                        evtSource.close();

                        // Fallback to polling if SSE fails
                        console.log('Falling back to polling...');
                        const pollInterval = setInterval(async () => {
                            try {
                                const statusResp = await fetch(`${UNDERWRITING_API_URL}/api/status/${sseSessionId}`, {
                                    headers: { 'ngrok-skip-browser-warning': 'true' }
                                });
                                const statusData = await statusResp.json();
                                console.log('Poll status (fallback):', statusData.status);

                                if (statusData.status === 'done') {
                                    clearInterval(pollInterval);
                                    submitButton.textContent = 'Complete';
                                    const underwritingUrl = `${UNDERWRITING_API_URL}/policy-detail/${formData.policy_number}`;
                                    let successHtml = `<strong>Email saved to Policy Center successfully</strong><br><br>`;
                                    if (processingType === 'claims') {
                                        successHtml += `<a href="${CLAIMS_API_URL}/claims" target="_blank" style="color:#0078d4;text-decoration:none;font-weight:600;display:block;padding:8px 12px;background:#f3f9fc;border-radius:4px;border-left:3px solid #0078d4;">Open Claims Center</a>`;
                                    } else {
                                        successHtml += `<a href="${underwritingUrl}" target="_blank" style="color:#0078d4;text-decoration:none;font-weight:600;display:block;padding:8px 12px;background:#f3f9fc;border-radius:4px;border-left:3px solid #0078d4;">Open Policy Center</a>`;
                                    }
                                    successMessage.innerHTML = successHtml;

                                } else if (statusData.status === 'error') {
                                    clearInterval(pollInterval);
                                    submitButton.textContent = 'Submit';
                                    submitButton.disabled = false;
                                    successMessage.innerHTML = `<strong>Processing failed:</strong> ${statusData.error || 'Unknown error'}`;
                                    successMessage.style.background = '#fde7e9';
                                    successMessage.style.color = '#a80000';
                                    successMessage.style.borderLeft = '4px solid #a80000';
                                }
                            } catch (pollErr) {
                                console.warn('Status poll error (fallback):', pollErr);
                            }
                        }, 2000);
                    };
                }

                return;
            }

            // Sync 200 success (shouldn't normally happen for /api/process now, but handle gracefully)
            successMessage.textContent = 'Processing completed successfully!';
            successMessage.classList.add('show');
            submitButton.textContent = 'Complete';
            submitButton.disabled = true;

            const underwritingUrl = `${UNDERWRITING_API_URL}/policy-detail/${formData.policy_number}`;
            const claimsUrl = `${CLAIMS_API_URL}/claims`;
            let successHtml = `<strong>Email saved to Policy Center successfully</strong><br><br>`;
            if (processingType === 'claims') {
                successHtml += `<a href="${claimsUrl}" target="_blank" style="color: #0078d4; text-decoration: none; font-weight: 600; display: block; padding: 8px 12px; background: #f3f9fc; border-radius: 4px; border-left: 3px solid #0078d4;">Open Claims Center</a>`;
            } else if (processingType === 'underwriting') {
                successHtml += `<a href="${underwritingUrl}" target="_blank" style="color: #0078d4; text-decoration: none; font-weight: 600; display: block; padding: 8px 12px; background: #f3f9fc; border-radius: 4px; border-left: 3px solid #0078d4;">Open Policy Center</a>`;
            }
            successMessage.innerHTML = successHtml;
            successMessage.classList.add('show');

            Office.context.mailbox.item.notificationMessages.removeAsync("progress");
            Office.context.mailbox.item.notificationMessages.removeAsync("formSuccess");
            Office.context.mailbox.item.notificationMessages.removeAsync("processSuccess");
            Office.context.mailbox.item.notificationMessages.addAsync("processComplete", {
                type: "informationalMessage",
                message: "This email's data has been saved in Policy Center.",
                icon: "Icon.80x80",
                persistent: true
            });
        } else {
            const errorText = await submitResponse.text();
            console.error('API Error Response:', errorText);
            throw new Error(`Submission failed: ${submitResponse.status}`);
        }


    } catch (error) {
        console.error('Error submitting form:', error);

        successMessage.textContent = 'Submission failed: ' + error.message;
        successMessage.style.background = '#fde7e9';
        successMessage.style.color = '#a80000';
        successMessage.style.borderLeft = '4px solid #a80000';
        successMessage.classList.add('show');

        submitButton.disabled = false;
        submitButton.textContent = 'Submit';

        setTimeout(() => {
            successMessage.classList.remove('show');
            successMessage.style.background = '#e8f4f8';
            successMessage.style.color = '#005a9e';
            successMessage.style.borderLeft = '4px solid #0078d4';
        }, 5000);
    }
}
