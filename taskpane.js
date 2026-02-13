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
        formFields.document_name = emailData.attachments[0].name || '';
        filename = emailData.attachments[0].name || '';
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

        // Step 3: Trigger Power Automate flow (async, don't wait)
        loadingText.textContent = 'Triggering Power Automate flow...';
        loadingSubtext.textContent = 'Sending email data for processing';

        const flowUrl = "https://default74afe875305e4ab4ba4ac1359a7629.ae.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/89c12382226642a4907cd110e9e7ab87/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Nbz7sUIbNoHlSBt_KVnF3CFKCCf9lPYn-LbIxZsWouA";

        // Trigger flow but don't wait for response (fire and forget)
        fetch(flowUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(emailData)
        }).then(response => {
            if (response.ok) {
                console.log('Flow triggered successfully');
            } else {
                console.warn('Flow trigger failed:', response.status);
            }
        }).catch(err => {
            console.warn('Flow trigger error:', err);
        });

        Office.context.mailbox.item.notificationMessages.removeAsync("progress");
        Office.context.mailbox.item.notificationMessages.removeAsync("formSuccess"); // Clear any previous notifications

        // Show success notification
        Office.context.mailbox.item.notificationMessages.addAsync(
            "formSuccess",
            {
                type: "informationalMessage",
                message: "Form data extracted! Opening form in taskpane...",
                icon: "Icon.80x80",
                persistent: false
            }
        );

        // Step 4: Show form and populate immediately
        loadingText.textContent = 'Loading form...';
        loadingSubtext.textContent = 'Preparing your insurance policy information';

        // Show form before populating (elements must be visible to access)
        loadingContainer.style.display = 'none';
        formContainer.style.display = 'block';

        // Small delay to ensure DOM is fully rendered
        await new Promise(resolve => setTimeout(resolve, 100));

        // Now populate the form
        populateForm(extractedData);

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
                message: "Processing failed: " + error.message
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

        // Combined request: Process file with email fields and PDF
        console.log('Processing file with email fields...');
        const apiPrefix = processingType === 'claims' ? '/claims-api' : '/api';
        const apiBaseURL = processingType === 'claims' ? CLAIMS_API_URL : UNDERWRITING_API_URL;

        const processResponse = await fetch(`${apiBaseURL}${apiPrefix}/process`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'ngrok-skip-browser-warning': 'true'
            },
            body: JSON.stringify({
                filename: filename,
                email_fields: formData,
                form_pdf: pdfBase64  // Include the PDF in base64 format
            })
        });

        if (!processResponse.ok) {
            throw new Error(`Processing failed: ${processResponse.status}`);
        }

        const result = await processResponse.json();
        console.log('Processing successful:', result);

        successMessage.textContent = 'Form submitted successfully! Generating report...';
        successMessage.classList.add('show');
        submitButton.textContent = 'Generating Report...';

        // Step 5: Poll for PDF report - start immediately with shorter interval
        console.log('Waiting for PDF report...');
        let pdfReady = false;
        let pdfPollingAttempts = 0;
        const maxPdfPollingAttempts = 60; // 2 minutes max (60 attempts * 2 seconds)

        while (pdfPollingAttempts < maxPdfPollingAttempts && !pdfReady) {
            try {
                const pdfResponse = await fetch(`${apiBaseURL}${apiPrefix}/output-pdf`, {
                    headers: {
                        'ngrok-skip-browser-warning': 'true',
                        'Accept': 'application/json'
                    }
                });

                if (pdfResponse.status === 200) {
                    const pdfData = await pdfResponse.json();
                    console.log('PDF is ready!', pdfData);

                    // Prepare URLs
                    const reportUrl = pdfData.pdf_url;
                    const sessionID = extractedData ? (extractedData.session_id || extractedData._session_id) : '';
                    // const underwritingUrl = `${UNDERWRITING_API_URL}/policy-detail/${formData.policy_number}`;
                    const underwritingUrl = `${UNDERWRITING_API_URL}/policy-center`;
                    const claimsUrl = `${CLAIMS_API_URL}/claims`;



                    // Update success message with professional styling (no emojis)
                    let successHtml = `<strong>Email processed successfully!</strong><br><br>`;
                    successHtml += `<a href="${reportUrl}" target="_blank" style="color: #0078d4; text-decoration: none; font-weight: 600; display: block; margin-bottom: 10px; padding: 8px 12px; background: #f3f9fc; border-radius: 4px; border-left: 3px solid #0078d4;">Open Processed Report</a>`;

                    if (processingType === 'claims') {
                        successHtml += `<a href="${claimsUrl}" target="_blank" style="color: #0078d4; text-decoration: none; font-weight: 600; display: block; padding: 8px 12px; background: #f3f9fc; border-radius: 4px; border-left: 3px solid #0078d4;">Open Claims Management</a>`;

                        // Also try automatic open for claims
                        // setTimeout(() => {
                        //     try { window.open(claimsUrl, '_blank', 'noopener,noreferrer'); } catch (e) { }
                        // }, 2000);
                    } else if (processingType === 'underwriting') {
                        successHtml += `<a href="${underwritingUrl}" target="_blank" style="color: #0078d4; text-decoration: none; font-weight: 600; display: block; padding: 8px 12px; background: #f3f9fc; border-radius: 4px; border-left: 3px solid #0078d4;">Open Policy Center</a>`;

                        // Also try automatic open for underwriting
                        // setTimeout(() => {
                        //     try { window.open(underwritingUrl, '_blank', 'noopener,noreferrer'); } catch (e) { }
                        // }, 2000);
                    }

                    // console.log('Opening report:', reportUrl);

                    // // Simulate Ctrl+Click to open tabs in background (without switching)
                    // function openInBackgroundTab(url) {
                    //     const a = document.createElement('a');
                    //     a.href = url;
                    //     a.target = '_blank';
                    //     a.rel = 'noopener noreferrer';
                    //     a.style.display = 'none';
                    //     document.body.appendChild(a);

                    //     // Simulate Ctrl+Click (Cmd+Click on Mac)
                    //     const evt = new MouseEvent('click', {
                    //         bubbles: true,
                    //         cancelable: true,
                    //         view: window,
                    //         ctrlKey: true,  // This is the key to opening in background!
                    //         metaKey: true   // For Mac users
                    //     });

                    //     a.dispatchEvent(evt);
                    //     document.body.removeChild(a);
                    // }

                    // try {
                    //     // Open report in background tab
                    //     openInBackgroundTab(reportUrl);

                    //     // Small delay to ensure tabs open in order
                    //     setTimeout(() => {
                    //         // Open the appropriate dashboard based on processing type
                    //         if (processingType === 'claims') {
                    //             openInBackgroundTab(claimsUrl);
                    //         } else if (processingType === 'underwriting') {
                    //             openInBackgroundTab(underwritingUrl);
                    //         }
                    //     }, 100);
                    // } catch (e) {
                    //     console.error('Error opening links:', e);
                    // }


                    successMessage.innerHTML = successHtml;
                    successMessage.classList.add('show');
                    pdfReady = true;
                    break;
                } else {
                    console.log('Waiting for PDF report...');
                }
            } catch (pollError) {
                console.warn('PDF polling attempt failed:', pollError.message);
            }

            await new Promise(resolve => setTimeout(resolve, 2000));
            pdfPollingAttempts++;
        }

        if (!pdfReady) {
            successMessage.textContent = 'Email processed successfully! Report is still processing...';
        }

        // Clear all previous notifications first to prevent duplicates
        Office.context.mailbox.item.notificationMessages.removeAsync("progress");
        Office.context.mailbox.item.notificationMessages.removeAsync("formSuccess");
        Office.context.mailbox.item.notificationMessages.removeAsync("processSuccess");

        // Add the final success notification (only this one should show)
        Office.context.mailbox.item.notificationMessages.addAsync(
            "processComplete",  // Use a unique ID for the final notification
            {
                type: "informationalMessage",
                message: "Email processed successfully!",
                icon: "Icon.80x80",
                persistent: true
            }
        );

        // Keep button as "Processed" and disabled permanently
        submitButton.textContent = 'Processed';
        submitButton.disabled = true;

        // Taskpane will remain open so user can see the success message and links
        // setTimeout(() => {
        //     console.log('Closing taskpane...');
        //     Office.context.ui.closeContainer();
        // }, 2000);

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
