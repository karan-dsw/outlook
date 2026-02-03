let mailboxItem = null;
let filename = '';

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        mailboxItem = Office.context.mailbox.item;
        console.log('Office.js initialized successfully');
        
        // Automatically trigger the flow when taskpane opens
        triggerFlowAndLoadForm();
    }
});

async function triggerFlowAndLoadForm() {
    const loadingContainer = document.getElementById('loadingContainer');
    const formContainer = document.getElementById('formContainer');
    const loadingText = document.querySelector('.loading-text');
    const loadingSubtext = document.querySelector('.loading-subtext');
    
    try {
        // Step 1: Get email data
        loadingText.textContent = 'Collecting email data...';
        const emailData = await getEmailData();
        console.log('Email data collected:', emailData);
        
        // Step 2: Trigger Power Automate flow
        loadingText.textContent = 'Triggering Power Automate flow...';
        loadingSubtext.textContent = 'Sending email data for processing';
        
        const flowUrl = "https://default74afe875305e4ab4ba4ac1359a7629.ae.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/89c12382226642a4907cd110e9e7ab87/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Nbz7sUIbNoHlSBt_KVnF3CFKCCf9lPYn-LbIxZsWouA";
        
        const response = await fetch(flowUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(emailData)
        });
        
        if (!response.ok) {
            const err = await response.text();
            throw new Error(`HTTP ${response.status}: ${err}`);
        }
        
        console.log('Flow triggered successfully');
        
        // Step 3: Poll for extracted form data
        loadingText.textContent = 'Extracting form data from attachments...';
        loadingSubtext.textContent = 'This may take a few moments';
        
        let extractedData = null;
        let pollingAttempts = 0;
        const maxPollingAttempts = 60; // 5 minutes max
        
        while (pollingAttempts < maxPollingAttempts) {
            try {
                const pendingResponse = await fetch("https://e7bcca93d9b7.ngrok-free.app/api/pending", {
                    headers: {
                        'ngrok-skip-browser-warning': 'true',
                        'Accept': 'application/json'
                    }
                });
                
                if (pendingResponse.ok) {
                    const data = await pendingResponse.json();
                    console.log("Pending API response:", data);
                    
                    if (data.success && data.count > 0 && data.files && data.files.length > 0) {
                        extractedData = data.files[0];
                        console.log("Extracted data received:", extractedData);
                        break;
                    }
                }
            } catch (pollError) {
                console.warn("Polling attempt failed:", pollError.message);
            }
            
            // Wait 5 seconds before next attempt
            await new Promise(resolve => setTimeout(resolve, 5000));
            pollingAttempts++;
            
            if (pollingAttempts % 6 === 0) {
                loadingSubtext.textContent = `Still processing... (${Math.floor(pollingAttempts / 12)} minute${Math.floor(pollingAttempts / 12) > 1 ? 's' : ''})`;
            }
        }
        
        if (!extractedData) {
            throw new Error('Timeout: Form data extraction did not complete in time');
        }
        
        // Step 4: Populate form with extracted data
        loadingText.textContent = 'Loading form...';
        loadingSubtext.textContent = 'Preparing your insurance policy information';
        
        populateForm(extractedData);
        
        // Show form, hide loading
        loadingContainer.style.display = 'none';
        formContainer.style.display = 'block';
        
    } catch (error) {
        console.error('Error:', error);
        loadingText.textContent = 'Error occurred';
        loadingSubtext.innerHTML = `<div class="error-message">${error.message}</div>`;
    }
}

function populateForm(extractedData) {
    // Store filename for later submission
    filename = extractedData.filename || '';
    
    // Get email_fields from the extracted data
    const data = extractedData.email_fields || extractedData.extracted_data || {};
    
    console.log("Populating form with data:", data);
    
    // Populate form fields
    if (data.broker_email) document.getElementById('brokerEmail').value = data.broker_email;
    if (data.broker_name) document.getElementById('brokerName').value = data.broker_name;
    if (data.underwriter_email) document.getElementById('underwriterEmail').value = data.underwriter_email;
    if (data.underwriter_name) document.getElementById('underwriterName').value = data.underwriter_name;
    if (data.policy_number) document.getElementById('policyNumber').value = data.policy_number;
    if (data.broker_agency_name) document.getElementById('agencyName').value = data.broker_agency_name;
    if (data.broker_agency_id || data.agency_id) document.getElementById('agencyId').value = data.broker_agency_id || data.agency_id;
    
    // Set timestamp
    const timestampField = document.getElementById('timestamp');
    if (extractedData.detected_at) {
        timestampField.value = extractedData.detected_at;
    } else {
        const now = new Date();
        timestampField.value = now.toLocaleString('en-US', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
            hour12: true
        });
    }
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
        console.log(`Processing attachment: ${att.name}`);
        
        try {
            await new Promise(resolve => {
                item.getAttachmentContentAsync(att.id, res => {
                    if (res.status === Office.AsyncResultStatus.Succeeded) {
                        results.push({
                            id: att.id,
                            name: att.name,
                            contentType: att.contentType,
                            size: att.size,
                            content: res.value.content,
                            format: res.value.format
                        });
                    } else {
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

// Handle form submission
document.addEventListener('DOMContentLoaded', function() {
    const insuranceForm = document.getElementById('insuranceForm');
    if (insuranceForm) {
        insuranceForm.addEventListener('submit', handleFormSubmit);
    }
});

async function handleFormSubmit(e) {
    e.preventDefault();
    
    const formData = {
        broker_email: document.getElementById('brokerEmail').value,
        broker_name: document.getElementById('brokerName').value,
        underwriter_email: document.getElementById('underwriterEmail').value,
        underwriter_name: document.getElementById('underwriterName').value,
        policy_number: document.getElementById('policyNumber').value,
        broker_agency_name: document.getElementById('agencyName').value,
        broker_agency_id: document.getElementById('agencyId').value,
        timestamp: document.getElementById('timestamp').value
    };
    
    const submitButton = document.querySelector('.submit-button');
    const successMessage = document.getElementById('successMessage');
    
    try {
        submitButton.disabled = true;
        submitButton.textContent = 'Submitting...';
        
        // Step 3: Confirm email fields
        console.log('Confirming email fields...');
        const confirmResponse = await fetch('https://e7bcca93d9b7.ngrok-free.app/api/email-fields', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'ngrok-skip-browser-warning': 'true'
            },
            body: JSON.stringify({
                filename: filename,
                email_fields: formData
            })
        });
        
        if (!confirmResponse.ok) {
            throw new Error(`Email fields confirmation failed: ${confirmResponse.status}`);
        }
        
        console.log('Email fields confirmed');
        
        // Step 4: Process the file
        console.log('Processing file...');
        submitButton.textContent = 'Processing...';
        
        const processResponse = await fetch('https://e7bcca93d9b7.ngrok-free.app/api/process', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'ngrok-skip-browser-warning': 'true'
            },
            body: JSON.stringify({
                filename: filename
            })
        });
        
        if (!processResponse.ok) {
            throw new Error(`Processing failed: ${processResponse.status}`);
        }
        
        const result = await processResponse.json();
        console.log('Processing successful:', result);
        
        successMessage.textContent = '✓ Form submitted successfully! Generating report...';
        successMessage.classList.add('show');
        submitButton.textContent = 'Generating Report...';
        
        // Step 5: Poll for PDF report
        console.log('Waiting for PDF report...');
        let pdfReady = false;
        let pdfPollingAttempts = 0;
        const maxPdfPollingAttempts = 60;
        
        while (pdfPollingAttempts < maxPdfPollingAttempts && !pdfReady) {
            try {
                const pdfResponse = await fetch('https://e7bcca93d9b7.ngrok-free.app/api/output-pdf', {
                    headers: {
                        'ngrok-skip-browser-warning': 'true',
                        'Accept': 'application/json'
                    }
                });
                
                if (pdfResponse.status === 200) {
                    const pdfData = await pdfResponse.json();
                    console.log('PDF is ready!', pdfData);
                    
                    const reportUrl = pdfData.pdf_url;
                    
                    if (reportUrl) {
                        console.log('Opening report:', reportUrl);
                        
                        // Use Office Dialog API to open the report
                        Office.context.ui.displayDialogAsync(
                            reportUrl,
                            { height: 90, width: 70, displayInIframe: false },
                            (result) => {
                                if (result.status === Office.AsyncResultStatus.Succeeded) {
                                    console.log('Report opened successfully');
                                    successMessage.textContent = '✓ Form submitted and report opened successfully!';
                                } else {
                                    console.error('Failed to open report:', result.error.message);
                                    // Provide clickable link as fallback
                                    successMessage.innerHTML = `✓ Form submitted! <a href="${reportUrl}" target="_blank" style="color: #0078d4; text-decoration: underline; font-weight: bold;">Click here to open report</a>`;
                                }
                            }
                        );
                    } else {
                        successMessage.textContent = '✓ Form submitted! Report URL not available.';
                    }
                    
                    pdfReady = true;
                    break;
                }
            } catch (pollError) {
                console.warn('PDF polling attempt failed:', pollError.message);
            }
            
            await new Promise(resolve => setTimeout(resolve, 5000));
            pdfPollingAttempts++;
        }
        
        if (!pdfReady) {
            successMessage.textContent = '✓ Form submitted! Report is still processing...';
        }
        
        submitButton.textContent = 'Submitted';
        
        setTimeout(() => {
            successMessage.classList.remove('show');
            submitButton.disabled = false;
            submitButton.textContent = 'Submit';
        }, 3000);
        
    } catch (error) {
        console.error('Error submitting form:', error);
        
        successMessage.textContent = '✗ Submission failed: ' + error.message;
        successMessage.style.background = '#fde7e9';
        successMessage.style.color = '#a80000';
        successMessage.style.borderLeft = '4px solid #a80000';
        successMessage.classList.add('show');
        
        submitButton.disabled = false;
        submitButton.textContent = 'Submit';
        
        setTimeout(() => {
            successMessage.classList.remove('show');
            successMessage.style.background = '#dff6dd';
            successMessage.style.color = '#0b7815';
            successMessage.style.borderLeft = '4px solid #0b7815';
        }, 5000);
    }
}
