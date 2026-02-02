let mailboxItem = null;
let flowTriggered = false;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Store the mailbox item reference
        mailboxItem = Office.context.mailbox.item;
        
        // Display current email info
        displayEmailInfo();
        
        // Set up form submission
        const form = document.getElementById('underwritingForm');
        if (form) {
            form.addEventListener('submit', handleFormSubmit);
        }
        
        console.log('Office.js initialized successfully - Form ready');
        
        // Automatically trigger the flow when taskpane opens
        if (!flowTriggered) {
            triggerFlowOnOpen();
        }
    }
});

async function triggerFlowOnOpen() {
    flowTriggered = true;
    
    try {
        showStatus('Triggering Power Automate flow...', 'info');
        
        // Get email data including body for analysis
        const emailData = await getEmailDataForFlow();
        
        console.log('Email data collected:', emailData);
        
        // Your Power Automate flow URL
        const flowUrl = 'https://default74afe875305e4ab4ba4ac1359a7629.ae.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/98694a6fe5ce4b1d8389f23d378bd9e0/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jEGsyOkQ2_bKCUTqufTMr99gQw9x4Q5oPphpSI7fMEA';
        
        // Call Power Automate
        const response = await fetch(flowUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(emailData)
        });
        
        if (response.ok) {
            console.log('Flow triggered successfully');
            showStatus('Flow triggered successfully! Please fill out the form below.', 'success');
        } else {
            const errorText = await response.text();
            throw new Error(`HTTP ${response.status}: ${errorText}`);
        }
        
    } catch (error) {
        console.error('Error triggering flow:', error);
        showStatus('Warning: Flow trigger failed. You can still fill out the form.', 'error');
    }
}

async function getEmailDataForFlow() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        
        if (!item) {
            reject(new Error('No email item available'));
            return;
        }
        
        // Initialize data object
        const emailData = {
            itemId: item.itemId || '',
            conversationId: item.conversationId || '',
            triggeredAt: new Date().toISOString(),
            userEmail: Office.context.mailbox.userProfile.emailAddress || '',
            subject: '',
            from: '',
            body: '',
            hasAttachments: false,
            attachmentCount: 0
        };
        
        // Get subject
        const getSubject = () => {
            return new Promise((resolveSubject) => {
                if (typeof item.subject === 'string') {
                    resolveSubject(item.subject);
                } else if (item.subject && typeof item.subject.getAsync === 'function') {
                    item.subject.getAsync((result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            resolveSubject(result.value);
                        } else {
                            resolveSubject('Unable to retrieve subject');
                        }
                    });
                } else {
                    resolveSubject('Subject not available');
                }
            });
        };
        
        // Get email body for analysis
        const getBody = () => {
            return new Promise((resolveBody) => {
                if (item.body && typeof item.body.getAsync === 'function') {
                    item.body.getAsync(Office.CoercionType.Text, (result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            resolveBody(result.value);
                        } else {
                            resolveBody('Unable to retrieve body');
                        }
                    });
                } else {
                    resolveBody('Body not available');
                }
            });
        };
        
        // Get sender
        const getFrom = () => {
            if (item.from) {
                if (typeof item.from === 'string') {
                    return item.from;
                } else if (item.from.emailAddress) {
                    return item.from.emailAddress;
                } else if (item.from.displayName) {
                    return item.from.displayName;
                }
            }
            return 'Unknown sender';
        };
        
        // Get attachment info
        const getAttachments = () => {
            if (item.attachments && item.attachments.length > 0) {
                return {
                    hasAttachments: true,
                    count: item.attachments.length,
                    attachments: item.attachments.map(att => ({
                        name: att.name,
                        contentType: att.contentType,
                        size: att.size,
                        id: att.id
                    }))
                };
            }
            return { hasAttachments: false, count: 0, attachments: [] };
        };
        
        // Collect all data
        Promise.all([getSubject(), getBody()]).then(([subject, body]) => {
            emailData.subject = subject;
            emailData.body = body;
            emailData.from = getFrom();
            
            const attachmentInfo = getAttachments();
            emailData.hasAttachments = attachmentInfo.hasAttachments;
            emailData.attachmentCount = attachmentInfo.count;
            emailData.attachments = attachmentInfo.attachments;
            
            resolve(emailData);
        }).catch(error => {
            reject(error);
        });
    });
}

function displayEmailInfo() {
    if (!mailboxItem) {
        console.error('Mailbox item not available');
        return;
    }
    
    try {
        // Get subject
        const subject = mailboxItem.subject || 'No subject';
        const subjectElement = document.getElementById('emailSubject');
        if (subjectElement) {
            subjectElement.textContent = subject;
        }
        
        // Get sender
        const from = mailboxItem.from 
            ? (mailboxItem.from.emailAddress || mailboxItem.from.displayName || 'Unknown')
            : 'Unknown';
        const fromElement = document.getElementById('emailFrom');
        if (fromElement) {
            fromElement.textContent = from;
        }
        
    } catch (error) {
        console.error('Error displaying email info:', error);
    }
}

async function handleFormSubmit(event) {
    event.preventDefault();
    
    const submitButton = document.getElementById('submitButton');
    const statusDiv = document.getElementById('statusMessage');
    
    // Get form data
    const formData = {
        policyNumber: document.getElementById('policyNumber').value,
        claimType: document.getElementById('claimType').value,
        claimAmount: document.getElementById('claimAmount').value || null,
        priority: document.getElementById('priority').value,
        notes: document.getElementById('notes').value,
        // Include email context
        emailSubject: mailboxItem.subject || 'No subject',
        emailFrom: mailboxItem.from ? mailboxItem.from.emailAddress : 'Unknown',
        itemId: mailboxItem.itemId || 'Unknown',
        conversationId: mailboxItem.conversationId || 'Unknown',
        submittedAt: new Date().toISOString(),
        submittedBy: Office.context.mailbox.userProfile.emailAddress
    };
    
    console.log('Form data collected:', formData);
    
    // Disable submit button
    submitButton.disabled = true;
    submitButton.textContent = 'Submitting...';
    
    showStatus('Saving form data...', 'info');
    
    try {
        // Store in localStorage
        const storageKey = `underwriting_${Date.now()}_${formData.itemId}`;
        localStorage.setItem(storageKey, JSON.stringify(formData));
        
        // Also maintain a list of all submissions
        let allSubmissions = JSON.parse(localStorage.getItem('underwriting_submissions') || '[]');
        allSubmissions.push({
            key: storageKey,
            timestamp: formData.submittedAt,
            policyNumber: formData.policyNumber,
            claimType: formData.claimType,
            priority: formData.priority
        });
        localStorage.setItem('underwriting_submissions', JSON.stringify(allSubmissions));
        
        console.log('✅ Form data saved successfully to localStorage');
        console.log('Storage Key:', storageKey);
        console.log('Total submissions:', allSubmissions.length);
        
        showStatus('Form data saved successfully!', 'success');
        
        // Reset form after successful submission
        setTimeout(() => {
            document.getElementById('underwritingForm').reset();
            submitButton.disabled = false;
            submitButton.textContent = 'Submit';
        }, 1500);
        
        // Optionally close the dialog after submission
        // setTimeout(() => closeForm(), 3000);
        
    } catch (error) {
        console.error('Error saving form data:', error);
        showStatus('Failed to save form: ' + error.message, 'error');
        submitButton.disabled = false;
        submitButton.textContent = 'Submit';
    }
}

function closeForm() {
    // If opened as a dialog, send a message to close
    if (Office.context.ui.messageParent) {
        Office.context.ui.messageParent(JSON.stringify({ action: 'close' }));
    }
    
    // If in taskpane, just show a message
    console.log('Close form requested');
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('statusMessage');
    if (statusDiv) {
        statusDiv.textContent = message;
        statusDiv.className = 'status ' + type;
        statusDiv.style.display = 'block';
        
        // Auto-hide after 5 seconds for success messages
        if (type === 'success') {
            setTimeout(() => {
                statusDiv.style.display = 'none';
            }, 5000);
        }
    }
}

// Helper function to retrieve all stored submissions (you can call this from console)
function getAllSubmissions() {
    const submissions = JSON.parse(localStorage.getItem('underwriting_submissions') || '[]');
    console.log('All submissions:', submissions);
    return submissions;
}

// Helper function to get a specific submission by key
function getSubmission(key) {
    const data = localStorage.getItem(key);
    return data ? JSON.parse(data) : null;
}

// Helper function to export all data (you can call this from console)
function exportAllData() {
    const submissions = JSON.parse(localStorage.getItem('underwriting_submissions') || '[]');
    const allData = submissions.map(sub => {
        const data = localStorage.getItem(sub.key);
        return data ? JSON.parse(data) : null;
    }).filter(item => item !== null);
    
    console.log('Exported all data:', allData);
    
    // Download as JSON file
    const dataStr = JSON.stringify(allData, null, 2);
    const dataBlob = new Blob([dataStr], {type: 'application/json'});
    const url = URL.createObjectURL(dataBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `underwriting_submissions_${new Date().toISOString()}.json`;
    link.click();
    
    return allData;
}

// Helper function to clear all stored data (you can call this from console)
function clearAllSubmissions() {
    const submissions = JSON.parse(localStorage.getItem('underwriting_submissions') || '[]');
    submissions.forEach(sub => {
        localStorage.removeItem(sub.key);
    });
    localStorage.removeItem('underwriting_submissions');
    console.log('All submissions cleared');
}
