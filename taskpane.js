let mailboxItem = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Store the mailbox item reference
        mailboxItem = Office.context.mailbox.item;
        
        // Display current email subject
        displayEmailInfo();
        
        console.log('Office.js initialized successfully');
    }
});

function displayEmailInfo() {
    if (!mailboxItem) {
        console.error('Mailbox item not available');
        return;
    }
    
    // For Outlook, subject is a property, not a method in read mode
    try {
        const subject = mailboxItem.subject;
        document.getElementById('emailSubject').textContent = subject || 'No subject';
        document.getElementById('emailInfo').style.display = 'block';
    } catch (error) {
        console.error('Error displaying email info:', error);
    }
}

async function triggerFlow() {
    const button = document.getElementById('triggerButton');
    const statusDiv = document.getElementById('statusMessage');
    
    // Check if Office context is available
    if (!mailboxItem) {
        showStatus('Outlook context not available. Please reload the add-in.', 'error');
        return;
    }
    
    // Disable button during request
    button.disabled = true;
    button.textContent = 'Triggering...';
    
    // Show info message
    showStatus('Sending request to Power Automate...', 'info');
    
    try {
        // Get email details to send to the flow
        const emailData = await getEmailData();
        
        // Replace with your Power Automate HTTP POST URL
        const flowUrl = 'https://default74afe875305e4ab4ba4ac1359a7629.ae.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/89c12382226642a4907cd110e9e7ab87/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Nbz7sUIbNoHlSBt_KVnF3CFKCCf9lPYn-LbIxZsWouA';
        // Make the request to Power Automate
        const response = await fetch(flowUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(emailData)
        });
        
        if (response.ok) {
            const responseText = await response.text();
            console.log('Flow response:', responseText);
            showStatus('Flow triggered successfully! Report will be available shortly.', 'success');

            // After 30s, construct and display the hardcoded OneDrive link
            setTimeout(() => {
                try {
                    const link = buildHardcodedReportLink();
                    if (link) {
                        displayReportLink(link);
                    } else {
                        console.warn('Could not determine attachment filename to build report link');
                    }
                } catch (e) {
                    console.error('Error building/displaying report link:', e);
                }
            }, 30000);
        } else {
            const errorText = await response.text();
            throw new Error(`HTTP error! status: ${response.status}, details: ${errorText}`);
        }
        
    } catch (error) {
        console.error('Error triggering flow:', error);
        showStatus('Failed to Auto Process: ' + error.message, 'error');
    } finally {
        // Re-enable button
        button.disabled = false;
        button.textContent = 'Auto Process';
    }
}

async function getEmailData() {
    return new Promise((resolve, reject) => {
        if (!mailboxItem) {
            reject(new Error('Mailbox item not available'));
            return;
        }
        
        try {
            // In read mode, these are properties, not async methods
            const subject = mailboxItem.subject || 'No subject';
            const from = mailboxItem.from ? mailboxItem.from.emailAddress : 'Unknown';
            const itemId = mailboxItem.itemId || 'Unknown';
            const internetMessageId = mailboxItem.internetMessageId || '';
            // For compose mode or if you need the body, you'd use async methods
            // But for basic info in read mode, properties work fine
            
            // Prepare data to send to Power Automate
            const emailData = {
                subject: subject,
                from: from,
                itemId: itemId,
                internetMessageId: internetMessageId,
                triggeredAt: new Date().toISOString(),
                email: Office.context.mailbox.userProfile.emailAddress,
                conversationId: mailboxItem.conversationId || 'Unknown'
            };
            
            console.log('Email data prepared:', emailData);
            resolve(emailData);
            
        } catch (error) {
            console.error('Error getting email data:', error);
            reject(error);
        }
    });
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('statusMessage');
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

// Build the hardcoded OneDrive report link using the first PDF attachment's base name
function buildHardcodedReportLink() {
    if (!mailboxItem || !mailboxItem.attachments || mailboxItem.attachments.length === 0) return null;

    // Prefer first attachment that ends with .pdf
    let att = mailboxItem.attachments.find(a => typeof a.name === 'string' && a.name.toLowerCase().endsWith('.pdf')) || mailboxItem.attachments[0];
    if (!att || !att.name) return null;

    // Remove extension and append .html
    const baseName = att.name.replace(/\.[^.]+$/, '');
    const htmlName = baseName + '.html';

    const baseUrl = 'https://datasciencewiizardsai-my.sharepoint.com/personal/karan_panchal_datasciencewizards_ai/Documents/Output_attachments/';
    return baseUrl + encodeURIComponent(htmlName);
}

// Update the UI to show the report link labeled 'Report'
function displayReportLink(url) {
    const container = document.getElementById('reportContainer');
    const linkElem = document.getElementById('reportLink');
    if (!container || !linkElem) return;

    // Create anchor
    const a = document.createElement('a');
    a.href = url;
    a.target = '_blank';
    a.textContent = 'Report';
    a.style.color = '#0078d4';

    // Clear previous and append
    linkElem.innerHTML = '';
    linkElem.appendChild(document.createTextNode('Your report is ready: '));
    linkElem.appendChild(a);

    container.style.display = 'block';
}
