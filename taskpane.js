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
        const flowUrl = '<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
</head>
<body>
    <script>
        Office.onReady(() => {
            console.log('Commands.html loaded and ready');
        });

        // This function is called when the ribbon button is clicked
        async function triggerFlowFromRibbon(event) {
            console.log('Underwriting button clicked from ribbon');
            
            try {
                // Show notification that processing has started
                Office.context.mailbox.item.notificationMessages.addAsync(
                    "progress",
                    {
                        type: "progressIndicator",
                        message: "Analyzing email and triggering flow..."
                    }
                );
                
                // Get email data including body for analysis
                const emailData = await getEmailData();
                
                console.log('Email data collected:', emailData);
                
                // Your Power Automate flow URL
                const flowUrl = 'https://default74afe875305e4ab4ba4ac1359a7629.ae.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/89c12382226642a4907cd110e9e7ab87/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Nbz7sUIbNoHlSBt_KVnF3CFKCCf9lPYn-LbIxZsWouA';
                
                // Call Power Automate
                const response = await fetch(flowUrl, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify(emailData)
                });
                
                // Remove progress indicator
                Office.context.mailbox.item.notificationMessages.removeAsync("progress");
                
                if (response.ok) {
                    // Show success notification
                    Office.context.mailbox.item.notificationMessages.addAsync(
                        "success",
                        {
                            type: "informationalMessage",
                            message: "Email analyzed and flow triggered successfully! ✓",
                            icon: "Icon.80x80",
                            persistent: false
                        }
                    );
                    
                    console.log('Flow triggered successfully');
                } else {
                    const errorText = await response.text();
                    throw new Error(`HTTP ${response.status}: ${errorText}`);
                }
                
            } catch (error) {
                console.error('Error triggering flow:', error);
                
                // Remove progress indicator if still showing
                Office.context.mailbox.item.notificationMessages.removeAsync("progress");
                
                // Show error notification
                Office.context.mailbox.item.notificationMessages.addAsync(
                    "error",
                    {
                        type: "errorMessage",
                        message: "Failed to Underwriting: " + error.message
                    }
                );
            } finally {
                // Signal that the function has completed
                event.completed();
            }
        }

        async function getEmailData() {
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
                        // In read mode, subject is a property
                        if (typeof item.subject === 'string') {
                            resolveSubject(item.subject);
                        } else if (item.subject && typeof item.subject.getAsync === 'function') {
                            item.subject.getAsync((result) => {
                                if (result.status === Office.AsyncResultStatus.Succeeded) {
                                    resolveSubject(result.value);
                                } else {
                                    console.error('Error getting subject:', result.error);
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
                                    console.error('Error getting body:', result.error);
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
                    
                    console.log('Email data prepared:', {
                        subject: emailData.subject,
                        from: emailData.from,
                        bodyLength: emailData.body.length,
                        attachments: emailData.attachmentCount
                    });
                    
                    resolve(emailData);
                }).catch(error => {
                    console.error('Error collecting email data:', error);
                    reject(error);
                });
            });
        }

        // CRITICAL: Register the function so Office.js can find it
        Office.actions.associate("triggerFlowFromRibbon", triggerFlowFromRibbon);
    </script>
</body>
</html>';
        
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
            showStatus('Flow triggered successfully!', 'success');
        } else {
            const errorText = await response.text();
            throw new Error(`HTTP error! status: ${response.status}, details: ${errorText}`);
        }
        
    } catch (error) {
        console.error('Error triggering flow:', error);
        showStatus('Failed to Underwriting: ' + error.message, 'error');
    } finally {
        // Re-enable button
        button.disabled = false;
        button.textContent = 'Underwriting';
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
            
            // For compose mode or if you need the body, you'd use async methods
            // But for basic info in read mode, properties work fine
            
            // Prepare data to send to Power Automate
            const emailData = {
                subject: subject,
                from: from,
                itemId: itemId,
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
