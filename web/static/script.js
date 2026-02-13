// Client-side JavaScript for the Malayalam Church Songs Generator

// Show message to user
function showMessage(message, type) {
    // Create or update alert div
    let alertDiv = document.querySelector('.alert');
    if (!alertDiv) {
        alertDiv = document.createElement('div');
        const container = document.querySelector('.container');
        const header = container.querySelector('header');
        container.insertBefore(alertDiv, header.nextSibling);
    }
    
    alertDiv.className = `alert alert-${type}`;
    alertDiv.textContent = message;
    alertDiv.style.display = 'block';
    
    // Auto-hide after 5 seconds
    setTimeout(() => {
        alertDiv.style.display = 'none';
    }, 5000);
}

// Extract hymns based on selected language
function extractHymns() {
    const language = document.getElementById('language').value;
    window.location.href = `/extract_hymns/${language}`;
}

// Update button text based on language
function updateExtractButtonText() {
    const language = document.getElementById('language').value;
    const extractBtn = document.getElementById('extractBtn');
    if (extractBtn) {
        extractBtn.textContent = `📊 Extract ${language} Hymns to Excel`;
    }
}

// Update default songs text based on language
function updateDefaultSongs() {
    const language = document.getElementById('language').value;
    const songsTextarea = document.getElementById('songs_text');
    
    if (songsTextarea) {
        const defaultMalayalam = songsTextarea.getAttribute('data-malayalam');
        const defaultEnglish = songsTextarea.getAttribute('data-english');
        
        // Try to load from sessionStorage first
        const storageKey = `songs_${language}`;
        const savedSongs = sessionStorage.getItem(storageKey);
        
        if (savedSongs) {
            // Use saved content for this language
            songsTextarea.value = savedSongs;
        } else {
            // Use default for the selected language
            if (language === 'English') {
                songsTextarea.value = defaultEnglish;
            } else {
                songsTextarea.value = defaultMalayalam;
            }
        }
    }
}

// Save current songs list to sessionStorage
function saveSongsToSession() {
    const language = document.getElementById('language').value;
    const songsTextarea = document.getElementById('songs_text');
    
    if (songsTextarea) {
        const storageKey = `songs_${language}`;
        sessionStorage.setItem(storageKey, songsTextarea.value);
    }
}

// Update PPT count when language changes
function updatePPTCount() {
    const language = document.getElementById('language').value;
    const countElement = document.getElementById('ppt-count');
    
    // Update extract button text
    updateExtractButtonText();
    
    // Update default songs
    updateDefaultSongs();
    
    if (countElement) {
        countElement.textContent = 'Loading...';
        
        fetch(`/ppt_count/${language}`)
            .then(response => response.json())
            .then(data => {
                if (data.count > 0) {
                    countElement.textContent = `${data.count} PowerPoint files available`;
                    countElement.style.color = '#28a745';
                } else {
                    countElement.textContent = 'No PowerPoint files found';
                    countElement.style.color = '#dc3545';
                }
            })
            .catch(err => {
                countElement.textContent = 'Could not load file count';
                countElement.style.color = '#666';
            });
    }
}

// Update count on language change
document.getElementById('language')?.addEventListener('change', updatePPTCount);

// Output console functions
function appendOutput(text) {
    const outputText = document.getElementById('outputText');
    
    if (outputText) {
        // Clear the default message on first output
        if (outputText.textContent === 'Ready to generate presentations...') {
            outputText.textContent = '';
        }
        
        outputText.textContent += text + '\n';
        // Auto-scroll to bottom
        outputText.scrollTop = outputText.scrollHeight;
    } else {
        console.error('Output text element not found');
    }
}

function clearOutput() {
    const outputText = document.getElementById('outputText');
    
    if (outputText) {
        outputText.textContent = 'Ready to generate presentations...';
    }
}

// Load initial count and setup form on page load
window.addEventListener('DOMContentLoaded', function() {
    console.log('DOM Content Loaded');
    
    updatePPTCount();
    
    // Auto-cleanup old files
    fetch('/cleanup')
        .catch(err => console.log('Cleanup skipped:', err));
    
    // Setup form submission handler
    const songForm = document.getElementById('songForm');
    console.log('Song form found:', songForm);
    
    if (!songForm) {
        console.error('Form with id="songForm" not found!');
        return;
    }
    
    songForm.addEventListener('submit', handleFormSubmit);
});

// Form submission handler
function handleFormSubmit(e) {
    e.preventDefault();
    
    console.log('Form submitted!');
    
    // Save current songs list to session before generating
    saveSongsToSession();
    
    const songsText = document.getElementById('songs_text').value.trim();
    
    if (!songsText) {
        alert('Please enter songs.');
        return false;
    }
    
    // Clear previous output and show console
    clearOutput();
    appendOutput('🚀 Starting presentation generation...');
    
    const formData = new FormData(e.target);
    const submitBtn = e.target.querySelector('button[type="submit"]');
    const originalText = submitBtn.innerHTML;
    
    submitBtn.disabled = true;
    submitBtn.classList.add('generating');
    submitBtn.textContent = 'Generating...';
    
    const language = document.getElementById('language').value;
    appendOutput(`📋 Language: ${language}`);
    appendOutput(`📝 Processing songs...`);
    
    fetch('/generate', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        appendOutput(`📡 Response received (status: ${response.status})`);
        
        if (!response.ok) {
            return response.text().then(text => {
                throw new Error(text || 'Generation failed');
            });
        }
        
        const contentType = response.headers.get('content-type');
        if (contentType && contentType.includes('application/json')) {
            return response.json();
        } else {
            appendOutput('💾 Downloading presentation...');
            return response.blob().then(blob => {
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                
                // Extract filename from Content-Disposition header
                const contentDisposition = response.headers.get('Content-Disposition');
                let filename = 'Holy_Communion_Service.pptx'; // fallback
                
                if (contentDisposition) {
                    const filenameMatch = contentDisposition.match(/filename="?(.+)"?/i);
                    if (filenameMatch && filenameMatch[1]) {
                        filename = filenameMatch[1].replace(/"/g, '');
                    }
                }
                
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                appendOutput('✅ Presentation generated successfully!');
                appendOutput(`📥 Download started: ${filename}`);
                
                // Check if there's a detailed log to fetch
                const genId = response.headers.get('X-Generation-ID');
                if (genId) {
                    appendOutput('\n📋 Detailed Generation Log:');
                    fetch(`/get_log/${genId}`)
                        .then(r => r.json())
                        .then(data => {
                            if (data.log && data.log.length > 0) {
                                data.log.forEach(line => {
                                    if (line.trim()) {
                                        // Replace server path with user-friendly message
                                        let outputLine = line.replace(/Presentation saved:.*\.pptx/, 'Presentation downloaded to your Downloads folder');
                                        outputLine = outputLine.replace(/^.*\/generated\/.*\.pptx$/, '');
                                        if (outputLine.trim()) {
                                            appendOutput(outputLine);
                                        }
                                    }
                                });
                                appendOutput('\n✅ File saved to your Downloads folder');
                                appendOutput('📁 Check your browser\'s download location');
                            }
                        })
                        .catch(err => console.error('Failed to fetch log:', err));
                }
                
                showMessage('Presentation generated successfully!', 'success');
            });
        }
    })
    .then(data => {
        if (data && data.message) {
            appendOutput('✅ ' + data.message);
            showMessage(data.message, 'success');
        }
    })
    .catch(error => {
        appendOutput('❌ Error: ' + error.message);
        showMessage('Error: ' + error.message, 'error');
        console.error('Generation error:', error);
    })
    .finally(() => {
        submitBtn.disabled = false;
        submitBtn.classList.remove('generating');
        submitBtn.innerHTML = originalText;
    });
    
    return false;
}
