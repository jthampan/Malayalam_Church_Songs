// Client-side JavaScript for the Malayalam Church Songs Generator

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

// Update PPT count when language changes
function updatePPTCount() {
    const language = document.getElementById('language').value;
    const countElement = document.getElementById('ppt-count');
    
    // Update extract button text
    updateExtractButtonText();
    
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

// Load initial count on page load
window.addEventListener('load', function() {
    updatePPTCount();
    
    // Auto-cleanup old files
    fetch('/cleanup')
        .catch(err => console.log('Cleanup skipped:', err));
});

// Form validation
document.querySelector('form')?.addEventListener('submit', function(e) {
    const songsText = document.getElementById('songs_text').value.trim();
    
    if (!songsText) {
        e.preventDefault();
        alert('Please enter songs.');
        return false;
    }
    
    // Show loading indicator
    const button = this.querySelector('.btn-generate');
    if (button) {
        button.textContent = 'Generating PowerPoint...';
        button.disabled = true;
    }
});
