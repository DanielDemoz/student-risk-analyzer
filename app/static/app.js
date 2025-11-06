// Student Risk Analyzer Frontend JavaScript

let currentResults = [];

// Get API endpoint from localStorage or use default
function getApiEndpoint() {
    const saved = localStorage.getItem('apiEndpoint');
    if (saved) {
        return saved;
    }
    // Default to localhost for local development, or detect if on GitHub Pages
    if (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1') {
        return 'http://localhost:8000';
    }
    // For GitHub Pages, you'll need to configure your API endpoint
    return window.location.origin; // Fallback to same origin
}

// Save API endpoint
function saveApiEndpoint() {
    const endpoint = document.getElementById('apiEndpoint').value.trim();
    if (endpoint) {
        // Remove trailing slash
        const cleanEndpoint = endpoint.replace(/\/$/, '');
        localStorage.setItem('apiEndpoint', cleanEndpoint);
        showSuccess('API endpoint saved successfully!');
        // Test the connection
        testApiConnection(cleanEndpoint);
    }
}

// Test API connection
async function testApiConnection(endpoint) {
    if (!endpoint) {
        showError('Please enter an API endpoint URL first.');
        return;
    }
    
    // Show loading state
    const errorAlert = document.getElementById('errorAlert');
    const successAlert = document.getElementById('successAlert');
    errorAlert.classList.add('d-none');
    successAlert.classList.add('d-none');
    
    try {
        const response = await fetch(`${endpoint}/health`, {
            method: 'GET',
            signal: AbortSignal.timeout(5000) // 5 second timeout
        });
        
        if (response.ok) {
            const data = await response.json();
            showSuccess(`✅ Connection successful! Server is running at ${endpoint}`);
        } else {
            showError(`⚠️ Server responded but with an error (${response.status}). Please check your endpoint.`);
        }
    } catch (error) {
        if (error.name === 'AbortError') {
            showError('⏱️ Connection timeout. Please check that the server is running and the URL is correct.');
        } else if (error.message.includes('Failed to fetch') || error.message.includes('NetworkError')) {
            showError(`❌ Cannot connect to server at ${endpoint}. Make sure:\n1. The server is running (use: python -m uvicorn app.main:app --reload)\n2. The URL is correct (e.g., http://localhost:8000)\n3. CORS is enabled on the server\n4. No firewall is blocking the connection`);
        } else {
            showError(`❌ Connection error: ${error.message}`);
        }
    }
}

// Test connection button handler
function testConnection() {
    const endpoint = document.getElementById('apiEndpoint').value.trim();
    if (!endpoint) {
        showError('Please enter an API endpoint URL first.');
        return;
    }
    // Remove trailing slash
    const cleanEndpoint = endpoint.replace(/\/$/, '');
    document.getElementById('apiEndpoint').value = cleanEndpoint;
    testApiConnection(cleanEndpoint);
}

// Initialize
document.addEventListener('DOMContentLoaded', function() {
    const uploadForm = document.getElementById('uploadForm');
    const searchInput = document.getElementById('searchInput');
    const exportCsvBtn = document.getElementById('exportCsvBtn');
    const copyEmailBtn = document.getElementById('copyEmailBtn');
    const apiEndpointInput = document.getElementById('apiEndpoint');

    // Load saved API endpoint
    if (apiEndpointInput) {
        apiEndpointInput.value = getApiEndpoint();
        // Test connection on page load if not localhost
        if (window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1') {
            const endpoint = getApiEndpoint();
            if (endpoint && endpoint !== window.location.origin) {
                // Only test if it's a different origin
                testApiConnection(endpoint);
            }
        }
    }

    // Upload form handler
    uploadForm.addEventListener('submit', handleUpload);

    // Search handler
    if (searchInput) {
        searchInput.addEventListener('input', handleSearch);
    }

    // Export CSV handler
    if (exportCsvBtn) {
        exportCsvBtn.addEventListener('click', handleExportCsv);
    }

    // Copy email handler
    if (copyEmailBtn) {
        copyEmailBtn.addEventListener('click', handleCopyEmail);
    }
});

// Handle file upload
async function handleUpload(e) {
    e.preventDefault();

    const fileInput = document.getElementById('fileInput');
    const processBtn = document.getElementById('processBtn');
    const spinner = document.getElementById('spinner');
    const errorAlert = document.getElementById('errorAlert');
    const successAlert = document.getElementById('successAlert');

    if (!fileInput.files || fileInput.files.length === 0) {
        showError('Please select a file to upload');
        return;
    }

    // Show spinner and progress
    processBtn.disabled = true;
    spinner.classList.remove('d-none');
    document.getElementById('btnText').textContent = 'Processing...';
    const progressBar = document.getElementById('progressBar');
    progressBar.classList.remove('d-none');
    errorAlert.classList.add('d-none');
    successAlert.classList.add('d-none');

    // Prepare form data
    const formData = new FormData();
    formData.append('file', fileInput.files[0]);

    try {
        const apiEndpoint = getApiEndpoint();
        const uploadUrl = `${apiEndpoint}/upload`;
        
        const response = await fetch(uploadUrl, {
            method: 'POST',
            body: formData
        });

        // Check if response is JSON
        const contentType = response.headers.get('content-type');
        let data;
        
        if (contentType && contentType.includes('application/json')) {
            data = await response.json();
        } else {
            // If not JSON, read as text to get error message
            const text = await response.text();
            throw new Error(`Server error: ${response.status} ${response.statusText}. ${text.substring(0, 200)}`);
        }

        if (!response.ok) {
            throw new Error(data.detail || data.message || 'Upload failed');
        }

        // Store results
        currentResults = data.results;

        // Display results
        displayResults(data);
        showSuccess(data.message);

    } catch (error) {
        // Handle JSON parsing errors
        if (error instanceof SyntaxError) {
            showError('Server returned invalid response. Please check the server logs.');
        } else if (error.message.includes('Failed to fetch') || error.message.includes('NetworkError')) {
            const endpoint = getApiEndpoint();
            showError(`Unable to connect to the server at ${endpoint}. Please:\n1. Make sure the server is running (python -m uvicorn app.main:app --reload)\n2. Check the API endpoint in Advanced settings\n3. Verify the URL is correct (e.g., http://localhost:8000)`);
        } else {
            showError(error.message || 'An error occurred while processing the file');
        }
    } finally {
        processBtn.disabled = false;
        spinner.classList.add('d-none');
        document.getElementById('btnText').textContent = 'Analyze Student Risk';
        document.getElementById('progressBar').classList.add('d-none');
    }
}

// Display results
function displayResults(data) {
    const resultsSection = document.getElementById('resultsSection');
    const resultsTableBody = document.getElementById('resultsTableBody');

    // Show results section
    resultsSection.classList.remove('d-none');

    // Update summary with correct category matching
    const failedCount = data.summary.Failed || 0;
    const highCount = data.summary.High || 0;
    const mediumCount = data.summary.Medium || 0;
    const lowCount = data.summary.Low || 0;
    
    document.getElementById('failedCount').textContent = failedCount;
    document.getElementById('highCount').textContent = highCount;
    document.getElementById('mediumCount').textContent = mediumCount;
    document.getElementById('lowCount').textContent = lowCount;
    
    // Debug: Log summary to console
    console.log('Summary counts:', { failedCount, highCount, mediumCount, lowCount, total: data.summary.Total });
    document.getElementById('totalCount').textContent = data.summary.Total || 0;
    
    // Calculate and display average grade and attendance
    if (data.results && data.results.length > 0) {
        const avgGrade = data.results.reduce((sum, r) => sum + r.grade_pct, 0) / data.results.length;
        const avgAttendance = data.results.reduce((sum, r) => sum + r.attendance_pct, 0) / data.results.length;
        
        document.getElementById('avgGrade').textContent = avgGrade.toFixed(1) + '%';
        document.getElementById('avgAttendance').textContent = avgAttendance.toFixed(1) + '%';
        document.getElementById('additionalStats').classList.remove('d-none');
    }

    // Clear table
    resultsTableBody.innerHTML = '';

    // Populate table
    data.results.forEach(result => {
        const row = createTableRow(result);
        resultsTableBody.appendChild(row);
    });

    // Initialize tooltips
    const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    tooltipTriggerList.map(function (tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl);
    });
}

// Create table row
function createTableRow(result) {
    const row = document.createElement('tr');

    // Student ID
    const studentId = result.student_id || 'N/A';
    
    // Risk category badge with color coding
    // Use risk_color from API if available, otherwise fallback to Bootstrap classes
    const riskColor = result.risk_color || '#6B7280';  // Default gray
    const badgeStyle = `background-color: ${riskColor}; color: white; font-weight: 600;`;
    const badge = `<span class="badge" style="${badgeStyle}">${result.risk_category.toUpperCase()}</span>`;

    // Student name with link (for Campus Login - but we'll show name separately in table)
    const studentNameLink = `<a href="${result.campus_login_url}" target="_blank">${result.student_name || 'Unknown'}</a>`;

    // Explanation tooltip
    let explanationAttr = '';
    if (result.explanation) {
        explanationAttr = `data-bs-toggle="tooltip" data-bs-placement="top" title="${result.explanation}"`;
    }

    // Risk Assessment combines score and category
    const riskAssessment = `${result.risk_score.toFixed(1)} (${result.risk_category})`;
    
    // Student name (not linked, just text)
    const studentName = result.student_name || 'Unknown';
    
    row.innerHTML = `
        <td>${studentName}</td>
        <td>${studentId}</td>
        <td>${result.program_name}</td>
        <td>${result.grade_pct.toFixed(1)}%</td>
        <td>${result.attendance_pct.toFixed(1)}%</td>
        <td ${explanationAttr}>
            <div class="d-flex align-items-center gap-2">
                <span>${result.risk_score.toFixed(1)}</span>
                ${badge}
            </div>
        </td>
        <td>
            <button class="btn btn-sm brukd-btn-primary btn-action" onclick="openCampusLogin('${result.campus_login_url.replace(/'/g, "\\'")}')">
                Open Campus Login
            </button>
            <button class="btn btn-sm brukd-btn-secondary btn-action" onclick="showEmailDraft('${result.student_id.replace(/'/g, "\\'")}', '${result.risk_category}', '${result.program_name.replace(/'/g, "\\'")}', ${result.grade_pct}, ${result.attendance_pct})">
                Email Draft
            </button>
        </td>
    `;

    return row;
}

// Handle search
function handleSearch(e) {
    const searchTerm = e.target.value.toLowerCase();
    const rows = document.querySelectorAll('#resultsTableBody tr');

    rows.forEach(row => {
        const text = row.textContent.toLowerCase();
        if (text.includes(searchTerm)) {
            row.style.display = '';
        } else {
            row.style.display = 'none';
        }
    });
}

// Open campus login
function openCampusLogin(url) {
    window.open(url, '_blank');
}

// Show email draft modal
async function showEmailDraft(studentId, riskCategory, program, gradePct, attendancePct) {
    try {
        const apiEndpoint = getApiEndpoint();
        const emailDraftUrl = `${apiEndpoint}/email-draft`;
        
        const response = await fetch(emailDraftUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                student_id: studentId,
                risk_category: riskCategory,
                program: program,
                grade_pct: gradePct,
                attendance_pct: attendancePct
            })
        });

        // Check if response is JSON
        const contentType = response.headers.get('content-type');
        let data;
        
        if (contentType && contentType.includes('application/json')) {
            data = await response.json();
        } else {
            // If not JSON, read as text to get error message
            const text = await response.text();
            throw new Error(`Server error: ${response.status} ${response.statusText}. ${text.substring(0, 200)}`);
        }

        if (!response.ok) {
            throw new Error(data.detail || data.message || 'Failed to generate email draft');
        }

        // Populate modal
        document.getElementById('emailSubject').value = data.subject;
        document.getElementById('emailBody').value = data.body;

        // Store email data for copying
        window.currentEmailData = data;

        // Show modal
        const modal = new bootstrap.Modal(document.getElementById('emailModal'));
        modal.show();

    } catch (error) {
        // Handle JSON parsing errors
        if (error instanceof SyntaxError) {
            alert('Server returned invalid response. Please check the server logs.');
        } else {
            alert('Error generating email draft: ' + error.message);
        }
    }
}

// Copy email to clipboard
function handleCopyEmail() {
    const emailData = window.currentEmailData;
    if (!emailData) {
        return;
    }

    const emailText = `Subject: ${emailData.subject}\n\n${emailData.body}`;

    navigator.clipboard.writeText(emailText).then(() => {
        const btn = document.getElementById('copyEmailBtn');
        const originalText = btn.textContent;
        btn.textContent = 'Copied!';
        btn.classList.add('btn-success');
        btn.classList.remove('btn-primary');

        setTimeout(() => {
            btn.textContent = originalText;
            btn.classList.remove('btn-success');
            btn.classList.add('btn-primary');
        }, 2000);
    }).catch(err => {
        alert('Failed to copy email: ' + err.message);
    });
}

// Export CSV
function handleExportCsv() {
    const apiEndpoint = getApiEndpoint();
    const csvUrl = `${apiEndpoint}/download.csv`;
    window.location.href = csvUrl;
}

// Show error message
function showError(message) {
    const errorAlert = document.getElementById('errorAlert');
    const errorMessage = document.getElementById('errorMessage');
    errorMessage.textContent = message;
    errorAlert.classList.remove('d-none');
    errorAlert.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// Show success message
function showSuccess(message) {
    const successAlert = document.getElementById('successAlert');
    const successMessage = document.getElementById('successMessage');
    successMessage.textContent = message;
    successAlert.classList.remove('d-none');
    successAlert.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

