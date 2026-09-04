// Initialize date display
document.addEventListener('DOMContentLoaded', function() {
    updateDateTime();
    initializeTooltips();
    setupFormValidation();
    setupAutoDismiss();
});

// Update date and time
function updateDateTime() {
    const dateElement = document.getElementById('currentDate');
    if (dateElement) {
        const options = { 
            weekday: 'long', 
            year: 'numeric', 
            month: 'long', 
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        };
        const now = new Date();
        dateElement.textContent = now.toLocaleDateString('en-US', options);
    }
}

// Initialize tooltips
function initializeTooltips() {
    document.querySelectorAll('[data-tooltip]').forEach(element => {
        element.addEventListener('mouseenter', showTooltip);
        element.addEventListener('mouseleave', hideTooltip);
    });
}

function showTooltip(e) {
    const tooltip = document.createElement('div');
    tooltip.className = 'tooltip';
    tooltip.textContent = e.target.dataset.tooltip;
    document.body.appendChild(tooltip);
    
    const rect = e.target.getBoundingClientRect();
    tooltip.style.top = rect.bottom + 10 + 'px';
    tooltip.style.left = rect.left + 'px';
}

function hideTooltip() {
    document.querySelectorAll('.tooltip').forEach(t => t.remove());
}

// Form validation
function setupFormValidation() {
    document.querySelectorAll('form').forEach(form => {
        form.addEventListener('submit', validateForm);
    });
}

function validateForm(e) {
    let isValid = true;
    const form = e.target;
    
    form.querySelectorAll('[required]').forEach(field => {
        if (!field.value.trim()) {
            field.style.borderColor = '#ef4444';
            isValid = false;
            
            // Add error message
            let error = field.parentNode.querySelector('.error-message');
            if (!error) {
                error = document.createElement('span');
                error.className = 'error-message';
                error.textContent = 'This field is required';
                field.parentNode.appendChild(error);
            }
        } else {
            field.style.borderColor = '#e2e8f0';
            const error = field.parentNode.querySelector('.error-message');
            if (error) error.remove();
        }
    });
    
    if (!isValid) {
        e.preventDefault();
        showNotification('Please fill all required fields', 'error');
    }
}

// Show notification
function showNotification(message, type = 'info') {
    const container = document.querySelector('.flash-container') || createFlashContainer();
    
    const notification = document.createElement('div');
    notification.className = `flash-message flash-${type}`;
    notification.innerHTML = `
        <i class="fas ${type === 'success' ? 'fa-check-circle' : 'fa-exclamation-circle'}"></i>
        <span>${message}</span>
        <button class="flash-close" onclick="this.parentElement.remove()">
            <i class="fas fa-times"></i>
        </button>
    `;
    
    container.appendChild(notification);
    
    setTimeout(() => {
        notification.style.animation = 'slideOut 0.3s ease forwards';
        setTimeout(() => notification.remove(), 300);
    }, 5000);
}

function createFlashContainer() {
    const container = document.createElement('div');
    container.className = 'flash-container';
    document.querySelector('.content-area').prepend(container);
    return container;
}

// Auto-dismiss flash messages
function setupAutoDismiss() {
    document.querySelectorAll('.flash-message').forEach(flash => {
        setTimeout(() => {
            flash.style.animation = 'slideOut 0.3s ease forwards';
            setTimeout(() => flash.remove(), 300);
        }, 5000);
    });
}

// Confirmation dialog
function confirmAction(message, callback) {
    const modal = document.createElement('div');
    modal.className = 'modal';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <h3>Confirm Action</h3>
                <button class="modal-close" onclick="this.closest('.modal').remove()">
                    <i class="fas fa-times"></i>
                </button>
            </div>
            <div class="modal-body">
                <p>${message}</p>
            </div>
            <div class="modal-footer">
                <button class="btn btn-secondary" onclick="this.closest('.modal').remove()">Cancel</button>
                <button class="btn btn-primary" id="confirmBtn">Confirm</button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    document.getElementById('confirmBtn').addEventListener('click', function() {
        modal.remove();
        callback();
    });
}

// Loading spinner
function showLoading() {
    const spinner = document.createElement('div');
    spinner.className = 'loading-spinner';
    spinner.innerHTML = '<div class="spinner"></div>';
    document.body.appendChild(spinner);
}

function hideLoading() {
    document.querySelectorAll('.loading-spinner').forEach(s => s.remove());
}

// Print functionality
function printReport() {
    window.print();
}

// Download as text file
function downloadAsText(content, filename) {
    const blob = new Blob([content], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}

// Copy to clipboard
function copyToClipboard(text) {
    navigator.clipboard.writeText(text).then(() => {
        showNotification('Copied to clipboard', 'success');
    }).catch(() => {
        showNotification('Failed to copy', 'error');
    });
}

// Debounce function
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// Throttle function
function throttle(func, limit) {
    let inThrottle;
    return function(...args) {
        if (!inThrottle) {
            func.apply(this, args);
            inThrottle = true;
            setTimeout(() => inThrottle = false, limit);
        }
    };
}