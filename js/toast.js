/**
 * Toast Notification System
 * Simple, non-blocking notifications to replace alert() calls
 */

class ToastManager {
    constructor() {
        this.container = null;
        this.init();
    }

    init() {
        // Create container if it doesn't exist
        if (!this.container) {
            this.container = document.createElement('div');
            this.container.className = 'toast-container';
            document.body.appendChild(this.container);
        }
    }

    /**
     * Show a toast notification
     * @param {string} message - The message to display
     * @param {Object} options - Configuration options
     * @param {string} options.type - 'info' | 'error' | 'success' | 'warning'
     * @param {string} options.title - Optional title
     * @param {number} options.duration - Auto-dismiss duration in ms (0 = no auto-dismiss)
     */
    show(message, options = {}) {
        const { type = 'info', title = '', duration = 5000 } = options;

        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;

        const icon = this.getIcon(type);

        toast.innerHTML = `
            <span class="toast-icon">${icon}</span>
            <div class="toast-content">
                ${title ? `<div class="toast-title">${this.escapeHtml(title)}</div>` : ''}
                <div class="toast-message">${this.escapeHtml(message)}</div>
            </div>
            <button class="toast-close" aria-label="Close">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M18 6L6 18M6 6l12 12"/>
                </svg>
            </button>
        `;

        // Close button handler
        const closeBtn = toast.querySelector('.toast-close');
        closeBtn.addEventListener('click', () => this.dismiss(toast));

        this.container.appendChild(toast);

        // Auto-dismiss
        if (duration > 0) {
            setTimeout(() => this.dismiss(toast), duration);
        }

        return toast;
    }

    dismiss(toast) {
        if (!toast || toast.classList.contains('toast-hiding')) return;

        toast.classList.add('toast-hiding');
        setTimeout(() => {
            if (toast.parentNode) {
                toast.parentNode.removeChild(toast);
            }
        }, 300);
    }

    getIcon(type) {
        switch (type) {
            case 'error':
                return `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <circle cx="12" cy="12" r="10"/>
                    <path d="M15 9l-6 6M9 9l6 6"/>
                </svg>`;
            case 'success':
                return `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <circle cx="12" cy="12" r="10"/>
                    <path d="M9 12l2 2 4-4"/>
                </svg>`;
            case 'warning':
                return `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M12 9v4M12 17h.01"/>
                    <path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/>
                </svg>`;
            default: // info
                return `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <circle cx="12" cy="12" r="10"/>
                    <path d="M12 16v-4M12 8h.01"/>
                </svg>`;
        }
    }

    escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    // Convenience methods
    info(message, title = '') {
        return this.show(message, { type: 'info', title });
    }

    error(message, title = 'Error') {
        return this.show(message, { type: 'error', title, duration: 8000 });
    }

    success(message, title = '') {
        return this.show(message, { type: 'success', title });
    }

    warning(message, title = 'Warning') {
        return this.show(message, { type: 'warning', title, duration: 6000 });
    }
}

// Create and export singleton instance
const toast = new ToastManager();
export { toast, ToastManager };
