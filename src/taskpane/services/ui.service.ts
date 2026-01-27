import { CONFIG } from "../utils/config";

export class UIService {

    static showNotification(message: string, type: 'success' | 'error' | 'info' = 'info'): void {
        console.log(`[${type.toUpperCase()}] ${message}`);
        // Simple toastr wrapper, assuming toastr function is available globally or via import if we were stricter
        // For this refactor, we keep it compatible with existing 'toaster' global if possible, or implement simple DOM manipulation
        const toastContainer = document.getElementById('toastr'); // Hypothetical, in reality the old app used 'toaster()'
        if (typeof (window as any).toaster === 'function') {
            (window as any).toaster(message, type);
        } else {
            // Fallback
            alert(message);
        }
    }

    static toggleLoader(loading: boolean): void {
        const loader = document.getElementById('page-loader');
        if (loader) loader.style.display = loading ? 'flex' : 'none';
    }

    static renderLoginPage(storedUrl: string, handleLoginCallback: (e: Event) => void, themeToggleCallback: () => void): void {
        const logoHeader = document.getElementById('logo-header');
        if (logoHeader) {
            logoHeader.innerHTML = `
            <img id="main-logo" src="${storedUrl}/assets/logo.png" alt="" class="logo">
            <div class="icon-nav me-3">
                <span id="theme-toggle"><i class="fa fa-moon c-pointer me-3"  title="Toggle Theme"></i><span>
            </div>`;
        }

        const appBody = document.getElementById('app-body');
        if (appBody) {
            appBody.innerHTML = `
            <div class="container pt-2">
            <form id="login-form" class="p-4 border rounded">
                <div class="mb-3">
                <label for="organization" class="form-label fw-bold">Organization</label>
                <input type="text" class="form-control" id="organization" required>
                </div>
                <div class="mb-3">
                <label for="username" class="form-label fw-bold">Username</label>
                <input type="text" class="form-control" id="username" required>
                </div>
                <div class="mb-3">
                <label for="password" class="form-label fw-bold">Password</label>
                <input type="password" class="form-control" id="password" required>
                </div>
                <div class="d-grid">
                <button type="submit" class="btn btn-primary bg-primary-clr">Login</button>
                </div>
            <div id="login-error" class="mt-3 text-danger" style="display: none;"></div>
            </form>
            </div>`;
        }

        // Attach Event Listeners
        document.getElementById('theme-toggle')?.addEventListener('click', themeToggleCallback);
        document.getElementById('login-form')?.addEventListener('submit', handleLoginCallback);
    }

    static applyTheme(theme: 'Light' | 'Dark'): void {
        const isDark = theme === 'Dark';
        const isLight = theme === 'Light';

        const safeApplyClass = (selector: string, darkClasses: string, lightClasses: string) => {
            const elements = document.querySelectorAll(selector);
            const darkClassList = darkClasses.split(' ');
            const lightClassList = lightClasses.split(' ');

            elements.forEach(elem => {
                if (!elem) return;
                elem.classList.remove(...darkClassList);
                elem.classList.remove(...lightClassList);
                if (isDark) elem.classList.add(...darkClassList);
                if (isLight) elem.classList.add(...lightClassList);
            });
        };

        // Apply Global Toggles
        document.body.classList.toggle('dark-theme', isDark);
        document.body.classList.toggle('light-theme', isLight);

        // Apply Specific Element Classes
        safeApplyClass('#app-body', 'bg-dark text-light', 'bg-white text-dark');
        safeApplyClass('#search-box', 'bg-secondary text-light border-0', 'bg-white text-dark border');
        safeApplyClass('.dropdown-menu', 'bg-dark text-light border-light', 'bg-white text-dark border');
        safeApplyClass('.list-group-item', 'bg-dark text-light', 'bg-white text-dark');
        safeApplyClass('.dropdown-toggle', 'bg-dark text-light border-0', 'bg-white text-dark border');
        safeApplyClass('.dropdown-item', 'bg-dark text-light', 'bg-white text-dark');
        safeApplyClass('.card', 'bg-dark text-light border-secondary', 'bg-white text-dark border');
        safeApplyClass('.card-header', 'bg-secondary text-light border-secondary', 'bg-light text-dark border-bottom');
        safeApplyClass('.box', 'bg-dark text-light border-secondary', 'bg-light text-dark border');
        safeApplyClass('.modal-content', 'bg-dark text-light border-secondary', 'bg-white text-dark border');

        safeApplyClass(
            '.list-group-item-action',
            'bg-dark text-light list-hover-dark',
            'bg-light text-dark list-hover-light'
        );

        safeApplyClass('#close-ai-window', 'fa-solid fa-circle-xmark bg-dark text-light', 'fa-solid fa-circle-xmark bg-light text-dark');
        safeApplyClass('#chatInput', 'bg-secondary text-light', 'bg-white text-dark');
        safeApplyClass('.prompt-text', 'bg-secondary text-light', 'bg-white text-dark');

        // Toggle Icon
        const themeToggle = document.getElementById('theme-toggle');
        const icon = themeToggle?.querySelector('i');
        if (icon) {
            if (isDark) {
                icon.classList.remove('fa-moon');
                icon.classList.add('fa-sun');
            } else {
                icon.classList.remove('fa-sun');
                icon.classList.add('fa-moon');
            }
        }
    }

    static attachDashboardEvents(
        handlers: {
            onHome: () => void,
            onSummary: () => void,
            onGlossary: () => void,
            onFormat: () => void,
            onRemoveFormat: () => void,
            onThemeToggle: () => void,
            onLogout: () => void
        }
    ): void {
        document.getElementById('home')?.addEventListener('click', handlers.onHome);
        document.getElementById('summary-mode')?.addEventListener('click', handlers.onSummary);
        document.getElementById('glossary')?.addEventListener('click', handlers.onGlossary);
        document.getElementById('define-formatting')?.addEventListener('click', handlers.onFormat);
        document.getElementById('removeFormatting')?.addEventListener('click', handlers.onRemoveFormat);
        document.getElementById('theme-toggle')?.addEventListener('click', handlers.onThemeToggle);
        document.getElementById('logout')?.addEventListener('click', handlers.onLogout);
    }
}
