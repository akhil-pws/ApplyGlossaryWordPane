/**
 * Shared "Reason for Change" popup, same behaviour as the LINK AI web app.
 *
 * Reason for change is configured per module (user role entity name) in LINK AI web › Settings ›
 * Reason for Change. Call requestReasonForChange() before a create / update / delete API call:
 *   - resolves null       → reason not required for the module(s): call the API without a reason
 *   - resolves a string   → the reason the user entered: send it as ReasonForChange
 *   - resolves undefined  → the user cancelled the popup: don't call the API
 *
 * Usage:
 *   const reason = await requestReasonForChange('Approval Log');
 *   if (reason === undefined) return;
 *   await api.save({ ...data, ReasonForChange: reason });
 */
import { StoreService } from "../services/store.service";
import { UIService } from "../services/ui.service";
import { getReasonForChangeModules, ReasonForChangeModule } from "../api/reason-for-change.api";
import { ReasonForChangePopup } from "./bodyelements";

const MAX_LENGTH = 500;

// Modules are fetched once per login and kept in memory; a new login token fetches again
let modulesPromise: Promise<ReasonForChangeModule[]> | null = null;
let cachedToken: string | null = null;

function getModules(): Promise<ReasonForChangeModule[]> {
  const token = StoreService.getInstance().jwt;
  if (!modulesPromise || cachedToken !== token) {
    cachedToken = token;
    modulesPromise = getReasonForChangeModules(token).catch((error) => {
      console.error('Error loading reason for change settings:', error);
      // don't cache failures, the next action tries again
      modulesPromise = null;
      return [];
    });
  }
  return modulesPromise;
}

/** True when any of the given modules is active and requires a reason. */
export async function isReasonForChangeRequired(moduleNames: string | string[]): Promise<boolean> {
  const names = Array.isArray(moduleNames) ? moduleNames : [moduleNames];
  const isCached = !!modulesPromise && cachedToken === StoreService.getInstance().jwt;
  if (!isCached) UIService.toggleLoader(true);
  try {
    const modules = await getModules();
    return modules.some(module =>
      names.includes(module.ModuleName) && module.Status === 1 && module.IsRequired === 1);
  } finally {
    if (!isCached) UIService.toggleLoader(false);
  }
}

/**
 * Asks for a reason when the module(s) require one.
 * Resolves null (not required), the entered reason, or undefined (cancelled).
 */
export async function requestReasonForChange(moduleNames: string | string[]): Promise<string | null | undefined> {
  if (!(await isReasonForChangeRequired(moduleNames))) {
    return null;
  }
  return openReasonForChangePopup();
}

/** Shows the popup; resolves the trimmed reason, or undefined when cancelled. */
export function openReasonForChangePopup(header = 'Reason for Change'): Promise<string | undefined> {
  return new Promise((resolve) => {
    // Hide other open popups meanwhile (e.g. the Approve dialog) so they don't stack
    const hiddenPopups = Array.from(document.querySelectorAll<HTMLElement>('#confirmation-popup > *'))
      .filter(el => el.style.display !== 'none');
    hiddenPopups.forEach(el => (el.style.display = 'none'));

    // Own host, so the popup can sit on top of a dialog rendered in #confirmation-popup
    const host = document.createElement('section');
    host.id = 'reason-for-change-popup';
    host.innerHTML = ReasonForChangePopup(header, MAX_LENGTH);
    document.body.appendChild(host);

    const textarea = host.querySelector('#reason-for-change-input') as HTMLTextAreaElement;
    const error = host.querySelector('#reason-for-change-error') as HTMLElement;
    let submitted = false;

    const validate = (): boolean => {
      const value = textarea.value.trim();
      const message = !value
        ? 'Reason for change is required.'
        : value.length > MAX_LENGTH
          ? `Reason for change cannot exceed ${MAX_LENGTH} characters.`
          : '';
      error.textContent = message;
      error.classList.toggle('d-none', !message);
      error.classList.toggle('d-block', !!message);
      textarea.classList.toggle('is-invalid', !!message);
      return !message;
    };

    const close = (reason: string | undefined) => {
      window.removeEventListener('keydown', onKeydown, true);
      host.remove();
      hiddenPopups.forEach(el => (el.style.display = ''));
      resolve(reason);
    };

    // Escape cancels only this popup, not the dialog underneath
    const onKeydown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopImmediatePropagation();
        close(undefined);
      }
    };
    window.addEventListener('keydown', onKeydown, true);

    textarea.addEventListener('input', () => {
      if (submitted) validate();
    });
    host.querySelector('#reason-for-change-save')!.addEventListener('click', () => {
      submitted = true;
      if (validate()) close(textarea.value.trim());
    });
    host.querySelector('#reason-for-change-cancel')!.addEventListener('click', () => close(undefined));

    textarea.focus();
  });
}
