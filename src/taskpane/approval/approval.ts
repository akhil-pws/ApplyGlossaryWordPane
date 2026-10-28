/**
 * Approve Document.
 *
 * A user with approval rights approves a document from the Tools (wrench) menu. Approving locks the
 * document against AI content: everyone sees a read-only banner, and the AI chat, insert and formatting
 * controls are disabled until a user with approval rights unapproves it. Each approve/unapprove is
 * recorded in the document's Approvals History.
 *
 * State comes from /api/report-approval (status, history, set). The status is cached per document so the
 * lock checks stay synchronous; it is reloaded whenever the header is rendered and after each change.
 */
import { StoreService } from "../services/store.service";
import { UIService } from "../services/ui.service";
import { toaster } from "../components/bodyelements";
import { requestReasonForChange } from "../components/reason-for-change";
import { ApprovalApiAction, ApprovalHistoryData, ApprovalStatusData, getApprovalHistory, getApprovalStatus, setApproval } from "./approval.api";

/**
 * Whether the current user may approve/unapprove. The API has no approval permission yet, so every
 * signed-in user gets the Approve/Unapprove menu item; set to false to preview the no-permission screen.
 */
export const DEMO_CAN_APPROVE = true;
const DIALOG_HISTORY_LIMIT = 5;
/** Reason for Change module (LINK AI web › Settings › Reason for Change) that covers approve / unapprove. */
const REASON_FOR_CHANGE_MODULE = 'Approval Log';
const LOCKED_TITLE = 'Locked while the document is approved';

type ApprovalAction = 'Approved' | 'Unapproved';

interface ApprovalEntry {
  action: ApprovalAction;
  user: string;
  at: string; // ISO date
}

// Controls locked while the document is approved (AI chat, insertion, copy, Action menu)
const LOCKED_CONTROL_SELECTORS = [
  '#action-dropdown-btn',
  '#addNewChatBtn',
  '#chatInput',
  '#insertTagButton',
  '#promptBuilderButton',
  '#sendPromptButton',
  '.response-checkbox',
  '[id^="copyPrompt-"]',
  '[id^="copyResponse-"]',
  '[id^="savePrompt-"]',
];

// Tools menu items disabled while approved
const LOCKED_TOOL_IDS = [
  'predefined-table',
  'customized-table',
  'default-text-style',
  'customized-style',
  'define-formatting',
  'glossary',
  'removeFormatting',
];

const GUARDED_SELECTOR = [
  ...LOCKED_CONTROL_SELECTORS,
  ...LOCKED_TOOL_IDS.map(id => `#${id}`),
  '.approval-insert-locked',
].join(', ');

const ICONS = {
  lock: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="5" y="11" width="14" height="10" rx="2"></rect><path d="M8 11V7a4 4 0 0 1 8 0v4"></path></svg>',
  unlock: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="5" y="11" width="14" height="10" rx="2"></rect><path d="M8 11V7a4 4 0 0 1 7.5-2"></path></svg>',
  shield: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3l8 3v6c0 5-3.5 8-8 9-4.5-1-8-4-8-9V6z"></path><path d="M9 12l2 2 4-4"></path></svg>',
  history: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 12a9 9 0 1 0 3-6.7L3 8"></path><path d="M3 3v5h5"></path><path d="M12 7v5l3 2"></path></svg>',
  warning: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3L2 20h20z"></path><path d="M12 10v4M12 17v.5"></path></svg>',
  chevronDown: '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 9l6 6 6-6"></path></svg>',
  chevronUp: '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 15l6-6 6 6"></path></svg>',
  chevronRight: '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 6l6 6-6 6"></path></svg>',
  close: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"></path></svg>',
};

let lockObserver: MutationObserver | null = null;
let clickGuardInstalled = false;
let lockScheduled = false;

// Cached approval status of the current document; null until loaded (or when loading failed)
let approvalStatus: ApprovalStatusData | null = null;
// Bumped on every status load / logout so a slow response can't overwrite newer state
let statusRequestId = 0;

// ---------------------------------------------------------------------------
// State (API)
// ---------------------------------------------------------------------------

function reportHeadId(): number | null {
  const store = StoreService.getInstance();
  const id = Number(store.dataList?.ID || store.documentID);
  return Number.isFinite(id) && id > 0 ? id : null;
}

/** Loads the current document's approval status into the cache. Leaves it null if it can't be loaded. */
async function loadApprovalStatus(): Promise<void> {
  const store = StoreService.getInstance();
  const id = reportHeadId();
  const requestId = ++statusRequestId;
  if (!id || !store.jwt) {
    approvalStatus = null;
    return;
  }
  try {
    const status = await getApprovalStatus(id, store.jwt);
    if (requestId === statusRequestId) approvalStatus = status;
  } catch (error) {
    console.error('Error loading approval status:', error);
    if (requestId === statusRequestId) approvalStatus = null;
  }
}

/** Approvals History for this document, newest first; null if it couldn't be loaded. */
async function loadApprovalHistory(): Promise<ApprovalEntry[] | null> {
  const store = StoreService.getInstance();
  const id = reportHeadId();
  if (!id || !store.jwt) return null;
  try {
    const rows = await getApprovalHistory(id, store.jwt);
    return rows
      .slice()
      .sort((a: ApprovalHistoryData, b: ApprovalHistoryData) =>
        (Date.parse(b.CreatedDate) - Date.parse(a.CreatedDate)) || (b.ID - a.ID))
      .map(row => ({
        action: row.Action === 'APPROVED' ? 'Approved' : 'Unapproved',
        user: row.CreatedBy || '—',
        at: row.CreatedDate,
      }));
  } catch (error) {
    console.error('Error loading approvals history:', error);
    return null;
  }
}

/** Approves or unapproves the current document; returns false (after telling the user) if it failed. */
async function submitApproval(action: ApprovalApiAction, reasonForChange: string | null): Promise<boolean> {
  const store = StoreService.getInstance();
  const id = reportHeadId();
  if (!id || !store.jwt) {
    toaster('This document is not linked to a LINK AI report', 'error');
    return false;
  }
  try {
    approvalStatus = await setApproval(id, action, store.jwt, reasonForChange);
    statusRequestId++;
    return true;
  } catch (error) {
    console.error('Error updating approval:', error);
    const message = (error as Error)?.message;
    toaster(message && message !== 'Network response was not ok'
      ? message
      : 'Could not update the approval. Please try again.', 'error');
    return false;
  }
}

export function isDocumentApproved(): boolean {
  const store = StoreService.getInstance();
  return !!store.jwt && !!approvalStatus?.IsApproved;
}

/** Whether the current user has approval rights (no permission in the API yet: DEMO_CAN_APPROVE). */
export function canApprove(): boolean {
  return DEMO_CAN_APPROVE;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function esc(value: unknown): string {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatWhen(iso: string): string {
  const d = new Date(iso);
  if (!iso || isNaN(d.getTime())) return iso || '—';
  const day = String(d.getDate()).padStart(2, '0');
  const month = d.toLocaleString('en-US', { month: 'short' });
  const time = `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
  return `${day} ${month} ${d.getFullYear()}, ${time}`;
}

/** Claims from the login token (the user's name isn't stored anywhere else). */
function tokenClaims(): Record<string, any> {
  try {
    const payload = StoreService.getInstance().jwt.split('.')[1];
    const base64 = payload.replace(/-/g, '+').replace(/_/g, '/');
    const binary = atob(base64.padEnd(Math.ceil(base64.length / 4) * 4, '='));
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return JSON.parse(new TextDecoder().decode(bytes)) || {};
  } catch {
    return {};
  }
}

function currentUserName(): string {
  const claims = tokenClaims();
  const role = StoreService.getInstance().UserRole || {};
  const fullName = [claims.given_name, claims.family_name].filter(Boolean).join(' ');
  return claims.name
    || fullName
    || claims['http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name']
    || claims.unique_name
    || role.FullName || role.UserFullName || role.UserName || role.Username
    || 'Current user';
}

/** Name of the open Word file, used when the API doesn't return a document name. */
function wordFileName(): string {
  try {
    const url = Office?.context?.document?.url;
    if (url) {
      const name = decodeURIComponent(url.split(/[\\/]/).pop() || '');
      if (name) return name;
    }
  } catch {
    // Office not available (e.g. browser preview)
  }
  return 'This document';
}

function documentName(status: ApprovalStatusData | null = approvalStatus): string {
  return status?.DocumentName || wordFileName();
}

// ---------------------------------------------------------------------------
// Public lifecycle
// ---------------------------------------------------------------------------

/**
 * Call after the header (logoheader) has been rendered. Safe to call repeatedly: re-renders with the
 * cached state straight away, then reloads the status from the API and re-renders.
 */
export async function initApproval(): Promise<void> {
  installClickGuard();
  installLockObserver();
  refreshApprovalUI();
  await loadApprovalStatus();
  refreshApprovalUI();
}

/** Re-renders the banner and Tools menu and applies/releases the lock for the current state. */
export function refreshApprovalUI() {
  const locked = isDocumentApproved();
  renderBanner(locked);
  renderToolsMenu(locked);
  if (locked) {
    applyLock();
  } else {
    releaseLock();
  }
  refreshTagSuggestions();
}

/** Call on logout: removes approval UI, releases any lock and forgets the cached status. */
export function removeApprovalUI() {
  approvalStatus = null;
  statusRequestId++;
  document.getElementById('approval-banner')?.remove();
  closeDialog();
  releaseLock();
  lockObserver?.disconnect();
  lockObserver = null;
}

// ---------------------------------------------------------------------------
// Read-only banner (screens 3–8)
// ---------------------------------------------------------------------------

function renderBanner(locked: boolean) {
  const existing = document.getElementById('approval-banner');
  if (!locked) {
    existing?.remove();
    return;
  }
  if (existing) return;

  const banner = document.createElement('div');
  banner.id = 'approval-banner';
  banner.className = 'approval-banner';
  banner.setAttribute('role', 'status');
  banner.innerHTML = `
    <span class="approval-banner-icon">${ICONS.lock}</span>
    <div class="approval-banner-body">
      <p class="approval-banner-text"><strong>Read-only.</strong> This document has been approved and is locked from AI content. Only users with approval rights can unapprove this document.</p>
      <button type="button" class="approval-link" id="approval-history-link">${ICONS.history}Approvals History</button>
    </div>`;

  const headerNav = document.getElementById('header-nav');
  if (headerNav) {
    headerNav.insertAdjacentElement('afterend', banner);
  } else {
    document.body.prepend(banner);
  }
  banner.querySelector('#approval-history-link')?.addEventListener('click', () => openHistoryDialog());
}

// ---------------------------------------------------------------------------
// Tools (wrench) menu (screens 1, 5, 8)
// ---------------------------------------------------------------------------

function renderToolsMenu(locked: boolean) {
  const menu = document.querySelector('ul[aria-labelledby="settingsDropdown"]');
  if (!menu) return;

  menu.querySelectorAll('[data-approval-menu]').forEach(el => el.remove());

  LOCKED_TOOL_IDS.forEach(id => {
    const item = document.getElementById(id);
    if (!item) return;
    item.classList.toggle('approval-disabled', locked);
    if (locked) {
      item.setAttribute('aria-disabled', 'true');
      item.setAttribute('title', LOCKED_TITLE);
    } else {
      item.removeAttribute('title');
      // glossary / removeFormatting manage their own disabled state through .disabled-link
      if (!item.classList.contains('disabled-link')) item.removeAttribute('aria-disabled');
    }
  });

  if (locked) {
    const header = document.createElement('li');
    header.setAttribute('data-approval-menu', '');
    header.className = 'approval-menu-header';
    header.innerHTML = `${ICONS.lock}Locked while approved`;
    menu.prepend(header);
  }

  // No Approve/Unapprove until the status is known (still loading, or it failed to load)
  if (!canApprove() || !approvalStatus) return;

  const divider = document.createElement('li');
  divider.setAttribute('data-approval-menu', '');
  divider.innerHTML = '<hr class="dropdown-divider">';

  const item = document.createElement('li');
  item.setAttribute('data-approval-menu', '');
  item.innerHTML = locked
    ? `<a class="dropdown-item approval-menu-item unapprove" href="#" id="unapprove-document">${ICONS.unlock}Unapprove</a>`
    : `<a class="dropdown-item approval-menu-item approve" href="#" id="approve-document">${ICONS.shield}Approve</a>`;

  menu.append(divider, item);

  item.querySelector('a')?.addEventListener('click', (e) => {
    e.preventDefault();
    if (locked) {
      openUnapproveDialog();
    } else {
      openApproveDialog();
    }
  });
}

// ---------------------------------------------------------------------------
// Locking of AI / insert controls
// ---------------------------------------------------------------------------

function installClickGuard() {
  if (clickGuardInstalled) return;
  clickGuardInstalled = true;

  // Capture phase so it runs before the controls' own listeners (and Bootstrap's dropdown toggles)
  document.addEventListener('click', (e) => {
    if (!isDocumentApproved()) return;
    const target = (e.target as HTMLElement)?.closest?.(GUARDED_SELECTOR) as HTMLElement | null;
    if (!target) return;
    e.preventDefault();
    e.stopPropagation();
    e.stopImmediatePropagation();
    if (target.classList.contains('approval-insert-locked')) {
      toaster('Insert is locked while the document is approved', 'error');
    }
  }, true);
}

function installLockObserver() {
  if (lockObserver) return;
  const appBody = document.getElementById('app-body');
  if (!appBody) return;
  // Views are re-rendered with innerHTML, so re-apply the lock whenever #app-body content changes
  lockObserver = new MutationObserver(() => {
    if (lockScheduled || !isDocumentApproved()) return;
    lockScheduled = true;
    requestAnimationFrame(() => {
      lockScheduled = false;
      if (isDocumentApproved()) applyLock();
    });
  });
  lockObserver.observe(appBody, { childList: true, subtree: true });
}

function applyLock() {
  document.querySelectorAll<HTMLElement>(LOCKED_CONTROL_SELECTORS.join(', ')).forEach(el => {
    if (el.hasAttribute('data-approval-locked')) return;
    el.setAttribute('data-approval-locked', '');
    el.classList.add('approval-locked-control');
    el.setAttribute('aria-disabled', 'true');

    if (el.hasAttribute('title')) el.setAttribute('data-approval-prev-title', el.getAttribute('title'));
    el.setAttribute('title', LOCKED_TITLE);

    if (el instanceof HTMLButtonElement || el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) {
      if (el.disabled) el.setAttribute('data-approval-was-disabled', '');
      el.disabled = true;
    }
    if (el instanceof HTMLTextAreaElement) {
      el.setAttribute('data-approval-prev-placeholder', el.placeholder || '');
      el.placeholder = 'Chat is read-only while locked';
    }
    if (el.id === 'action-dropdown-btn') {
      el.insertAdjacentHTML('afterbegin', `<span class="approval-lock-icon">${ICONS.lock}</span>`);
    }
  });
}

function releaseLock() {
  document.querySelectorAll<HTMLElement>('[data-approval-locked]').forEach(el => {
    el.removeAttribute('data-approval-locked');
    el.classList.remove('approval-locked-control');
    el.removeAttribute('aria-disabled');

    const prevTitle = el.getAttribute('data-approval-prev-title');
    if (prevTitle !== null) {
      el.setAttribute('title', prevTitle);
      el.removeAttribute('data-approval-prev-title');
    } else {
      el.removeAttribute('title');
    }

    if (el instanceof HTMLButtonElement || el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) {
      el.disabled = el.hasAttribute('data-approval-was-disabled');
      el.removeAttribute('data-approval-was-disabled');
    }
    if (el instanceof HTMLTextAreaElement) {
      el.placeholder = el.getAttribute('data-approval-prev-placeholder') || '';
      el.removeAttribute('data-approval-prev-placeholder');
    }
    el.querySelector(':scope > .approval-lock-icon')?.remove();
  });
}

/** Re-runs the home tag search so Properties / AI Tags pick up the "Insert locked" / "View only" notes. */
function refreshTagSuggestions() {
  const searchBox = document.getElementById('search-box') as HTMLInputElement | null;
  if (searchBox && searchBox.value.trim() && document.getElementById('suggestion-list')) {
    searchBox.dispatchEvent(new Event('input'));
  }
}

/** Note shown on home search results while approved: Properties/Images can't be inserted, AI tags are view only. */
export function approvalTagNoteHtml(isAITag: boolean): string {
  if (!isDocumentApproved()) return '';
  return isAITag
    ? '<span class="approval-item-note">View only</span>'
    : `<span class="approval-item-note">${ICONS.lock}Insert locked</span>`;
}

// ---------------------------------------------------------------------------
// Dialogs (screens 2, 2b, 6, 7)
// ---------------------------------------------------------------------------

function closeDialog() {
  document.getElementById('approval-modal')?.remove();
  document.removeEventListener('keydown', onDialogKeydown);
}

function onDialogKeydown(e: KeyboardEvent) {
  if (e.key === 'Escape') closeDialog();
}

function showDialog(html: string): HTMLElement {
  closeDialog();
  const host = document.getElementById('confirmation-popup') || document.body;
  const wrapper = document.createElement('div');
  wrapper.id = 'approval-modal';
  wrapper.className = 'approval-modal-backdrop';
  wrapper.innerHTML = html;
  host.appendChild(wrapper);
  document.addEventListener('keydown', onDialogKeydown);
  // Focus the dialog itself so keyboard users start inside it without a focus ring on a button
  (wrapper.querySelector('.approval-dialog') as HTMLElement | null)?.focus();
  return wrapper;
}

/** Runs the API loads a dialog needs behind the page loader. */
async function withLoader<T>(load: () => Promise<T>): Promise<T> {
  UIService.toggleLoader(true);
  try {
    return await load();
  } finally {
    UIService.toggleLoader(false);
  }
}

/** Fresh status + history for a dialog; falls back to the cached status if the status call fails. */
async function loadDialogData(): Promise<{ status: ApprovalStatusData | null, history: ApprovalEntry[] | null }> {
  const previous = approvalStatus;
  const [, history] = await Promise.all([loadApprovalStatus(), loadApprovalHistory()]);
  if (!approvalStatus) approvalStatus = previous;
  return { status: approvalStatus, history };
}

function historyRowsHtml(entries: ApprovalEntry[]): string {
  if (entries.length === 0) {
    return '<li class="approval-history-empty">No approvals recorded yet.</li>';
  }
  return entries.map(entry => `
    <li class="approval-history-row">
      <span class="approval-history-action ${entry.action === 'Approved' ? 'approved' : 'unapproved'}">${entry.action}</span>
      <span class="approval-history-user" title="${esc(entry.user)}">${esc(entry.user)}</span>
      <span class="approval-history-when">${formatWhen(entry.at)}</span>
    </li>`).join('');
}

function historyTableHtml(entries: ApprovalEntry[] | null, label: string): string {
  if (!entries) {
    return '<p class="approval-history-empty approval-history-error">Couldn\'t load the approvals history. Please try again.</p>';
  }
  return `
    <ol class="approval-history-list" aria-label="${label}">
      <li class="approval-history-row approval-history-head"><span>Action</span><span>User</span><span>Date &amp; time</span></li>
      ${historyRowsHtml(entries)}
    </ol>`;
}

/** Collapsible "Approvals History (n)" section used by the approve / unapprove dialogs. */
function dialogHistoryHtml(history: ApprovalEntry[] | null, open: boolean): string {
  if (!history) {
    return `<div class="approval-dialog-history">${historyTableHtml(null, '')}</div>`;
  }
  const shown = history.slice(0, DIALOG_HISTORY_LIMIT);
  return `
    <div class="approval-dialog-history">
      <button type="button" class="approval-link" id="approval-history-toggle" aria-expanded="${open}" aria-controls="approval-history-panel">
        ${ICONS.history}Approvals History (${history.length})<span class="approval-caret">${ICONS.chevronDown}${ICONS.chevronUp}</span>
      </button>
      <div id="approval-history-panel" class="approval-history-panel" ${open ? '' : 'hidden'}>
        ${historyTableHtml(shown, 'Latest approvals, newest first')}
        <div class="approval-history-footer">
          <span>Latest ${shown.length} of ${history.length}</span>
          <button type="button" class="approval-link" id="approval-view-full-history">View full history${ICONS.chevronRight}</button>
        </div>
      </div>
    </div>`;
}

function wireDialogHistory(dialog: HTMLElement, history: ApprovalEntry[] | null, reopen: (historyOpen: boolean) => void) {
  const toggle = dialog.querySelector('#approval-history-toggle') as HTMLButtonElement;
  const panel = dialog.querySelector('#approval-history-panel') as HTMLElement;
  toggle?.addEventListener('click', () => {
    const open = panel.hidden;
    panel.hidden = !open;
    toggle.setAttribute('aria-expanded', String(open));
  });
  dialog.querySelector('#approval-view-full-history')?.addEventListener('click', () => {
    // Closing the full history returns to this dialog
    openHistoryDialog(() => reopen(true), history);
  });
}

/** Disables the dialog buttons and shows progress on the confirm button while a change is saved. */
function setDialogBusy(dialog: HTMLElement, busy: boolean, busyLabel: string) {
  const confirmBtn = dialog.querySelector('#approval-confirm') as HTMLButtonElement;
  const cancelBtn = dialog.querySelector('#approval-cancel') as HTMLButtonElement;
  if (busy) {
    confirmBtn.dataset.label = confirmBtn.innerHTML;
    confirmBtn.innerHTML = `<span class="spinner-border spinner-border-sm" aria-hidden="true"></span>${busyLabel}`;
  } else if (confirmBtn.dataset.label) {
    confirmBtn.innerHTML = confirmBtn.dataset.label;
  }
  confirmBtn.disabled = busy;
  cancelBtn.disabled = busy;
}

interface ApproveDialogOptions {
  checked?: boolean;
  historyOpen?: boolean;
  history?: ApprovalEntry[] | null;
}

export async function openApproveDialog(options: ApproveDialogOptions = {}) {
  const store = StoreService.getInstance();
  if (store.isPendingResponse) {
    toaster('Wait for the current AI response to finish before approving', 'error');
    return;
  }
  const { checked = false, historyOpen = false } = options;
  const { status, history } = options.history !== undefined
    ? { status: approvalStatus, history: options.history }
    : await withLoader(loadDialogData);

  if (status?.IsApproved) {
    // Approved meanwhile (e.g. by another user): show the locked state instead
    refreshApprovalUI();
    toaster('This document has already been approved', 'error');
    return;
  }

  const dialog = showDialog(`
    <section class="approval-dialog" role="dialog" aria-modal="true" tabindex="-1" aria-labelledby="approval-dialog-title">
      <div class="approval-dialog-head">
        <span class="approval-dialog-icon approve">${ICONS.lock}</span>
        <h2 id="approval-dialog-title">Approve and lock this document?</h2>
      </div>

      <div class="approval-details">
        <span>Document</span><span class="approval-doc-name">${esc(documentName(status))}</span>
        <span>Type</span><span>${esc(status?.DocumentType || '—')}</span>
        <span>Template</span><span>${esc(status?.DocumentTemplate || '—')}</span>
        <span>Approver</span><span>${esc(currentUserName())}</span>
      </div>

      <p class="approval-note info">${ICONS.shield}<span>By approving, you certify that you have fully reviewed this document and have the expertise to do so. LINK AI will lock the document against AI queries and text insertion. Your name, date and time are recorded.</span></p>

      <label class="approval-certify">
        <input type="checkbox" id="approval-certify-checkbox" ${checked ? 'checked' : ''}>
        <span>I have fully reviewed this document and have the expertise to approve it.</span>
      </label>

      ${dialogHistoryHtml(history, historyOpen)}

      <div class="approval-dialog-actions">
        <button type="button" class="approval-btn secondary" id="approval-cancel">Cancel</button>
        <button type="button" class="approval-btn approve" id="approval-confirm" ${checked ? '' : 'disabled'} title="${checked ? '' : 'Confirm the review statement to enable'}">${ICONS.lock}Approve &amp; Lock</button>
      </div>
    </section>`);

  const checkbox = dialog.querySelector('#approval-certify-checkbox') as HTMLInputElement;
  const confirmBtn = dialog.querySelector('#approval-confirm') as HTMLButtonElement;

  checkbox.addEventListener('change', () => {
    confirmBtn.disabled = !checkbox.checked;
    confirmBtn.title = checkbox.checked ? '' : 'Confirm the review statement to enable';
  });
  wireDialogHistory(dialog, history, (open) => openApproveDialog({ checked: checkbox.checked, historyOpen: open, history }));
  dialog.querySelector('#approval-cancel')?.addEventListener('click', closeDialog);
  confirmBtn.addEventListener('click', async () => {
    if (!checkbox.checked) return;
    // Ask for a reason first when Reason for Change is enabled; cancelling keeps this dialog open
    confirmBtn.disabled = true;
    const reason = await requestReasonForChange(REASON_FOR_CHANGE_MODULE);
    if (reason === undefined) {
      confirmBtn.disabled = !checkbox.checked;
      return;
    }
    setDialogBusy(dialog, true, 'Approving…');
    checkbox.disabled = true;
    if (await submitApproval('APPROVED', reason)) {
      closeDialog();
      refreshApprovalUI();
      toaster('Document approved and locked', 'success');
    } else {
      setDialogBusy(dialog, false, '');
      checkbox.disabled = false;
    }
  });
}

export async function openUnapproveDialog(options: { historyOpen?: boolean, history?: ApprovalEntry[] | null } = {}) {
  const { historyOpen = false } = options;
  const { status, history } = options.history !== undefined
    ? { status: approvalStatus, history: options.history }
    : await withLoader(loadDialogData);

  if (status && !status.IsApproved) {
    // Unapproved meanwhile (e.g. by another user): show the unlocked state instead
    refreshApprovalUI();
    toaster('This document is no longer approved', 'error');
    return;
  }

  // The status records who made the current (approved) state; fall back to the latest approval in the history
  const lastApproval = history?.find(entry => entry.action === 'Approved');
  const approvedBy = status?.IsApproved ? status.CreatedBy : lastApproval?.user;
  const approvedOn = status?.IsApproved ? status.CreatedDate : lastApproval?.at;

  const dialog = showDialog(`
    <section class="approval-dialog" role="dialog" aria-modal="true" tabindex="-1" aria-labelledby="approval-dialog-title">
      <div class="approval-dialog-head">
        <span class="approval-dialog-icon unapprove">${ICONS.unlock}</span>
        <h2 id="approval-dialog-title">Unapprove and unlock this document?</h2>
      </div>

      <div class="approval-details wide">
        <span>Document</span><span class="approval-doc-name">${esc(documentName(status))}</span>
        <span>Approved by</span><span>${esc(approvedBy || '—')}</span>
        <span>Approved on</span><span>${approvedOn ? formatWhen(approvedOn) : '—'}</span>
      </div>

      <p class="approval-note warning">${ICONS.warning}<span>This removes the human-review certification and unlocks the document for AI queries and text insertion. It should be reviewed and approved again before it is treated as fully human-reviewed. Your name, date and time are recorded.</span></p>

      ${dialogHistoryHtml(history, historyOpen)}

      <div class="approval-dialog-actions">
        <button type="button" class="approval-btn secondary" id="approval-cancel">Cancel</button>
        <button type="button" class="approval-btn unapprove" id="approval-confirm">${ICONS.unlock}Unapprove &amp; Unlock</button>
      </div>
    </section>`);

  wireDialogHistory(dialog, history, (open) => openUnapproveDialog({ historyOpen: open, history }));
  dialog.querySelector('#approval-cancel')?.addEventListener('click', closeDialog);
  const confirmBtn = dialog.querySelector('#approval-confirm') as HTMLButtonElement;
  confirmBtn.addEventListener('click', async () => {
    // Ask for a reason first when Reason for Change is enabled; cancelling keeps this dialog open
    confirmBtn.disabled = true;
    const reason = await requestReasonForChange(REASON_FOR_CHANGE_MODULE);
    if (reason === undefined) {
      confirmBtn.disabled = false;
      return;
    }
    setDialogBusy(dialog, true, 'Unapproving…');
    if (await submitApproval('UNAPPROVED', reason)) {
      closeDialog();
      refreshApprovalUI();
      toaster('Document unapproved and unlocked', 'success');
    } else {
      setDialogBusy(dialog, false, '');
    }
  });
}

/** Full Approvals History for this document, newest first (screen 7). Pass history to reuse what a dialog loaded. */
export async function openHistoryDialog(onClose?: () => void, loadedHistory?: ApprovalEntry[] | null) {
  const history = loadedHistory !== undefined ? loadedHistory : await withLoader(loadApprovalHistory);
  const close = () => {
    closeDialog();
    onClose?.();
  };
  const count = history ? `${history.length} ${history.length === 1 ? 'entry' : 'entries'}` : '';

  const dialog = showDialog(`
    <section class="approval-dialog" role="dialog" aria-modal="true" tabindex="-1" aria-labelledby="approval-dialog-title">
      <div class="approval-dialog-head">
        <span class="approval-dialog-icon history">${ICONS.history}</span>
        <div class="approval-dialog-titles">
          <h2 id="approval-dialog-title">Approvals History</h2>
          <span class="approval-dialog-subtitle">${esc(documentName())}${count ? ` · ${count}` : ''}</span>
        </div>
        <button type="button" class="approval-icon-btn" id="approval-close-x" aria-label="Close">${ICONS.close}</button>
      </div>

      <div class="approval-history-scroll">
        ${historyTableHtml(history, 'Approvals history, latest first')}
      </div>

      <p class="approval-footnote">The log across all documents is in LINK AI web › Reports › Approval Log.</p>

      <div class="approval-dialog-actions">
        <button type="button" class="approval-btn secondary" id="approval-cancel">Close</button>
      </div>
    </section>`);

  dialog.querySelector('#approval-close-x')?.addEventListener('click', close);
  dialog.querySelector('#approval-cancel')?.addEventListener('click', close);
}
