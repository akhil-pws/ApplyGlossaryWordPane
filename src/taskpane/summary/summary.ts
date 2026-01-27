import { generateCheckboxHistory } from "../draft/home";
import { addGenAITags } from "../taskpane";
import {
  activateSummaryMode,
  getSummaryTagsByReportHeadId,
  getSummaryTagHistory,
  getSummaryTagStatus,
  refreshSummaryMode,
  addSummaryHistory
} from "./summary.api";
import { CONFIG } from "../utils/config";
import { UIService } from "../services/ui.service";
import { StoreService } from "../services/store.service";
import { Confirmationpopup, toaster } from "../components/bodyelements";

export var summarySelectedNames: string[] = [];

let isSummaryLoading = false;
let allSummaryTags: any[] = [];
let currentSummaryStatus = 0;
let currentSummaryInstance = 0;

export async function loadSummarypage(availableKeys: any[]) {
  const instanceId = ++currentSummaryInstance;
  const store = StoreService.getInstance();
  const searchBoxClass = store.theme === 'Dark' ? 'bg-secondary text-light' : 'bg-white text-dark';

  document.getElementById('app-body').innerHTML = `
    <div class="container pt-3">

      <!-- Top bar -->
      <div class="d-flex justify-content-end px-2">
        <div class="dropdown">
          <button class="btn btn-default dropdown-toggle" type="button" data-bs-toggle="dropdown" aria-expanded="false">
            Action
          </button>
          <ul class="dropdown-menu dropdown-menu-end">
            <li>
              <a class="dropdown-item" href="#" id="add-btn-tag">
                <i class="fa-solid fa-plus me-2"></i> Add
              </a>
            </li>

            <li>
              <a class="dropdown-item" href="#" id="reanalyze-draft">
                <i class="fa-solid fa-rotate-right me-2"></i> Re-analyze Draft
              </a>
            </li>
          </ul>
        </div>
      </div>

      <!-- Search -->
      <div class="form-group px-2 pt-2">
        <div class="input-group">
          <input type="text" id="search-box"
            class="form-control ${searchBoxClass}"
            placeholder="Search Summary Tags ..."
            autocomplete="off" />
          <span class="input-group-text">
            <i class="fa-solid fa-magnifying-glass text-muted"></i>
          </span>
        </div>
      </div>

      <!-- Card Panel (always visible now) -->
      <div class="card mt-3 mx-2 mb-3" id="summary-card">
        <div class="card-header fw-semibold">Summary Tags</div>

        <div class="list-group list-group-summary list-group-flush" id="summary-tag-list">
          <div class="d-flex justify-content-center align-items-center p-5">
            <div class="loader"></div>
          </div>
        </div>

        <!-- Pagination Footer -->
        <div class="card-footer d-flex justify-content-between align-items-center flex-wrap gap-2">
          <div class="btn-group btn-group-sm" role="group" aria-label="pagination">
            <button class="btn btn-outline-secondary" id="page-first" title="First">
              <i class="fa-solid fa-backward-fast"></i>
            </button>
            <button class="btn btn-outline-secondary" id="page-prev" title="Previous">
              <i class="fa-solid fa-backward-step"></i>
            </button>

            <div class="btn-group btn-group-sm" id="page-buttons"></div>

            <button class="btn btn-outline-secondary" id="page-next" title="Next">
              <i class="fa-solid fa-forward-step"></i>
            </button>
            <button class="btn btn-outline-secondary" id="page-last" title="Last">
              <i class="fa-solid fa-forward-fast"></i>
            </button>
          </div>

          <div class="text-muted small" id="page-count-label"></div>
        </div>
      </div>
    </div>
  `;

  const searchBox = document.getElementById('search-box') as HTMLInputElement;
  const list = document.getElementById('summary-tag-list') as HTMLElement;
  const pageButtons = document.getElementById('page-buttons') as HTMLElement;
  const pageCountLabel = document.getElementById('page-count-label') as HTMLElement;

  const btnFirst = document.getElementById('page-first') as HTMLButtonElement;
  const btnPrev = document.getElementById('page-prev') as HTMLButtonElement;
  const btnNext = document.getElementById('page-next') as HTMLButtonElement;
  const btnLast = document.getElementById('page-last') as HTMLButtonElement;

  const addBtn = document.getElementById('add-btn-tag') as HTMLAnchorElement;
  const reanalyzeBtn = document.getElementById('reanalyze-draft') as HTMLAnchorElement;

  function disableActionButtons(disabled: boolean) {
    if (addBtn) addBtn.classList.toggle("disabled", disabled);
  }

  function setReanalyzeButtonState(enabled: boolean) {
    if (reanalyzeBtn) reanalyzeBtn.classList.toggle("disabled", !enabled);
  }

  // ✅ helper to normalize API response into string[]
  function normalizeNames(res: any): string[] {
    let names =
      res?.Data ??
      res?.SelectedNames ??
      res?.selectedNames ??
      res?.data ??
      res ??
      [];

    if (typeof names === "string") {
      return names.split(",").map((x: string) => x.trim()).filter(Boolean);
    }

    if (Array.isArray(names) && names.length > 0 && typeof names[0] === "object") {
      return names
        .map((x: any) => x.DisplayName || x.Name || x.TagName)
        .filter(Boolean);
    }

    return Array.isArray(names) ? names : [];
  }

  // ✅ Pagination + filter setup
  const pageSize = 8;
  let currentPage = 1;
  let filtered = allSummaryTags;

  function getTotalPages() {
    return Math.max(1, Math.ceil(filtered.length / pageSize));
  }

  function slicePage() {
    const start = (currentPage - 1) * pageSize;
    return filtered.slice(start, start + pageSize);
  }

  function renderRows() {
    list.replaceChildren();

    const pageItems = slicePage();

    if (pageItems.length === 0) {
      if (isSummaryLoading) {
        list.innerHTML = `<div class="d-flex justify-content-center align-items-center p-5"><div class="loader"></div></div>`;
      } else {
        list.innerHTML = `<div class="p-3 text-muted">No tags found</div>`;
      }
      return;
    }

    pageItems.forEach(tag => {
      const row = document.createElement('button');
      row.type = 'button';

      const isActive = currentSummaryStatus === 2;
      const themeClasses = store.theme === 'Dark'
        ? `bg-dark text-light ${isActive ? 'list-hover-dark' : 'opacity-50'}`
        : `bg-light text-dark ${isActive ? 'list-hover-light' : 'opacity-50'}`;

      row.className =
        `list-group-item list-group-item-action d-flex justify-content-between align-items-center ${themeClasses}`;

      if (!isActive) {
        row.disabled = true;
        row.style.cursor = 'not-allowed';
      }

      const tagStatus = (tag.Status === undefined || tag.Status === null) ? "1" : String(tag.Status);

      let statusIcon = "";
      if (tagStatus === "0") {
        statusIcon = `<i class="fa fa-spinner fa-spin text-muted"></i>`;
      } else if (tagStatus === "2") {
        statusIcon = `<i class="fa-solid fa-circle-info text-warning c-pointer" id="reprocess-${tag.ID || tag.ReportHeadSummaryTagID}" title="Click to Reprocess"></i>`;
      } else {
        // Default or status 1
        statusIcon = `<i class="fa-solid fa-circle-check light-navy-blue"></i>`;
      }

      row.innerHTML = `
        <div class="text-truncate pe-2">${tag.Name}</div>
        <div class="d-flex align-items-center gap-3">
          ${statusIcon}
          <i class="fa-solid fa-angles-right text-muted"></i>
        </div>
      `;

      if (isActive) {
        row.onclick = async () => {
          try {
            const appBody = document.getElementById('app-body');
            appBody.innerHTML = `
            <div id="button-container">
              <div class="loader" id="loader"></div>
            </div>
          `;
            const html = await generateCheckboxHistory(tag, "Summary");
            appBody.innerHTML = html;
          } catch {
            document.getElementById('app-body').innerHTML =
              '<div class="text-danger p-2">Error loading data</div>';
          }
        };
      }

      list.appendChild(row);

      if (tagStatus === "2") {
        const infoBtn = row.querySelector(`#reprocess-${tag.ID || tag.ReportHeadSummaryTagID}`);
        if (infoBtn) {
          infoBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            reprocessTag(tag);
          });
        }
      }
    });
  }

  function renderPager() {
    const totalPages = getTotalPages();
    currentPage = Math.min(Math.max(1, currentPage), totalPages);

    const startItem = filtered.length === 0 ? 0 : (currentPage - 1) * pageSize + 1;
    const endItem = Math.min(currentPage * pageSize, filtered.length);
    pageCountLabel.textContent = `${startItem} - ${endItem} of ${filtered.length} items`;

    btnFirst.toggleAttribute('disabled', currentPage === 1);
    btnPrev.toggleAttribute('disabled', currentPage === 1);
    btnNext.toggleAttribute('disabled', currentPage === totalPages);
    btnLast.toggleAttribute('disabled', currentPage === totalPages);

    pageButtons.replaceChildren();

    const maxButtons = 5;
    let start = Math.max(1, currentPage - Math.floor(maxButtons / 2));
    let end = Math.min(totalPages, start + maxButtons - 1);
    start = Math.max(1, end - maxButtons + 1);

    for (let p = start; p <= end; p++) {
      const b = document.createElement('button');
      b.className = `btn ${p === currentPage ? 'btn-primary' : 'btn-outline-secondary'}`;
      b.textContent = String(p);
      b.onclick = () => {
        currentPage = p;
        renderAll();
      };
      pageButtons.appendChild(b);
    }
  }

  function renderAll() {
    renderRows();
    renderPager();
  }

  function applySearch() {
    const term = searchBox.value.trim().toLowerCase();
    filtered = term
      ? allSummaryTags.filter(t => (t.Name || '').toLowerCase().includes(term))
      : allSummaryTags;

    currentPage = 1;
    renderAll();
  }

  let debounceTimeout: any;
  searchBox.addEventListener('input', () => {
    clearTimeout(debounceTimeout);
    debounceTimeout = setTimeout(applySearch, 250);
  });

  btnFirst.addEventListener('click', () => { currentPage = 1; renderAll(); });
  btnPrev.addEventListener('click', () => { currentPage = Math.max(1, currentPage - 1); renderAll(); });
  btnNext.addEventListener('click', () => { currentPage = Math.min(getTotalPages(), currentPage + 1); renderAll(); });
  btnLast.addEventListener('click', () => { currentPage = getTotalPages(); renderAll(); });

  // ✅ Action button wiring
  addBtn?.addEventListener('click', () => {
    if (!isSummaryLoading) addGenAITags();
  });

  // --------------------------
  // ✅ NEW FLOW: First call GET tags API and read SummaryTagGenerated from it
  // --------------------------
  let hasActivated = false;
  function deduplicateSummarySources(sources: any[]) {
    const map = new Map();
    sources.forEach(source => {
      const key = source.FileName;
      if (!map.has(key)) {
        map.set(key, source);
      } else {
        const existing = map.get(key);
        if (String(existing.Status) !== "1" && String(source.Status) === "1") {
          map.set(key, source);
        }
      }
    });
    return Array.from(map.values());
  }

  async function firstLoadAndRender() {
    isSummaryLoading = true;
    disableActionButtons(true);

    try {
      setReanalyzeButtonState(false);

      // ✅ 1) FIRST call GET tags API
      const getRes = await getSummaryTagsByReportHeadId(store.documentID, store.jwt);

      // ✅ 2) check Data.SummaryTagGenerated from GET API response
      currentSummaryStatus = getRes?.Data?.SummaryTagGenerated;

      // ✅ 3) Render tags immediately (no loader/table hide logic)
      allSummaryTags = getRes?.Data?.SummaryTags || [];
      store.sourceSummaryList = deduplicateSummarySources(getRes?.Data?.SummarySources || []);
      filtered = allSummaryTags;
      renderAll();

      const summaryStatus = currentSummaryStatus;

      if (summaryStatus === 2) {
        setReanalyzeButtonState(true);

        if (allSummaryTags && allSummaryTags.length > 0) {
          summarySelectedNames = normalizeNames(allSummaryTags);
        }
        return;
      }

      // status 0 -> activate once, then poll until 2
      if (summaryStatus === 0) {
        if (!hasActivated) {
          hasActivated = true;

          const base64Data = await getWordAsBase64();
          const payload = {
            ReportHeadID: store.documentID,
            ActiveDocument: base64Data
          };

          await activateSummaryMode(payload, store.jwt);
        }

        await pollSummaryUntilDone();
      }

      // status 1 -> poll until 2
      if (summaryStatus === 1) {
        await pollSummaryUntilDone();
      }

    } catch (err) {
      console.error("Summary load failed:", err);
      if (list) list.innerHTML = `<div class="p-3 text-danger">Failed to load Summary mode</div>`;
    } finally {
      isSummaryLoading = false;
      disableActionButtons(false);
    }
  }

  // ✅ Reprocess a single tag
  async function reprocessTag(tag: any) {
    try {
      tag.Status = "0"; // Show spinner immediately
      renderAll();

      const jwt = store.jwt;
      const historyRes = await getSummaryTagHistory(tag.ID || tag.ReportHeadSummaryTagID, jwt);
      const history = historyRes?.Data || [];

      if (history.length === 0) {
        toaster("No history found to reprocess", "error");
        tag.Status = "2";
        renderAll();
        return;
      }

      // Latest history is usually index 0 after unshift, but let's check or just take first in raw list
      const lastHistory = history[0];
      debugger
      const payload = {
        ReportHeadID: Number(store.documentID),
        ReportHeadSummaryTagID: tag.ID || tag.ReportHeadSummaryTagID,
        Prompt: lastHistory.Prompt,
        Response: "",
        Selected: 1,
        SourceVector: lastHistory.SourceVector ? lastHistory.SourceVector : '',
        Name: tag.Name
      };

      await addSummaryHistory(payload, jwt);

      // Start polling if not already running (status 1)
      if (currentSummaryStatus !== 1) {
        currentSummaryStatus = 1;
        pollSummaryUntilDone();
      }

    } catch (err) {
      console.error("Reprocess failed:", err);
      toaster("Reprocess failed", "error");
      tag.Status = "2";
      renderAll();
    }
  }

  // ✅ Ping API every 10 sec until status becomes 2 and all tags processed
  async function pollSummaryUntilDone() {
    while (instanceId === currentSummaryInstance) {
      const statusRes = await getSummaryTagStatus(store.documentID, store.jwt);
      const data = statusRes?.Data;
      const status = data?.SummaryTagGenerated;
      const tagStatuses = data?.SummaryTagStatus || [];

      currentSummaryStatus = status;

      // Update individual tag statuses in our local list if they exist in the response
      if (tagStatuses && tagStatuses.length > 0) {
        tagStatuses.forEach((ts: any) => {
          const matchingTag = allSummaryTags.find(t => (t.ReportHeadSummaryTagID || t.ID) === ts.ReportHeadSummaryTagID);
          if (matchingTag) {
            matchingTag.Status = ts.Status;
          }
        });
        renderAll();
      }

      // Check if all tags are processed (status != "0")
      // We consider it done only if status is 2 and there are no tags with Status "0"
      const allTagsProcessed = tagStatuses.length > 0 ? tagStatuses.every((t: any) => String(t.Status) !== "0") : true;

      if (status === 2 && allTagsProcessed) {
        setReanalyzeButtonState(true);

        // after done -> fetch tags again and render
        const getRes2 = await getSummaryTagsByReportHeadId(store.documentID, store.jwt);
        allSummaryTags = getRes2?.Data?.SummaryTags || [];
        filtered = allSummaryTags;

        if (allSummaryTags && allSummaryTags.length > 0) {
          summarySelectedNames = normalizeNames(allSummaryTags);
        }

        renderAll();
        break;
      }

      await new Promise(r => setTimeout(r, 10000));
    }
  }

  // ✅ Reanalyze wiring (same logic but triggers refresh + poll)
  reanalyzeBtn?.addEventListener('click', async () => {
    if (isSummaryLoading || reanalyzeBtn.classList.contains("disabled")) return;

    const popupContainer = document.getElementById("confirmation-popup");
    if (!popupContainer) return;

    popupContainer.innerHTML = Confirmationpopup("Do you want to refresh all summary tags of the draft?");

    // Change button labels to No/Yes as requested
    const cancelBtn = document.getElementById("confirmation-popup-cancel");
    const confirmBtn = document.getElementById("confirmation-popup-confirm");

    if (cancelBtn) cancelBtn.innerText = "No";
    if (confirmBtn) confirmBtn.innerText = "Yes";

    const handleAction = async (refresh: boolean) => {
      popupContainer.innerHTML = ""; // Close popup
      try {
        setReanalyzeButtonState(false);
        disableActionButtons(true);
        isSummaryLoading = true;
        currentSummaryStatus = 1;

        // Start individual tag spinners immediately
        allSummaryTags.forEach(t => t.Status = "0");
        renderRows();

        const base64Data = await getWordAsBase64();
        const payload = {
          ReportHeadID: Number(store.documentID),
          RefreshSummaryTag: refresh,
          ActiveDocument: base64Data
        };

        const res = await refreshSummaryMode(payload, store.jwt);

        // Immediately update state and UI from response
        if (res?.Data) {
          allSummaryTags = res.Data.SummaryTags || [];
          store.sourceSummaryList = deduplicateSummarySources(res.Data.SummarySources || []);
          currentSummaryStatus = res.Data.SummaryTagGenerated;
          filtered = allSummaryTags;

          if (allSummaryTags && allSummaryTags.length > 0) {
            summarySelectedNames = normalizeNames(allSummaryTags);
          }

          renderAll();
        }

        // poll again until done
        await pollSummaryUntilDone();

      } catch (err) {
        console.error("Refresh failed:", err);
        setReanalyzeButtonState(true);
      }
    };

    document.getElementById("confirmation-popup-confirm")?.addEventListener("click", () => handleAction(true));
    document.getElementById("confirmation-popup-cancel")?.addEventListener("click", () => handleAction(false));
  });

  // ✅ Final: run new logic
  await firstLoadAndRender();
}

// ------------------ Base64 utils (unchanged) ------------------

export function getWordAsBase64(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 1024 * 1024 },
      result => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(result.error);
          return;
        }

        const file = result.value;
        const slices: string[] = [];
        let index = 0;

        const getSlice = () => {
          file.getSliceAsync(index, slice => {
            if (slice.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync();
              reject(slice.error);
              return;
            }

            const data = slice.value.data;

            if (typeof data === "string") {
              slices.push(data);
            } else {
              const bytes = new Uint8Array(data);
              slices.push(uint8ToBase64(bytes));
            }

            index++;

            if (index < file.sliceCount) {
              getSlice();
            } else {
              file.closeAsync();
              resolve(slices.join(""));
            }
          });
        };

        getSlice();
      }
    );
  });
}

function uint8ToBase64(bytes: Uint8Array): string {
  let binary = '';
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return window.btoa(binary);
}
