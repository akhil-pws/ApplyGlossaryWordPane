import { theme } from "../taskpane";
import { wordTableStyles } from "./tablestyles";

function addtagbody(sponsorOptions) {
  const body = `<div class="modal-dialog">
  <div class="modal-content">
    <div class="modal-body p-3 pt-0">
      <form id="genai-form" autocomplete="off" novalidate>
        <!-- Name Field -->
        <div class="mb-3">
          <label for="name" class="form-label"><span class="text-danger">*</span> Name</label>
          <input type="text" class="form-control" id="name" required>
          <div class="invalid-feedback">Name is required.</div>
          <div id="submition-error" class="invalid-feedback" style="display: none;"></div>

        </div>

        <!-- Description Field -->
        <div class="mb-3">
          <label for="description" class="form-label">Description</label>
          <textarea class="form-control" id="description" rows="6"></textarea>
        </div>

        <!-- Prompt Field -->
        <div class="mb-3 prompt-box">
          <label for="prompt" class="form-label"><span class="text-danger">*</span> Prompt 
            <small class="text-secondary">(Note: Use # tag for content suggestions)</small>
          </label>
          <textarea class="form-control" id="prompt" rows="6"  required></textarea>
          <div class="invalid-feedback">Prompt is required.</div>
          <div id="mention-dropdown" class="dropdown-menu"></div>
        </div>

        <!-- Save Globally Checkbox -->
        <div class="form-check mb-3">
          <input type="checkbox" class="form-check-input" id="saveGlobally">
          <label class="form-check-label" for="saveGlobally">Save Globally</label>
        </div>

        <!-- Available to All Sponsors Checkbox -->
        <div class="form-check mb-3">
          <input type="checkbox" class="form-check-input" id="isAvailableForAll" disabled>
          <label class="form-check-label" for="isAvailableForAll">Available to All Sponsors</label>
        </div>

        <!-- Sponsor Dropdown -->
        <div class="mb-3">
          <label for="sponsor" class="form-label"><span class="text-danger">*</span> Sponsor</label>
          <div class="dropdown w-100">
            <button 
              class="btn btn-white border w-100 text-start d-flex justify-content-between align-items-center dropdown-toggle" 
              type="button" 
              id="sponsorDropdown" 
              data-bs-toggle="dropdown" 
              aria-expanded="false" 
              disabled>
              <span id="sponsorDropdownLabel">Select Sponsors</span>
              <span class="dropdown-toggle-icon"></span>
            </button>
            <ul class="dropdown-menu w-100 p-2" aria-labelledby="sponsorDropdown" style="box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
              <li class="dropdown-item p-2" style="cursor: pointer;">
                <div class="form-check">
                  <input class="form-check-input" type="checkbox" value="selectAll" id="selectAll">
                  <label class="form-check-label" for="selectAll">Select All</label>
                </div>
              </li>
              ${sponsorOptions}
            </ul>
          </div>
        </div>

        <!-- Action Buttons -->
        <div class="mt-3 d-flex justify-content-between">
          <span id="cancel-btn-gen-ai" class="fw-bold text-primary my-auto c-pointer">Cancel</span>
          <button type="submit" class="btn btn-primary" id="text-gen-save">Save</button>
        </div>
      </form>
    </div>
  </div>
</div>`

  return body
}


function Confirmationpopup(content: string) {
  const isDark = theme === 'Dark';
  const popupClass = isDark ? 'bg-dark text-light' : 'bg-light text-dark';

  const body = `
<div class="modal show d-block" tabindex="-1">
  <div class="modal-dialog">
    <div class="modal-content ${popupClass}">
      <div class="modal-header border-0">
        <h5 class="fw-bold">Confirmation</h5>
      </div>

      <div class="modal-body">
        <p>${content}</p>
      </div>

      <div class="modal-footer border-0">
        <button type="button" class="btn btn-link ${isDark ? 'text-info' : 'text-primary'}" id="confirmation-popup-cancel">Cancel</button>
        <button type="button" class="btn btn-primary text-white" id="confirmation-popup-confirm">Ok</button>
      </div>
    </div>
  </div>
</div>`;
  return body;
}


function customizeTablePopup(selectedValue: string) {
  const isDark = theme === 'Dark';
  const popupClass = isDark ? 'bg-dark text-light' : 'bg-light text-dark';

  const dropdown = `
  
    <select class="form-select mb-3 ${popupClass}" id="confirmation-popup-dropdown">
      ${wordTableStyles
      .map(
        opt =>
          `<option value="${opt.style}" ${opt.style === selectedValue ? 'selected' : ''}>
              ${opt.name}
            </option>`
      )
      .join('')}
    </select>
  `;

  const tablePreview = `
    <div class="table-responsive">
      <table class="table table-bordered table-sm" id="confirmation-popup-table-preview">
        <thead>
          <tr>
            <th>Header 1</th>
            <th>Header 2</th>
            <th>Header 3</th>
          </tr>
        </thead>
        <tbody>
          <tr style="color:black;">
            <td>Data 1</td>
            <td>Data 2</td>
            <td>Data 3</td>
          </tr>
          <tr style="background-color:white;color:black;">
            <td>Data 4</td>
            <td>Data 5</td>
            <td>Data 6</td>
          </tr>
          <tr style="color:black;" >
            <td>Data 7</td>
            <td>Data 8</td>
            <td>Data 9</td>
          </tr>
        </tbody>
      </table>
    </div>
  `;

  return `
<div class="modal show d-block" tabindex="-1">
  <div class="modal-dialog">
    <div class="modal-content ${popupClass}">
      <div class="modal-header border-0">
        <h5 class="fw-bold">Customize Table</h5>
      </div>

      <div class="modal-body">
        ${dropdown}
        ${tablePreview}
      </div>

      <div class="modal-footer border-0">
        <button type="button" class="btn btn-link ${isDark ? 'text-info' : 'text-primary'}" id="confirmation-popup-cancel">Cancel</button>
        <button type="button" class="btn btn-primary text-white" id="confirmation-popup-confirm">Ok</button>
      </div>
    </div>
  </div>
</div>
  `;
}


function DataModalPopup(selectedData) {
  const isDark = theme === 'Dark';
  const popupClass = isDark ? 'bg-dark text-light' : 'bg-light text-dark';

  return `
<div class="modal show d-block" tabindex="-1">
  <div class="modal-dialog modal-lg">
    <div class="modal-content p-3 ${popupClass}">
      <div class="modal-header flex-column align-items-start border-0">
        <span class="fw-bold mb-3">${selectedData?.Name || ''}</span>
        <span class="d-block list-height">${selectedData?.UserValue || ''}</span>
        ${selectedData?.Sources ? `
        <hr class="${isDark ? 'border-light' : 'border-dark'}">
        <div class="d-flex align-items-start flex-wrap">
          <span class="fw-bold me-2">Selected Sources :</span>
          <div class="d-flex flex-wrap gap-1">
            ${selectedData.Sources.map(source => `
              <span class="badge ${isDark ? 'text-bg-secondary' : 'text-bg-info'}">${source.FileName}</span>
            `).join('')}
          </div>
        </div>` : ''}
      </div>

      <div class="modal-body p-3 add-ai-gen">
        <div class="row g-2 list-height">
          ${selectedData?.Data?.map(item => `
            <div class="col-md-12 mt-3">
              <div class="border rounded p-2 ${isDark ? 'bg-secondary text-light' : 'bg-light text-dark'} shadow-sm h-100">
                <div class="fw-bold small text-truncate" title="${item.FileName}">
                  ${item.FileName}
                </div>
                <div class="text-muted small mb-1">Page: ${item.PageNumber}</div>
                <div class="small" style="white-space: normal;">
                  ${item.Sentence}
                </div>
              </div>
            </div>
          `).join('')}
        </div>

        <div class="d-flex w-100 justify-content-end mt-3 align-items-center">
          <button type="button" class="btn btn-primary text-white" id="datamodel-popup-ok">OK</button>
        </div>
      </div>
    </div>
  </div>
</div>`;
}



function toaster(message: string, type: string) {
  const icon = type === 'success' ? 'fa-check-circle' : 'fa-exclamation-circle';
  // const color = type === 'success' ? '#28a745' : '#dc3545';
  const color = `#ffffff`
  const body = `<div class="toast show" style="position: fixed; top: 10px; right: 10px; z-index: 1050; max-width: fit-content; background-color: #808080; color: #ffffff;">
    <div class="toast-body">
         <i class="fa ${icon} me-2" style="color: ${color};"></i> ${message}
    </div>
  </div>`;

  document.getElementById('toastr').innerHTML = body;
  setTimeout(() => {
    document.getElementById('toastr').innerHTML = ``;
  }, 4000);
}

function logoheader(storedUrl) {
  const themeicon = theme === 'Dark' ? 'fa-sun' : 'fa-moon'
  const body = `
    <img id="main-logo" src="${storedUrl}/assets/logo.png" alt="" class="logo">
    <div class="icon-nav me-3">
      <i class="fa fa-home c-pointer me-3" title="Home" id="home"></i>
      <div class="dropdown d-inline">
        <i class="fa fa-tools c-pointer me-3" id="settingsDropdown" data-bs-toggle="dropdown" aria-expanded="false" title="Settings"></i>
        <ul class="dropdown-menu" aria-labelledby="settingsDropdown">
        <li>
        <li>
            <a class="dropdown-item" href="#" id="define-formatting">
              <i class="fa fa-sliders-h me-2" aria-hidden="true"></i> Define Formatting
            </a>
          </li>
          <li>
            <a class="dropdown-item disabled-link" href="#" id="glossary" tabindex="-1" aria-disabled="true">
              <i class="fa fa-book me-2" aria-hidden="true"></i> Apply Glossary
            </a>
          </li>
          <li>
            <a class="dropdown-item disabled-link" href="#" id="removeFormatting" tabindex="-1" aria-disabled="true">
              <i class="fa fa-eraser me-2" aria-hidden="true"></i> Remove Formatting
            </a>
          </li>
        </ul>
      </div>

      <!-- Theme Toggle Icon -->
      <span id="theme-toggle"><i class="fa ${themeicon} c-pointer me-3" title="Toggle Theme"></i></span>

      <i class="fa fa-sign-out c-pointer me-3" id="logout" title="Logout"></i>
    </div>    
  `
  return body;
}

const navTabs = `<ul class="nav nav-tabs" id="tabList" role="tablist">
  <li class="nav-item">
    <a class="nav-link active" id="tag-tab" data-bs-toggle="tab" href="#tag" role="tab">Tag</a>
  </li>
  <li class="nav-item">
    <a class="nav-link" id="prompt-tab" data-bs-toggle="tab" href="#prompt" role="tab">Prompt builder</a>
  </li>
</ul>

<div class="tab-content p-3 border border-top-0">
  <div class="tab-pane fade show active" id="add-tag-body" role="tabpanel" aria-labelledby="tag-tab">
  </div>
  <div class="tab-pane fade" id="add-prompt-template" role="tabpanel" aria-labelledby="prompt-tab">
  </div>
</div>
`



const promptbuilderbody = `<div>hi</div>`


export { navTabs, addtagbody, promptbuilderbody, logoheader, Confirmationpopup, toaster, DataModalPopup, customizeTablePopup };