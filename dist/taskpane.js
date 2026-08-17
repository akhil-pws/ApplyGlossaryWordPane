/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/taskpane/components/bodyelements.ts":
/*!*************************************************!*\
  !*** ./src/taskpane/components/bodyelements.ts ***!
  \*************************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Confirmationpopup: function() { return /* binding */ Confirmationpopup; },
/* harmony export */   DataModalPopup: function() { return /* binding */ DataModalPopup; },
/* harmony export */   PromptBuilderModalPopup: function() { return /* binding */ PromptBuilderModalPopup; },
/* harmony export */   addtagbody: function() { return /* binding */ addtagbody; },
/* harmony export */   customizeTablePopup: function() { return /* binding */ customizeTablePopup; },
/* harmony export */   customizeTextStylePopup: function() { return /* binding */ customizeTextStylePopup; },
/* harmony export */   customizedStylePopup: function() { return /* binding */ customizedStylePopup; },
/* harmony export */   logoheader: function() { return /* binding */ logoheader; },
/* harmony export */   navTabs: function() { return /* binding */ navTabs; },
/* harmony export */   promptbuilderbody: function() { return /* binding */ promptbuilderbody; },
/* harmony export */   toaster: function() { return /* binding */ toaster; }
/* harmony export */ });
/* harmony import */ var _services_store_service__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../services/store.service */ "./src/taskpane/services/store.service.ts");
/* harmony import */ var _tablestyles__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./tablestyles */ "./src/taskpane/components/tablestyles.ts");


function addtagbody(sponsorOptions, sourceOptions, isSummaryMode = false) {
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

      

        <div class="mb-3">
          <label for="source" class="form-label">
            ${isSummaryMode ? '' : '<span class="text-danger">*</span> '}
            ${isSummaryMode ? 'Additional Source Type' : 'Primary Source Type'}
          </label>
          <div class="dropdown w-100">
            <button 
              class="btn btn-white border w-100 text-start d-flex justify-content-between align-items-center dropdown-toggle" 
              type="button" 
              id="sourceDropdown" 
              data-bs-toggle="dropdown" 
              aria-expanded="false" 
              >
              <span id="sourceDropdownLabel">Select Source</span>
              <span class="dropdown-toggle-icon"></span>
            </button>
            <ul class="dropdown-menu w-100 p-2" aria-labelledby="sourceDropdown" style="box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
              <li class="source-dropdown-item dropdown-item p-2" style="cursor: pointer;">
                <div class="form-check">
                  <input class="form-check-input" type="checkbox" value="selectAll" id="sourceSelectAll">
                  <label class="form-check-label" for="sourceSelectAll">Select All</label>
                </div>
              </li>
              ${sourceOptions}
            </ul>
          </div>
          <div class="invalid-feedback" id="primarySourceError" style="display:none;">
            ${isSummaryMode ? 'Additional Source Type is required.' : 'Primary Source is required.'}
          </div>

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
              <li class="sponsor-dropdown-item dropdown-item p-2" style="cursor: pointer;">
                <div class="form-check">
                  <input class="form-check-input" type="checkbox" value="selectAll" id="sponsorSelectAll">
                  <label class="form-check-label" for="sponsorSelectAll">Select All</label>
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
</div>`;
  return body;
}
function Confirmationpopup(content) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_0__.StoreService.getInstance();
  const isDark = store.theme === 'Dark';
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
function customizeTablePopup(selectedValue, type) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_0__.StoreService.getInstance();
  const isDark = store.theme === "Dark";
  const popupClass = isDark ? "bg-dark text-light" : "bg-light text-dark";
  const header = type === "Custom" ? "Customized Tables" : "Default Tables";
  const sourceList = type === "Custom" ? store.customTableStyle : _tablestyles__WEBPACK_IMPORTED_MODULE_1__.wordTableStyles;
  const warningContent = type === "Custom" ? '' : '<small class="text-secondary font-italic">Warning: The previewed table colors may vary depending on the accent or theme currently selected in Word.</small>';
  const dropdown = `
    <select class="form-select mb-2 ${popupClass}" id="confirmation-popup-dropdown">
      ${sourceList.map(opt => {
    const value = type === "Custom" ? opt.Name : opt.style;
    const isSelected = value === selectedValue;
    const text = type === "Custom" ? opt.Name : opt.style;
    return `<option value="${value}" ${isSelected ? "selected" : ""}>
                    ${text}
                  </option>`;
  }).join("")}
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
        <h5 class="fw-bold">${header}</h5>
      </div>

      <div class="modal-body">
        ${dropdown}
        ${tablePreview}
        ${warningContent}
      </div>

      <div class="modal-footer border-0">
        <button type="button" class="btn btn-link ${isDark ? "text-info" : "text-primary"}" id="confirmation-popup-cancel">Cancel</button>
        <button type="button" class="btn btn-primary text-white" id="confirmation-popup-confirm">Ok</button>
      </div>
    </div>
  </div>
</div>
  `;
}
function DataModalPopup(selectedData) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_0__.StoreService.getInstance();
  const isDark = store.theme === 'Dark';
  const popupClass = isDark ? 'bg-dark text-light' : 'bg-light text-dark';
  const hasEvidences = selectedData?.Data && selectedData.Data.length > 0;
  let bodyContent = '';
  if (hasEvidences) {
    bodyContent = `
      <div class="row g-2 list-height">
        ${selectedData.Data.map(item => `
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
      </div>`;
  } else {
    const message = "No reference details are available for this response. The AI model did not provide source references.";
    bodyContent = `
      <div class="p-3 mb-3" style="background-color: #f0f8ff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border-left: 4px solid #0056b3 !important;">
        <div style="color: #6c757d; white-space: normal; font-size: 0.9rem; line-height: 1.4;">
          ${message}
        </div>
      </div>`;
  }
  return `
<div class="modal show d-block" tabindex="-1">
  <div class="modal-dialog modal-lg">
    <div class="modal-content p-2 ${popupClass}">
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
        ${bodyContent}

        <div class="d-flex w-100 justify-content-end mt-3 align-items-center">
          <button type="button" class="btn btn-primary text-white" id="datamodel-popup-ok">OK</button>
        </div>
      </div>
    </div>
  </div>
</div>`;
}
function toaster(message, type) {
  const icon = type === 'success' ? 'fa-check-circle' : 'fa-exclamation-circle';
  // const color = type === 'success' ? '#28a745' : '#dc3545';
  const color = `#ffffff`;
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
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_0__.StoreService.getInstance();
  const themeicon = store.theme === 'Dark' ? 'fa-sun' : 'fa-moon';
  const modeIcon = store.mode === 'Home' ? 'fa-home' : store.mode === 'Summary' ? 'fa-wand-magic-sparkles' : 'fa-magnifying-glass';
  const body = `
    <img id="main-logo" src="${storedUrl}/assets/logo.png" alt="" class="logo">
    <div class="icon-nav me-3">
    <div class="dropdown d-inline">
        <i class="fa ${modeIcon} c-pointer me-3" id="modeDropdown" data-bs-toggle="dropdown" aria-expanded="false" title="Mode"></i>
        <ul class="dropdown-menu" aria-labelledby="modeDropdown" appendTo="body">
        <li>
            <a class="dropdown-item" href="#" id="home">
              <i class="fa fa-home me-2" aria-hidden="true"></i> Draft Mode
            </a>
          </li>
          <li>
            <a class="dropdown-item" href="#" id="summary-mode">
              <i class="fa fa-wand-magic-sparkles me-2" aria-hidden="true"></i> Summary Mode
            </a>
          </li>
          <li>
            <a class="dropdown-item disabled-link" href="#" id="review-mode">
              <i class="fa fa-magnifying-glass me-2" aria-hidden="true"></i> Review Mode
            </a>
          </li>
        </ul>
      </div>


      <div class="dropdown d-inline">
        <i class="fa fa-tools c-pointer me-3" id="settingsDropdown" data-bs-toggle="dropdown" aria-expanded="false" title="Settings"></i>
        <ul class="dropdown-menu" aria-labelledby="settingsDropdown">
         <li>
            <a class="dropdown-item" href="#" id="predefined-table">
              <i class="fa fa-table me-2" aria-hidden="true"></i> Default Tables
            </a>
          </li>
          <li>
            <a class="dropdown-item" href="#" id="customized-table">
              <i class="fa fa-brush me-2" aria-hidden="true"></i> Customized Tables
            </a>
          </li>
          <li>
            <a class="dropdown-item" href="#" id="default-text-style">
              <i class="fa fa-font me-2" aria-hidden="true"></i> Default Text Style
            </a>
          </li>
          <li>
            <a class="dropdown-item" href="#" id="customized-style">
              <i class="fa fa-palette me-2" aria-hidden="true"></i> Customized Styles
            </a>
          </li>
          <li><hr class="dropdown-divider"></li>
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
  `;
  return body;
}
const navTabs = `<ul class="nav nav-tabs" id="tabList" role="tablist">
  <li class="nav-item">
    <a class="nav-link active" id="tag-tab" data-bs-toggle="tab" href="#tag" role="tab">Tag</a>
  </li>
  <li class="nav-item">
    <a class="nav-link" id="prompt-tab" data-bs-toggle="tab" href="#prompt" role="tab">Prompt Builder</a>
  </li>
</ul>

<div class="tab-content p-3 border border-top-0">
  <div class="tab-pane fade show active" id="add-tag-body" role="tabpanel" aria-labelledby="tag-tab">
  </div>
  <div class="tab-pane fade" id="add-prompt-template" role="tabpanel" aria-labelledby="prompt-tab">
  </div>
</div>
`;
const promptbuilderbody = `<div>hi</div>`;
function customizeTextStylePopup(selectedValue, availableStyles) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_0__.StoreService.getInstance();
  const isDark = store.theme === "Dark";
  const popupClass = isDark ? "bg-dark text-light" : "bg-light text-dark";
  const selectClass = isDark ? "bg-dark text-light border-light" : "bg-white text-dark border-dark";
  if (selectedValue && !availableStyles.includes(selectedValue)) {
    availableStyles.push(selectedValue);
    availableStyles.sort((a, b) => a.localeCompare(b));
  }
  const dropdown = `
    <div class="mb-3">
      <label for="text-style-dropdown" class="form-label fw-bold">Select Default Paragraph Style</label>
      <select class="form-select ${selectClass}" id="text-style-dropdown" style="max-height: 200px;">
        ${availableStyles.map(style => {
    const isSelected = style === selectedValue;
    return `<option value="${style}" ${isSelected ? "selected" : ""}>${style}</option>`;
  }).join("")}
      </select>
    </div>
  `;
  const previewContainer = `
    <div class="mb-3">
      <label class="form-label fw-bold">Style Preview</label>
      <div id="text-style-preview" class="p-3 border rounded text-center style-preview-box" style="max-height: 150px; max-width: 100%; overflow: auto; min-height: 80px; display: flex; flex-direction: column; transition: all 0.3s ease; box-sizing: border-box;">
        Preview Text
      </div>
    </div>
  `;
  return `
<div class="modal show d-block" tabindex="-1">
  <div class="modal-dialog">
    <div class="modal-content ${popupClass}">
      <div class="modal-header border-0">
        <h5 class="fw-bold">Default Text Style</h5>
      </div>

      <div class="modal-body">
        ${dropdown}
        ${previewContainer}
        <div class="alert alert-info py-2 small" role="alert">
          This lists all paragraph styles available in the current Word document. The selected style will be applied by default to all inserted AI text content.
        </div>
      </div>

      <div class="modal-footer border-0">
        <button type="button" class="btn btn-link ${isDark ? "text-info" : "text-primary"}" id="text-style-popup-cancel">Cancel</button>
        <button type="button" class="btn btn-primary text-white" id="text-style-popup-confirm">Ok</button>
      </div>
    </div>
  </div>
</div>
  `;
}
function customizedStylePopup(selectedValue, availableStyles) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_0__.StoreService.getInstance();
  const isDark = store.theme === "Dark";
  const popupClass = isDark ? "bg-dark text-light" : "bg-light text-dark";
  const selectClass = isDark ? "bg-dark text-light border-light" : "bg-white text-dark border-dark";
  const dropdown = `
    <div class="mb-3">
      <label for="customized-style-dropdown" class="form-label fw-bold">Select Customized Style</label>
      <select class="form-select ${selectClass} mb-3" id="customized-style-dropdown">
        ${availableStyles.map(style => {
    const isSelected = style.id === selectedValue;
    return `<option value="${style.id}" ${isSelected ? "selected" : ""}>${style.name}</option>`;
  }).join("")}
      </select>
    </div>
  `;
  const previewContainer = `
    <div class="mb-3">
      <label class="form-label fw-bold">Style Preview</label>
      <div id="customized-style-preview" class="p-3 border rounded text-center style-preview-box" style="max-height: 150px; max-width: 100%; overflow: auto; min-height: 80px; display: flex; flex-direction: column; transition: all 0.3s ease; box-sizing: border-box;">
        Preview Text
      </div>
    </div>
  `;
  return `
<div class="modal show d-block" tabindex="-1">
  <div class="modal-dialog">
    <div class="modal-content ${popupClass}">
      <div class="modal-header border-0">
        <h5 class="fw-bold">Customized Styles</h5>
      </div>

      <div class="modal-body">
        ${dropdown}
        ${previewContainer}
        <div class="alert alert-info py-2 small" role="alert">
          Select a style configuration to apply when inserting generated AI text.
        </div>
      </div>

      <div class="modal-footer border-0">
        <button type="button" class="btn btn-link ${isDark ? "text-info" : "text-primary"}" id="customized-style-popup-cancel">Cancel</button>
        <button type="button" class="btn btn-primary text-white" id="customized-style-popup-confirm">Ok</button>
      </div>
    </div>
  </div>
</div>
  `;
}
function PromptBuilderModalPopup() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_0__.StoreService.getInstance();
  const isDark = store.theme === 'Dark';
  const popupClass = isDark ? 'bg-dark text-light' : 'bg-light text-dark';
  return `
<div class="modal show d-block" tabindex="-1">
  <div class="modal-dialog">
    <div class="modal-content ${popupClass}">
      <div class="modal-header border-0">
        <h5 class="fw-bold">Prompt Builder</h5>
      </div>

      <div class="modal-body" style="max-height: 400px; overflow-y: auto;">
        <div class="form-group mb-3">
          <label class='form-label'><span class="text-danger">*</span> Prompt Builder Template</label>
          <select id="promptBuilderTemplatePopup" class="form-select ${isDark ? 'bg-secondary text-light border-secondary' : ''}">
            <option value="" disabled selected>Select a template</option>
          </select>
          <div id="templateErrorPopup" class="invalid-feedback d-none">Type is required.</div>
        </div>

        <div id="fieldsContainerPopup"></div>

        <div class="form-group mb-3" id="previewContainerPopup" style="display: none;">
          <label class="mb-2 fw-bold">Preview</label>
          <div id="previewPopup" class="form-control border p-2 ${isDark ? 'bg-secondary text-light border-secondary' : 'bg-light text-dark'}" style="min-height: 50px; white-space: pre-wrap; word-break: break-word;"></div>
        </div>
      </div>

      <div class="modal-footer border-0">
        <button type="button" class="btn btn-link ${isDark ? 'text-info' : 'text-primary'}" id="prompt-builder-popup-cancel">Cancel</button>
        <button type="button" class="btn btn-primary text-white" id="prompt-builder-popup-insert" disabled>Insert</button>
      </div>
    </div>
  </div>
</div>`;
}


/***/ }),

/***/ "./src/taskpane/components/customstyles.ts":
/*!*************************************************!*\
  !*** ./src/taskpane/components/customstyles.ts ***!
  \*************************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   mapApiStyleToCustomStyle: function() { return /* binding */ mapApiStyleToCustomStyle; }
/* harmony export */ });
function mapApiStyleToCustomStyle(apiStyle) {
  const props = apiStyle.Properties;
  const hasProperties = props && typeof props === "object" && Object.keys(props).length > 0;
  return {
    id: apiStyle.Name,
    name: apiStyle.Name,
    ID: apiStyle.ID,
    properties: hasProperties ? {
      bold: props.Bold ?? false,
      italic: props.Italic ?? false,
      underline: props.Underline ?? false,
      fontFamily: props.FontFamily ?? "Calibri",
      size: props.Size ?? "11pt",
      fontColor: props.FontColor ?? props.fontColor ?? "#000000",
      backgroundColor: props.BackgroundColor ?? props.backgroundColor ?? "transparent"
    } : null
  };
}

/***/ }),

/***/ "./src/taskpane/components/tablestyles.ts":
/*!************************************************!*\
  !*** ./src/taskpane/components/tablestyles.ts ***!
  \************************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   wordTableStyleLocales: function() { return /* binding */ wordTableStyleLocales; },
/* harmony export */   wordTableStyles: function() { return /* binding */ wordTableStyles; }
/* harmony export */ });
const wordTableStyles = [
// Plain Tables
{
  name: "Plain White Table",
  style: "Table Grid",
  headerClass: "",
  tableClass: "background-color:#ffffff;border:1px solid #000;",
  rowClass: "background-color:#ffffff;color:#000000;",
  format: "empty",
  sideHeader: false
}, {
  name: "Plain Table with Gray Alternating Rows",
  style: "Plain Table 1",
  headerClass: "",
  tableClass: "background-color:#ffffff;border:1px solid #000;",
  rowClass: "background-color:#f2f2f2;color:#000000;",
  format: "empty",
  sideHeader: true
},
// Grid Table 1 Variants (Header + Alternating Rows)
{
  name: "Plain Header Table with Light Gray Alternating Rows",
  style: "Grid Table 2",
  headerClass: "border:1px solid white;",
  tableClass: "background-color:#ffffff;",
  rowClass: "background-color:#f2f2f2;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Plain Header Table with Light Blue Alternating Rows",
  style: "Grid Table 2 - Accent 1",
  headerClass: "border:1px solid white;",
  tableClass: "background-color:#ffffff;border:1px solid #d4d4d1ff;",
  rowClass: "background-color:#dbe5f1;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Plain Header Table with Pink Alternating Rows",
  style: "Grid Table 2 - Accent 2",
  headerClass: "border:1px solid white;",
  tableClass: "background-color:#ffffff;border:1px solid #d4d4d1ff;",
  rowClass: "background-color:#f2dbdb;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Plain Header Table with Light Green Alternating Rows",
  style: "Grid Table 2 - Accent 3",
  headerClass: "border:1px solid white;",
  tableClass: "background-color:#ffffff;border:1px solid #d4d4d1ff;",
  rowClass: "background-color:#eaf1dd;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Plain Header Table with Light Purple Alternating Rows",
  style: "Grid Table 2 - Accent 4",
  headerClass: "border:1px solid white;",
  tableClass: "background-color:#ffffff;border:1px solid #d4d4d1ff;",
  rowClass: "background-color:#e5dfec;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Plain Header Table with Light Teal Alternating Rows",
  style: "Grid Table 2 - Accent 5",
  headerClass: "border:1px solid white;",
  tableClass: "background-color:#ffffff;border:1px solid #d4d4d1ff;",
  rowClass: "background-color:#daeef3;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Plain Header Table with Light Orange Alternating Rows",
  style: "Grid Table 2 - Accent 6",
  headerClass: "border:1px solid white;",
  tableClass: "background-color:#ffffff;border:1px solid #d4d4d1ff;",
  rowClass: "background-color:#fde9d9;color:#000000;",
  format: "empty",
  sideHeader: true
},
// Header Only Tables
{
  name: "Black Header Only Table with White Rows",
  style: "List Table 3",
  headerClass: "",
  tableClass: "background-color:#000000;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#ffffff;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Blue Header Only Table with White Rows",
  style: "List Table 3 - Accent 1",
  headerClass: "",
  tableClass: "background-color:#4F81BD;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#ffffff;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Red Header Only Table with White Rows",
  style: "List Table 3 - Accent 2",
  headerClass: "",
  tableClass: "background-color:#c0504d;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#ffffff;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Green Header Only Table with White Rows",
  style: "List Table 3 - Accent 3",
  headerClass: "",
  tableClass: "background-color:#9bbb59;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#ffffff;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Purple Header Only Table with White Rows",
  style: "List Table 3 - Accent 4",
  headerClass: "",
  tableClass: "background-color:#8064a2;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#ffffff;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Teal Header Only Table with White Rows",
  style: "List Table 3 - Accent 5",
  headerClass: "",
  tableClass: "background-color:#4bacc6;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#ffffff;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Orange Header Only Table with White Rows",
  style: "List Table 3 - Accent 6",
  headerClass: "",
  tableClass: "background-color:#f79646;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#ffffff;color:#000000;",
  format: "empty",
  sideHeader: true
},
// Grid Table 4 Variants
{
  name: "Black Header Table with Light Gray Alternating Rows",
  style: "Grid Table 4",
  headerClass: "",
  tableClass: "background-color:#000000;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#f2f2f2;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Blue Header Table with Light Blue Alternating Rows",
  style: "Grid Table 4 - Accent 1",
  headerClass: "",
  tableClass: "background-color:#4F81BD;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#dbe5f1;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Red Header Table with Pink Alternating Rows",
  style: "Grid Table 4 - Accent 2",
  headerClass: "",
  tableClass: "background-color:#c0504d;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#f2dbdb;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Green Header Table with Light Green Alternating Rows",
  style: "Grid Table 4 - Accent 3",
  headerClass: "",
  tableClass: "background-color:#9bbb59;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#eaf1dd;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Purple Header Table with Light Purple Alternating Rows",
  style: "Grid Table 4 - Accent 4",
  headerClass: "",
  tableClass: "background-color:#8064a2;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#e5dfec;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Teal Header Table with Light Teal Alternating Rows",
  style: "Grid Table 4 - Accent 5",
  headerClass: "",
  tableClass: "background-color:#4bacc6;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#daeef3;color:#000000;",
  format: "empty",
  sideHeader: true
}, {
  name: "Orange Header Table with Light Orange Alternating Rows",
  style: "Grid Table 4 - Accent 6",
  headerClass: "",
  tableClass: "background-color:#f79646;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#fde9d9;color:#000000;",
  format: "empty",
  sideHeader: true
},
// Grid Table 5 Variants (Partial)
{
  name: "Black Header Table with Light Gray Alternating Rows - Side header same color",
  style: "Grid Table 5 Dark",
  headerClass: "",
  tableClass: "background-color:#000000;border:1px solid #f2f2f2;color:#ffffff;",
  rowClass: "background-color:#f2f2f2;color:#000000;",
  format: "partial",
  sideHeader: true
}, {
  name: "Blue Header Table with Light Blue Alternating Rows - Side header same color",
  style: "Grid Table 5 Dark - Accent 1",
  headerClass: "",
  tableClass: "background-color:#4F81BD;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#dbe5f1;color:#000000;",
  format: "partial",
  sideHeader: true
}, {
  name: "Red Header Table with Pink Alternating Rows - Side header same color",
  style: "Grid Table 5 Dark - Accent 2",
  headerClass: "",
  tableClass: "background-color:#c0504d;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#f2dbdb;color:#000000;",
  format: "partial",
  sideHeader: true
}, {
  name: "Green Header Table with Light Green Alternating Rows - Side header same color",
  style: "Grid Table 5 Dark - Accent 3",
  headerClass: "",
  tableClass: "background-color:#9bbb59;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#eaf1dd;color:#000000;",
  format: "partial",
  sideHeader: true
}, {
  name: "Purple Header Table with Light Purple Alternating Rows - Side header same color",
  style: "Grid Table 5 Dark - Accent 4",
  headerClass: "",
  tableClass: "background-color:#8064a2;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#e5dfec;color:#000000;",
  format: "partial",
  sideHeader: true
}, {
  name: "Teal Header Table with Light Teal Alternating Rows - Side header same color",
  style: "Grid Table 5 Dark - Accent 5",
  headerClass: "",
  tableClass: "background-color:#4bacc6;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#daeef3;color:#000000;",
  format: "partial",
  sideHeader: true
}, {
  name: "Orange Header Table with Light Orange Alternating Rows - Side header same color",
  style: "Grid Table 5 Dark - Accent 6",
  headerClass: "",
  tableClass: "background-color:#f79646;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#fde9d9;color:#000000;",
  format: "partial",
  sideHeader: true
}, {
  name: "Black Table - Whole table is black",
  style: "List Table 5 Dark",
  headerClass: "border:1px solid #000000; border-bottom: none;",
  tableClass: "background-color:#000000;border:1px solid #f2f2f2;color:#ffffff;",
  rowClass: "background-color:#f2f2f2;color:#000000;",
  format: "full",
  sideHeader: true
}, {
  name: "Blue Table - Whole table is blue",
  style: "List Table 5 Dark - Accent 1",
  headerClass: "border:1px solid #4F81BD; border-bottom: none;",
  tableClass: "background-color:#4F81BD;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#dbe5f1;color:#000000;",
  format: "full",
  sideHeader: true
}, {
  name: "Red Table - Whole table is red",
  style: "List Table 5 Dark - Accent 2",
  headerClass: "border:1px solid #c0504d; border-bottom: none;",
  tableClass: "background-color:#c0504d;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#f2dbdb;color:#000000;",
  format: "full",
  sideHeader: true
}, {
  name: "Green Table - Whole table is green",
  style: "List Table 5 Dark - Accent 3",
  headerClass: "border:1px solid #9bbb59; border-bottom: none;",
  tableClass: "background-color:#9bbb59;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#eaf1dd;color:#000000;",
  format: "full",
  sideHeader: true
}, {
  name: "Purple Table - Whole table is purple",
  style: "List Table 5 Dark - Accent 4",
  headerClass: "border:1px solid #8064a2; border-bottom: none;",
  tableClass: "background-color:#8064a2;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#e5dfec;color:#000000;",
  format: "full",
  sideHeader: true
}, {
  name: "Teal Table - Whole table is teal",
  style: "List Table 5 Dark - Accent 5",
  headerClass: "border:1px solid #4bacc6; border-bottom: none;",
  tableClass: "background-color:#4bacc6;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#daeef3;color:#000000;",
  format: "full",
  sideHeader: true
}, {
  name: "Orange Table - Whole table is orange",
  style: "List Table 5 Dark - Accent 6",
  headerClass: "border:1px solid #f79646; border-bottom: none;",
  tableClass: "background-color:#f79646;border:1px solid #d4d4d1ff;color:#ffffff;",
  rowClass: "background-color:#fde9d9;color:#000000;",
  format: "full",
  sideHeader: true
}];
const wordTableStyleLocales = {
  /* ===================== ENGLISH (source of truth) ===================== */
  en: {
    "Table Grid": "Table Grid",
    "Table Grid 1": "Table Grid",
    "Table Grid 2": "Table Grid",
    "Plain Table 1": "Plain Table 1",
    "Grid Table 2": "Grid Table 2",
    "Grid Table 2 - Accent 1": "Grid Table 2 - Accent 1",
    "Grid Table 2 - Accent 2": "Grid Table 2 - Accent 2",
    "Grid Table 2 - Accent 3": "Grid Table 2 - Accent 3",
    "Grid Table 2 - Accent 4": "Grid Table 2 - Accent 4",
    "Grid Table 2 - Accent 5": "Grid Table 2 - Accent 5",
    "Grid Table 2 - Accent 6": "Grid Table 2 - Accent 6",
    "List Table 3": "List Table 3",
    "List Table 3 - Accent 1": "List Table 3 - Accent 1",
    "List Table 3 - Accent 2": "List Table 3 - Accent 2",
    "List Table 3 - Accent 3": "List Table 3 - Accent 3",
    "List Table 3 - Accent 4": "List Table 3 - Accent 4",
    "List Table 3 - Accent 5": "List Table 3 - Accent 5",
    "List Table 3 - Accent 6": "List Table 3 - Accent 6",
    "Grid Table 4": "Grid Table 4",
    "Grid Table 4 - Accent 1": "Grid Table 4 - Accent 1",
    "Grid Table 4 - Accent 2": "Grid Table 4 - Accent 2",
    "Grid Table 4 - Accent 3": "Grid Table 4 - Accent 3",
    "Grid Table 4 - Accent 4": "Grid Table 4 - Accent 4",
    "Grid Table 4 - Accent 5": "Grid Table 4 - Accent 5",
    "Grid Table 4 - Accent 6": "Grid Table 4 - Accent 6",
    "Grid Table 5 Dark": "Grid Table 5 Dark",
    "Grid Table 5 Dark - Accent 1": "Grid Table 5 Dark - Accent 1",
    "Grid Table 5 Dark - Accent 2": "Grid Table 5 Dark - Accent 2",
    "Grid Table 5 Dark - Accent 3": "Grid Table 5 Dark - Accent 3",
    "Grid Table 5 Dark - Accent 4": "Grid Table 5 Dark - Accent 4",
    "Grid Table 5 Dark - Accent 5": "Grid Table 5 Dark - Accent 5",
    "Grid Table 5 Dark - Accent 6": "Grid Table 5 Dark - Accent 6",
    "List Table 5 Dark": "List Table 5 Dark",
    "List Table 5 Dark - Accent 1": "List Table 5 Dark - Accent 1",
    "List Table 5 Dark - Accent 2": "List Table 5 Dark - Accent 2",
    "List Table 5 Dark - Accent 3": "List Table 5 Dark - Accent 3",
    "List Table 5 Dark - Accent 4": "List Table 5 Dark - Accent 4",
    "List Table 5 Dark - Accent 5": "List Table 5 Dark - Accent 5",
    "List Table 5 Dark - Accent 6": "List Table 5 Dark - Accent 6"
  },
  /* ===================== GERMAN ===================== */
  de: {
    "Table Grid": "Tabellenraster",
    "Table Grid 1": "Tabellenraster",
    "Table Grid 2": "Tabellenraster",
    "Plain Table 1": "Einfache Tabelle 1",
    "Plain Table 2": "Einfache Tabelle 2",
    "Plain Table 3": "Einfache Tabelle 3",
    "Plain Table 5": "Einfache Tabelle 5",
    "Grid Table 2": "Gitternetztabelle 2",
    "Grid Table 2 - Accent 1": "Gitternetztabelle 2 – Akzent 1",
    "Grid Table 2 - Accent 2": "Gitternetztabelle 2 – Akzent 2",
    "Grid Table 2 - Accent 3": "Gitternetztabelle 2 – Akzent 3",
    "Grid Table 2 - Accent 4": "Gitternetztabelle 2 – Akzent 4",
    "Grid Table 2 - Accent 5": "Gitternetztabelle 2 – Akzent 5",
    "Grid Table 2 - Accent 6": "Gitternetztabelle 2 – Akzent 6",
    "List Table 2": "Listentabelle 2",
    "List Table 3": "Listentabelle 3",
    "List Table 3 - Accent 1": "Listentabelle 3 – Akzent 1",
    "List Table 3 - Accent 2": "Listentabelle 3 – Akzent 2",
    "List Table 3 - Accent 3": "Listentabelle 3 – Akzent 3",
    "List Table 3 - Accent 4": "Listentabelle 3 – Akzent 4",
    "List Table 3 - Accent 5": "Listentabelle 3 – Akzent 5",
    "List Table 3 - Accent 6": "Listentabelle 3 – Akzent 6",
    "Grid Table 4": "Gitternetztabelle 4",
    "Grid Table 4 - Accent 1": "Gitternetztabelle 4 – Akzent 1",
    "Grid Table 4 - Accent 2": "Gitternetztabelle 4 – Akzent 2",
    "Grid Table 4 - Accent 3": "Gitternetztabelle 4 – Akzent 3",
    "Grid Table 4 - Accent 4": "Gitternetztabelle 4 – Akzent 4",
    "Grid Table 4 - Accent 5": "Gitternetztabelle 4 – Akzent 5",
    "Grid Table 4 - Accent 6": "Gitternetztabelle 4 – Akzent 6",
    "Grid Table 5 Dark": "Gitternetztabelle 5 dunkel",
    "Grid Table 5 Dark - Accent 1": "Gitternetztabelle 5 dunkel  – Akzent 1",
    "Grid Table 5 Dark - Accent 2": "Gitternetztabelle 5 dunkel  – Akzent 2",
    "Grid Table 5 Dark - Accent 3": "Gitternetztabelle 5 dunkel  – Akzent 3",
    "Grid Table 5 Dark - Accent 4": "Gitternetztabelle 5 dunkel  – Akzent 4",
    "Grid Table 5 Dark - Accent 5": "Gitternetztabelle 5 dunkel  – Akzent 5",
    "Grid Table 5 Dark - Accent 6": "Gitternetztabelle 5 dunkel  – Akzent 6",
    "List Table 5 Dark": "Listentabelle 5 dunkel",
    "List Table 5 Dark - Accent 1": "Listentabelle 5 dunkel  – Akzent 1",
    "List Table 5 Dark - Accent 2": "Listentabelle 5 dunkel  – Akzent 2",
    "List Table 5 Dark - Accent 3": "Listentabelle 5 dunkel  – Akzent 3",
    "List Table 5 Dark - Accent 4": "Listentabelle 5 dunkel  – Akzent 4",
    "List Table 5 Dark - Accent 5": "Listentabelle 5 dunkel  – Akzent 5",
    "List Table 5 Dark - Accent 6": "Listentabelle 5 dunkel  – Akzent 6"
  },
  /* ===================== FRENCH ===================== */
  fr: {
    "Table Grid": "Quadrillage",
    "Plain Table 1": "Tableau simple 1",
    "Grid Table 2": "Tableau quadrillé 2",
    "Grid Table 2 - Accent 1": "Tableau quadrillé 2 Accent 1",
    "Grid Table 2 - Accent 2": "Tableau quadrillé 2 Accent 2",
    "Grid Table 2 - Accent 3": "Tableau quadrillé 2 Accent 3",
    "Grid Table 2 - Accent 4": "Tableau quadrillé 2 Accent 4",
    "Grid Table 2 - Accent 5": "Tableau quadrillé 2 Accent 5",
    "Grid Table 2 - Accent 6": "Tableau quadrillé 2 Accent 6",
    "List Table 3": "Tableau de liste 3",
    "List Table 3 - Accent 1": "Tableau de liste 3 Accent 1",
    "List Table 3 - Accent 2": "Tableau de liste 3 Accent 2",
    "List Table 3 - Accent 3": "Tableau de liste 3 Accent 3",
    "List Table 3 - Accent 4": "Tableau de liste 3 Accent 4",
    "List Table 3 - Accent 5": "Tableau de liste 3 Accent 5",
    "List Table 3 - Accent 6": "Tableau de liste 3 Accent 6",
    "Grid Table 4": "Tableau quadrillé 4",
    "Grid Table 4 - Accent 1": "Tableau quadrillé 4 Accent 1",
    "Grid Table 4 - Accent 2": "Tableau quadrillé 4 Accent 2",
    "Grid Table 4 - Accent 3": "Tableau quadrillé 4 Accent 3",
    "Grid Table 4 - Accent 4": "Tableau quadrillé 4 Accent 4",
    "Grid Table 4 - Accent 5": "Tableau quadrillé 4 Accent 5",
    "Grid Table 4 - Accent 6": "Tableau quadrillé 4 Accent 6",
    "Grid Table 5 Dark": "Tableau quadrillé 5 foncé",
    "Grid Table 5 Dark - Accent 1": "Tableau quadrillé 5 foncé Accent 1",
    "Grid Table 5 Dark - Accent 2": "Tableau quadrillé 5 foncé Accent 2",
    "Grid Table 5 Dark - Accent 3": "Tableau quadrillé 5 foncé Accent 3",
    "Grid Table 5 Dark - Accent 4": "Tableau quadrillé 5 foncé Accent 4",
    "Grid Table 5 Dark - Accent 5": "Tableau quadrillé 5 foncé Accent 5",
    "Grid Table 5 Dark - Accent 6": "Tableau quadrillé 5 foncé Accent 6",
    "List Table 5 Dark": "Tableau de liste 5 foncé",
    "List Table 5 Dark - Accent 1": "Tableau de liste 5 foncé Accent 1",
    "List Table 5 Dark - Accent 2": "Tableau de liste 5 foncé Accent 2",
    "List Table 5 Dark - Accent 3": "Tableau de liste 5 foncé Accent 3",
    "List Table 5 Dark - Accent 4": "Tableau de liste 5 foncé Accent 4",
    "List Table 5 Dark - Accent 5": "Tableau de liste 5 foncé Accent 5",
    "List Table 5 Dark - Accent 6": "Tableau de liste 5 foncé Accent 6"
  },
  /* ===================== SPANISH ===================== */
  es: {
    "Table Grid": "Cuadrícula de tabla",
    "Plain Table 1": "Tabla simple 1",
    "Grid Table 2": "Tabla de cuadrícula 2",
    "Grid Table 2 - Accent 1": "Tabla de cuadrícula 2 Acento 1",
    "Grid Table 2 - Accent 2": "Tabla de cuadrícula 2 Acento 2",
    "Grid Table 2 - Accent 3": "Tabla de cuadrícula 2 Acento 3",
    "Grid Table 2 - Accent 4": "Tabla de cuadrícula 2 Acento 4",
    "Grid Table 2 - Accent 5": "Tabla de cuadrícula 2 Acento 5",
    "Grid Table 2 - Accent 6": "Tabla de cuadrícula 2 Acento 6",
    "List Table 3": "Tabla de lista 3",
    "List Table 3 - Accent 1": "Tabla de lista 3 Acento 1",
    "List Table 3 - Accent 2": "Tabla de lista 3 Acento 2",
    "List Table 3 - Accent 3": "Tabla de lista 3 Acento 3",
    "List Table 3 - Accent 4": "Tabla de lista 3 Acento 4",
    "List Table 3 - Accent 5": "Tabla de lista 3 Acento 5",
    "List Table 3 - Accent 6": "Tabla de lista 3 Acento 6",
    "Grid Table 4": "Tabla de cuadrícula 4",
    "Grid Table 4 - Accent 1": "Tabla de cuadrícula 4 Acento 1",
    "Grid Table 4 - Accent 2": "Tabla de cuadrícula 4 Acento 2",
    "Grid Table 4 - Accent 3": "Tabla de cuadrícula 4 Acento 3",
    "Grid Table 4 - Accent 4": "Tabla de cuadrícula 4 Acento 4",
    "Grid Table 4 - Accent 5": "Tabla de cuadrícula 4 Acento 5",
    "Grid Table 4 - Accent 6": "Tabla de cuadrícula 4 Acento 6",
    "Grid Table 5 Dark": "Tabla de cuadrícula 5 oscuro",
    "Grid Table 5 Dark - Accent 1": "Tabla de cuadrícula 5 oscuro Acento 1",
    "Grid Table 5 Dark - Accent 2": "Tabla de cuadrícula 5 oscuro Acento 2",
    "Grid Table 5 Dark - Accent 3": "Tabla de cuadrícula 5 oscuro Acento 3",
    "Grid Table 5 Dark - Accent 4": "Tabla de cuadrícula 5 oscuro Acento 4",
    "Grid Table 5 Dark - Accent 5": "Tabla de cuadrícula 5 oscuro Acento 5",
    "Grid Table 5 Dark - Accent 6": "Tabla de cuadrícula 5 oscuro Acento 6",
    "List Table 5 Dark": "Tabla de lista 5 oscuro",
    "List Table 5 Dark - Accent 1": "Tabla de lista 5 oscuro Acento 1",
    "List Table 5 Dark - Accent 2": "Tabla de lista 5 oscuro Acento 2",
    "List Table 5 Dark - Accent 3": "Tabla de lista 5 oscuro Acento 3",
    "List Table 5 Dark - Accent 4": "Tabla de lista 5 oscuro Acento 4",
    "List Table 5 Dark - Accent 5": "Tabla de lista 5 oscuro Acento 5",
    "List Table 5 Dark - Accent 6": "Tabla de lista 5 oscuro Acento 6"
  }
};

/***/ }),

/***/ "./src/taskpane/draft/draft-functions.ts":
/*!***********************************************!*\
  !*** ./src/taskpane/draft/draft-functions.ts ***!
  \***********************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   applyCustomTextStyleToCell: function() { return /* binding */ applyCustomTextStyleToCell; },
/* harmony export */   chatfooter: function() { return /* binding */ chatfooter; },
/* harmony export */   colorTable: function() { return /* binding */ colorTable; },
/* harmony export */   confirmSwitchChatHistory: function() { return /* binding */ confirmSwitchChatHistory; },
/* harmony export */   copyText: function() { return /* binding */ copyText; },
/* harmony export */   detectTableCase: function() { return /* binding */ detectTableCase; },
/* harmony export */   generateChatHistoryHtml: function() { return /* binding */ generateChatHistoryHtml; },
/* harmony export */   insertLineWithHeadingStyle: function() { return /* binding */ insertLineWithHeadingStyle; },
/* harmony export */   mapImagesToComponentObjects: function() { return /* binding */ mapImagesToComponentObjects; },
/* harmony export */   parseHtmlTableToGrid: function() { return /* binding */ parseHtmlTableToGrid; },
/* harmony export */   removeQuotes: function() { return /* binding */ removeQuotes; },
/* harmony export */   renderSelectedTags: function() { return /* binding */ renderSelectedTags; },
/* harmony export */   resolveWordTableStyle: function() { return /* binding */ resolveWordTableStyle; },
/* harmony export */   selectMatchingBookmarkFromSelection: function() { return /* binding */ selectMatchingBookmarkFromSelection; },
/* harmony export */   svgBase64ToPngBase64: function() { return /* binding */ svgBase64ToPngBase64; },
/* harmony export */   switchModeIcon: function() { return /* binding */ switchModeIcon; },
/* harmony export */   switchToAddTag: function() { return /* binding */ switchToAddTag; },
/* harmony export */   switchToPromptBuilder: function() { return /* binding */ switchToPromptBuilder; },
/* harmony export */   transposeGrid: function() { return /* binding */ transposeGrid; },
/* harmony export */   updateEditorFinalTable: function() { return /* binding */ updateEditorFinalTable; }
/* harmony export */ });
/* harmony import */ var _components_bodyelements__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../components/bodyelements */ "./src/taskpane/components/bodyelements.ts");
/* harmony import */ var _home__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./home */ "./src/taskpane/draft/home.ts");
/* harmony import */ var _components_tablestyles__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../components/tablestyles */ "./src/taskpane/components/tablestyles.ts");
/* harmony import */ var _services_store_service__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../services/store.service */ "./src/taskpane/services/store.service.ts");
/* harmony import */ var _utils_doc_storage__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/doc-storage */ "./src/taskpane/utils/doc-storage.ts");
/* harmony import */ var _utils_fontawesome_icons__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/fontawesome-icons */ "./src/taskpane/utils/fontawesome-icons.ts");






function insertLineWithHeadingStyle(paragraph, line) {
  let builtInStyle = Word.BuiltInStyleName.normal;
  let text = line;
  if (line.startsWith("###### ")) {
    builtInStyle = Word.BuiltInStyleName.heading6;
    text = line.substring(7).trim();
  } else if (line.startsWith("##### ")) {
    builtInStyle = Word.BuiltInStyleName.heading5;
    text = line.substring(6).trim();
  } else if (line.startsWith("#### ")) {
    builtInStyle = Word.BuiltInStyleName.heading4;
    text = line.substring(5).trim();
  } else if (line.startsWith("### ")) {
    builtInStyle = Word.BuiltInStyleName.heading3;
    text = line.substring(4).trim();
  } else if (line.startsWith("## ")) {
    builtInStyle = Word.BuiltInStyleName.heading2;
    text = line.substring(3).trim();
  } else if (line.startsWith("# ")) {
    builtInStyle = Word.BuiltInStyleName.heading1;
    text = line.substring(2).trim();
  }

  // Apply configured default text style if it's the normal style
  try {
    const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
    if (builtInStyle === Word.BuiltInStyleName.normal) {
      const styleName = store.defaultTextStyle || "Normal";
      paragraph.style = styleName;
    } else {
      paragraph.styleBuiltIn = builtInStyle;
    }

    // Apply customized style properties if configured for normal text
    if (builtInStyle === Word.BuiltInStyleName.normal && store.customizedTextStyle && store.customizedTextStyle.properties) {
      const props = store.customizedTextStyle.properties;
      paragraph.font.bold = props.bold;
      paragraph.font.italic = props.italic;
      paragraph.font.underline = props.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
      if (props.fontFamily) {
        paragraph.font.name = props.fontFamily;
      }
      if (props.size) {
        const numericSize = parseFloat(props.size);
        if (!isNaN(numericSize)) {
          paragraph.font.size = numericSize;
        }
      }
      if (props.fontColor) {
        paragraph.font.color = props.fontColor;
      }
      if (props.backgroundColor && props.backgroundColor !== "transparent") {
        try {
          paragraph.font.highlightColor = props.backgroundColor;
        } catch (e) {}
      }
    }
  } catch (e) {
    console.error("Error setting paragraph style, falling back to normal:", e);
    try {
      paragraph.styleBuiltIn = Word.BuiltInStyleName.normal;
    } catch (err) {}
  }
  const regex = /(\*\*(.+?)\*\*)|(\*(.+?)\*)|(_(.+?)_)/g;
  let lastIndex = 0;
  let match;
  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      paragraph.insertText(text.substring(lastIndex, match.index), Word.InsertLocation.end);
    }
    let content = "";
    let bold = false;
    let italic = false;
    let underline = false;
    if (match[1]) {
      content = match[2];
      bold = true;
    } else if (match[3]) {
      content = match[4];
      italic = true;
    } else if (match[5]) {
      content = match[6];
      underline = true;
    }
    const r = paragraph.insertText(content, Word.InsertLocation.end);
    r.font.bold = bold;
    r.font.italic = italic;
    r.font.underline = underline ? Word.UnderlineType.single : Word.UnderlineType.none;
    lastIndex = regex.lastIndex;
  }
  if (lastIndex < text.length) {
    paragraph.insertText(text.substring(lastIndex), Word.InsertLocation.end);
  }
}
function removeQuotes(value) {
  return value ? value.replace(/^"|"$/g, '').replace(/\\n/g, '').replace(/\*\*/g, '').replace(/\\r/g, '') : '';
}
function copyText(text) {
  // Copy text to clipboard logic
  const tempTextArea = document.createElement('textarea');
  tempTextArea.value = text;
  document.body.appendChild(tempTextArea);
  tempTextArea.select();
  document.execCommand('copy');
  document.body.removeChild(tempTextArea);
  (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_0__.toaster)('Copied to clipboard successfully!', 'success');
}
function confirmSwitchChatHistory(onConfirm, onCancel) {
  const chatInput = document.getElementById("chatInput");
  if (chatInput && chatInput.value.trim().length > 0) {
    const container = document.getElementById('confirmation-popup');
    if (container) {
      container.innerHTML = (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_0__.Confirmationpopup)('Are you sure you want to discard changes and proceed?');
      setTimeout(() => {
        const cancelBtn = document.getElementById('confirmation-popup-cancel');
        const confirmBtn = document.getElementById('confirmation-popup-confirm');
        cancelBtn?.addEventListener('click', () => {
          container.innerHTML = '';
          if (onCancel) onCancel();
        });
        confirmBtn?.addEventListener('click', () => {
          container.innerHTML = '';
          onConfirm();
        });
      }, 0);
    } else {
      if (confirm('You have unsaved changes in your chat input. Do you want to switch?')) {
        onConfirm();
      } else {
        if (onCancel) onCancel();
      }
    }
  } else {
    onConfirm();
  }
}
function switchToPromptBuilder() {
  // Remove active class from current tab
  document.querySelector('.nav-link.active')?.classList.remove('active');
  document.querySelector('.tab-pane.show.active')?.classList.remove('show', 'active');

  // Add active class to Prompt Builder tab
  document.getElementById('prompt-tab').classList.add('active');
  document.getElementById('add-prompt-template').classList.add('show', 'active');
}
function switchToAddTag() {
  // Remove active class from current tab
  document.querySelector('.nav-link.active')?.classList.remove('active');
  document.querySelector('.tab-pane.show.active')?.classList.remove('show', 'active');

  // Add active class to Prompt Builder tab
  document.getElementById('tag-tab').classList.add('active');
  document.getElementById('add-tag-body').classList.add('show', 'active');
}
function updateEditorFinalTable(data) {
  const regex = /<TableStart>([\s\S]*?)<TableEnd>/gi;
  let match;
  let tables = [];
  while ((match = regex.exec(data)) !== null) {
    try {
      const parsedContent = JSON.parse(match[1]);
      tables.push(jsonToHtmlTable(parsedContent));
    } catch (error) {
      console.error("Failed to parse JSON:", error, match[1]);
    }
  }
  let tableIndex = 0;
  return data.replace(regex, () => tables[tableIndex++] || "");
}
function jsonToHtmlTable(jsonData) {
  if (!jsonData || Array.isArray(jsonData) && jsonData.length === 0) {
    return '<p>No data available</p>';
  }
  let headers = new Set();
  let rows = [];
  function flattenObject(obj, prefix = "", result = {}) {
    Object.keys(obj).forEach(key => {
      const value = obj[key];
      const newKey = prefix ? `${prefix} > ${key}` : key;
      if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
        flattenObject(value, newKey, result);
      } else if (Array.isArray(value)) {
        result[newKey] = value.map(item => {
          return typeof item === 'object' ? Object.entries(item).map(([k, v]) => `<strong>${k}:</strong> ${v}`).join('<br>') : item;
        }).join('<br>');
      } else {
        result[newKey] = value;
      }
    });
    return result;
  }
  let normalizedData = Array.isArray(jsonData) ? jsonData : [jsonData];

  // Narrative summary tables come back as [{ Label: Value }, ...].
  // Render them as key/value rows instead of one very wide row.
  if (normalizedData.length > 1 && normalizedData.every(item => typeof item === 'object' && item !== null && Object.keys(item).length === 1)) {
    const keyValueRows = normalizedData.map(item => {
      const [key, value] = Object.entries(flattenObject(item))[0] || ["", ""];
      return {
        key,
        value
      };
    }).filter(({
      key,
      value
    }) => {
      const cellValue = value === undefined || value === null ? "" : String(value);
      return key.trim() !== "" || cellValue.trim() !== "";
    });
    if (keyValueRows.length === 0) return '<p>No data available</p>';
    let table = '<table border="1" cellspacing="0" cellpadding="5">';
    keyValueRows.forEach(({
      key,
      value
    }) => {
      const cellValue = value === undefined || value === null ? "" : value;
      table += `<tr><th>${key}</th><td>${cellValue}</td></tr>`;
    });
    table += '</table>';
    return table;
  }
  normalizedData.forEach(item => {
    let flattenedItem = flattenObject(item);
    Object.keys(flattenedItem).forEach(key => headers.add(key));
    rows.push(flattenedItem);
  });
  let table = '<table border="1" cellspacing="0" cellpadding="5">';
  table += '<tr>' + [...headers].map(header => `<th>${header}</th>`).join('') + '</tr>';
  rows.forEach(row => {
    table += '<tr>' + [...headers].map(header => `<td>${row[header] === undefined || row[header] === null ? "" : row[header]}</td>`).join('') + '</tr>';
  });
  table += '</table>';
  return table;
}
function formatChatDate(dateStr) {
  if (!dateStr) return '';
  try {
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return dateStr;
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    const hh = String(date.getHours()).padStart(2, '0');
    const min = String(date.getMinutes()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd} ${hh}:${min}`;
  } catch (e) {
    return dateStr;
  }
}
function generateChatHistoryHtml(chatList) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
  const promptclass = store.theme === 'Dark' ? 'bg-secondary text-light' : 'bg-white text-dark';
  const globalPromptUpdate = store.UserRole.UserRoleEntityAccessList.find(item => item.UserRoleEntity === 'Global Prompt Update');
  return chatList.map((chat, index) => {
    const includeSaveIcon = globalPromptUpdate?.UserRoleAccessID === 3;
    const includeReferenceIcon = true;
    const formattedDate = formatChatDate(chat.CreatedDate);
    return `
      <div class="row chat-entry m-0 p-0">
        <div class="col-md-12 mt-2 p-2">
          <div class="d-flex justify-content-between align-items-start">
            <!-- Prompt Box -->
            <div class="form-control h-34 d-flex align-items-center dynamic-height prompt-text ${promptclass}" style="width: 95%;">
              ${chat.Prompt}
            </div>

            <!-- Icons Stack -->
            <div class="d-flex flex-column align-items-center ms-2">
              <i class="fa fa-copy text-secondary c-pointer mb-2" title="Copy Prompt" id="copyPrompt-${index}"></i>
              ${includeSaveIcon ? `<i class="fa fa-save text-secondary c-pointer mb-2" title="Save Prompt" id="savePrompt-${index}"></i>` : ''}
              <div class="ngb-tooltip d-inline-block">
                <span class="tooltiptext info-tooltip-text">${chat.DocumentInstruction ? `<strong>Document Instruction:</strong> ${chat.DocumentInstruction}<br>` : ''}<strong>Created By:</strong> ${chat.CreatedByName || 'Unknown User'}<br><strong>Date:</strong> ${formattedDate || 'N/A'}</span>
                <i class="fa-solid fa-circle-info text-secondary c-pointer" id="infoPrompt-${index}"></i>
              </div>
            </div>
          </div>
        </div>

        <div class="col-md-12 mb-2 p-2 d-flex">
          <span class="d-flex align-items-baseline w-100">
            <div class="flex-grow-1 c-pointer ai-response-container px-2 pe-3 pt-3 ai-selected-response" id="responseContainer-${index}">
              <input
                class="form-check-input c-pointer me-2 response-checkbox"
                type="checkbox"
                id="checkbox-${index}"
                ${chat.Selected === 1 ? 'checked' : ''}>
              <span id="responseText-${index}">${chat.Response}</span>
              <i class="fa fa-copy text-secondary c-pointer ms-2"
                title="Copy Response"
                id="copyResponse-${index}"></i>
              ${includeReferenceIcon ? (0,_utils_fontawesome_icons__WEBPACK_IMPORTED_MODULE_5__.getIconSvg)(_utils_fontawesome_icons__WEBPACK_IMPORTED_MODULE_5__.faFolderGear, 'text-secondary c-pointer ms-2', '', `title="Open Reference" id="openRefferance-${index}"`) : ''}
            </div>
          </span>
        </div>
      </div>`;
  }).join('');
}
function chatfooter(tag) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
  const promptclass = store.theme === 'Dark' ? 'bg-secondary text-light' : 'bg-white text-dark';
  const tooltipButton = tag.Sources && tag.Sources.length > 0 ? `  <span class="tooltiptext">${tag.Sources}</span>` : '<span class="tooltiptext">Source</span>';
  return ` <textarea class="form-control ${promptclass}"
                      rows="7"
                      id="chatInput"
                      ></textarea>
            <div id="mention-dropdown" class="dropdown-menu"></div>
            <div class="d-flex flex-column align-self-end me-3">
              <button class="btn btn-secondary text-light ms-2 mb-2 ngb-tooltip" id="insertTagButton">
                <span class="tooltiptext">Insert</span>
                <i class="fa fa-plus text-light c-pointer"></i>
              </button>
              <button class="btn btn-secondary text-light ms-2 mb-2 ngb-tooltip" id="promptBuilderButton">
                <span class="tooltiptext">Prompt Builder</span>
                <i class="fa fa-keyboard text-light c-pointer"></i>
              </button>
              <button class="btn btn-secondary text-light ms-2 mb-2 ngb-tooltip" id="changeSourceButton">
                ${tooltipButton}
                <i class="fa fa-file-lines text-light c-pointer"></i>
              </button>

              <button type="submit" class="btn btn-primary bg-primary-clr ms-2 text-white ngb-tooltip" id="sendPromptButton">
                <span class="tooltiptext">Send</span>
                <i class="fa fa-paper-plane text-white"></i>
              </button>
            </div>`;
}
function renderSelectedTags(selectedNames, availableKeys) {
  const badgeWrapper = document.getElementById('tag-badge-wrapper');
  if (badgeWrapper) {
    badgeWrapper.innerHTML = '';
    // Filter out duplicates (case-insensitive)
    const uniqueNames = [...new Set(selectedNames.map(name => name.toLowerCase()))].map(lowerName => selectedNames.find(name => name.toLowerCase() === lowerName));
    const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
    uniqueNames.forEach(name => {
      if (store.mode === 'Summary') {
        let summaryTag;
        if (/^SM\d+$/i.test(name)) {
          summaryTag = store.summaryTagList?.find(k => `sm${k.ID || k.ReportHeadSummaryTagID}`.toLowerCase() === name.toLowerCase());
        } else {
          summaryTag = store.summaryTagList?.find(k => k.Name?.toLowerCase() === name.toLowerCase());
        }
        if (summaryTag?.Name) {
          const badge = document.createElement('span');
          badge.className = 'badge rounded-pill border bg-white text-dark px-3 py-2 shadow-sm d-flex align-items-center badge-clickable';
          badge.style.cursor = 'pointer';
          badge.innerHTML = `${summaryTag.Name} <i class="fa-solid fa-wand-magic-sparkles ms-2 text-muted" aria-label="Summary Tag"></i>`;
          badge.addEventListener('click', async () => {
            const tagId = summaryTag.ID || summaryTag.ReportHeadSummaryTagID;
            if (store.currentChatTagId !== -1 && store.currentChatTagId !== undefined && store.currentChatTagId !== null && String(store.currentChatTagId) === String(tagId)) {
              return;
            }
            confirmSwitchChatHistory(async () => {
              await selectMatchingBookmarkFromSelection(name);
              if (summaryTag) {
                const appBody = document.getElementById('app-body');
                appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';
                (0,_home__WEBPACK_IMPORTED_MODULE_1__.generateCheckboxHistory)(summaryTag, "Summary").then(html => {
                  appBody.innerHTML = html;
                });
              }
            });
          });
          badgeWrapper.appendChild(badge);
        }
      } else {
        let aiTag;
        if (/^ID\d+$/i.test(name)) {
          aiTag = availableKeys.find(mention => mention.AIFlag === 1 && `id${mention.ID}`.toLowerCase() === name.toLowerCase());
        } else {
          aiTag = availableKeys.find(mention => mention.AIFlag === 1 && mention.DisplayName.toLowerCase() === name.toLowerCase());
        }
        if (aiTag?.DisplayName) {
          const badge = document.createElement('span');
          badge.className = 'badge rounded-pill border bg-white text-dark px-3 py-2 shadow-sm d-flex align-items-center badge-clickable';
          badge.style.cursor = 'pointer';
          badge.innerHTML = `${aiTag.DisplayName} ${(0,_utils_fontawesome_icons__WEBPACK_IMPORTED_MODULE_5__.getIconSvg)(_utils_fontawesome_icons__WEBPACK_IMPORTED_MODULE_5__.faMicrochipAi, 'ms-2 text-muted', '', 'aria-label="AI Suggested"')}`;
          badge.addEventListener('click', async () => {
            const tagId = aiTag.ID || aiTag.ReportHeadSummaryTagID;
            if (store.currentChatTagId !== -1 && store.currentChatTagId !== undefined && store.currentChatTagId !== null && String(store.currentChatTagId) === String(tagId)) {
              return;
            }
            confirmSwitchChatHistory(async () => {
              await selectMatchingBookmarkFromSelection(name);
              if (aiTag) {
                const appBody = document.getElementById('app-body');
                appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';
                (0,_home__WEBPACK_IMPORTED_MODULE_1__.generateCheckboxHistory)(aiTag, "AITag").then(html => {
                  appBody.innerHTML = html;
                });
              }
            });
          });
          badgeWrapper.appendChild(badge);
        }
      }
    });
  }
}

// applyThemeClasses and swicthThemeIcon moved to UIService

const modeIconMap = {
  Home: "fa-home",
  Summary: "fa-wand-magic-sparkles",
  Review: "fa-magnifying-glass"
};
function switchModeIcon() {
  const icon = document.getElementById("modeDropdown");
  if (!icon) return;
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
  icon.classList.remove(...Object.values(modeIconMap));
  icon.classList.add(modeIconMap[store.mode]);
  _utils_doc_storage__WEBPACK_IMPORTED_MODULE_4__.DocStorage.setItem("mode", store.mode);
}
async function selectMatchingBookmarkFromSelection(displayName) {
  return Word.run(async context => {
    const selection = context.document.getSelection();
    const bookmarks = selection.getBookmarks(); // ClientResult<string[]>
    await context.sync();
    const targetBookmarkName = bookmarks.value.find(bookmark => {
      const cleanName = bookmark.split('_Split_')[0].replace(/_/g, ' ');
      return cleanName.toLowerCase() === displayName.toLowerCase();
    });
    if (targetBookmarkName) {
      const range = context.document.getBookmarkRangeOrNullObject(targetBookmarkName);
      range.load('isNullObject');
      await context.sync();
      if (!range.isNullObject) {
        range.select(); // Select the entire bookmark
      }
    }
  });
}
function applyCustomTextStyleToCell(cell, store) {
  try {
    const styleName = store.defaultTextStyle || "Normal";
    cell.body.style = styleName;
  } catch (e) {
    console.error("Error setting cell body style:", e);
  }
  if (store.customizedTextStyle && store.customizedTextStyle.properties) {
    const props = store.customizedTextStyle.properties;
    const font = cell.body.font;
    font.bold = props.bold;
    font.italic = props.italic;
    font.underline = props.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
    if (props.fontFamily) {
      font.name = props.fontFamily;
    }
    if (props.size) {
      const numericSize = parseFloat(props.size);
      if (!isNaN(numericSize)) {
        font.size = numericSize;
      }
    }
    if (props.fontColor) {
      font.color = props.fontColor;
    }
    if (props.backgroundColor && props.backgroundColor !== "transparent") {
      try {
        font.highlightColor = props.backgroundColor;
      } catch (e) {}
    }
  }
}
async function colorTable(table, rows, context, isReversed = false) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();

  // ------------------------------------------------------------
  // 1) Copy cell values DOM -> Word + MERGE FIRST COLUMN GROUPS
  // Skip this phase if reversed, as the grid population is handled externally 
  // and merging is disabled for transposed tables.
  // ------------------------------------------------------------
  if (!isReversed) {
    let lastParamRowIndex = -1; // last row index where 1st col had value

    rows.forEach((row, rowIndex) => {
      const cells = Array.from(row.querySelectorAll("td, th"));
      let cellIndex = 0;
      let firstColText = "";
      cells.forEach(cell => {
        const text = cell.innerText.trim();

        // capture first column text
        if (cellIndex === 0) {
          firstColText = text;
        }
        const tableCell = table.getCell(rowIndex, cellIndex);
        tableCell.value = text;
        applyCustomTextStyleToCell(tableCell, store);
        cellIndex++;
      });

      // ✅ Merge first column when empty (skip header row)
      if (rowIndex > 0) {
        if (firstColText) {
          // new group starts
          lastParamRowIndex = rowIndex;
        } else {
          // empty first col = merge with last non-empty parameter row
          if (lastParamRowIndex !== -1) {
            const topCell = table.getCell(lastParamRowIndex, 0);
            const bottomCell = table.getCell(rowIndex, 0);
            topCell.merge(bottomCell);

            // ✅ center align merged parameter cell
            try {
              topCell.verticalAlignment = Word.VerticalAlignment.center;
              topCell.body.paragraphs.getFirst().alignment = Word.Alignment.center;
            } catch (e) {
              // ignore
            }
          }
        }
      }
    });
    await context.sync();
  }

  // ------------------------------------------------------------
  // 2) Load rows for formatting
  // ------------------------------------------------------------
  table.rows.load("items");
  await context.sync();
  table.rows.items.forEach(row => row.cells.load("items"));
  await context.sync();
  table.rows.items.forEach(row => {
    row.cells.items.forEach(cell => {
      applyCustomTextStyleToCell(cell, store);
    });
  });

  // Helper to check if color is dark
  const isColorDark = hexColor => {
    const hex = hexColor.replace("#", "");
    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return luminance < 0.5;
  };
  const applyBoldIfNeeded = (cell, rowIndex, cellIndex) => {
    let weight = "Normal";
    if (store.colorPallete.IsHeaderBold && rowIndex === 0) weight = "Bold";
    if (store.colorPallete.IsSideHeaderBold && cellIndex === 0) weight = "Bold";
    cell.body.font.bold = weight === "Bold";
  };

  // Helper to apply shading, font color, and border to a row or cell
  const applyColor = (cellOrRow, bgColor) => {
    cellOrRow.shadingColor = bgColor;
    try {
      cellOrRow.font.color = isColorDark(bgColor) ? "#FFFFFF" : "#000000";

      // Apply thin light grey border for all sides
      const borderColor = "#D3D3D3"; // light grey
      if (cellOrRow.getBorder) {
        ["Top", "Bottom", "Left", "Right"].forEach(side => {
          cellOrRow.getBorder(side).type = "Single";
          cellOrRow.getBorder(side).color = borderColor;
          cellOrRow.getBorder(side).width = 1; // 1pt thin
        });
      }
    } catch (e) {
      // row objects might not have font directly
    }
  };

  // Determine base table type
  const base = store.tableStyle.split(" - ")[0].trim();

  // ------------------------------------------------------------
  // Plain Table 3
  // ------------------------------------------------------------
  if (base === "Plain Table 3") {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = store.colorPallete.Primary;
        if (rowIndex === 0 || cellIndex === 0) {
          bgColor = store.colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1 ? store.colorPallete.Primary : store.colorPallete.Secondary;
        }
        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
  }

  // ------------------------------------------------------------
  // Plain Table 2
  // ------------------------------------------------------------
  else if (base === "Plain Table 2") {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    const rowCount = table.rows.items.length;
    const firstGapIndex = 0; // gap between row 0 and row 1
    const lastGapIndex = rowCount - 1; // gap between last-2 and last row

    // STEP 2 — Turn OFF all gaps EXCEPT first + last
    table.rows.items.forEach((row, rowIndex) => {
      // rowIndex = gap BELOW this row
      if (rowIndex !== firstGapIndex && rowIndex !== lastGapIndex) {
        const bottom = row.getBorder(Word.BorderLocation.bottom);
        bottom.type = Word.BorderType.none;
      }
    });
    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = store.colorPallete.Primary;
        if (rowIndex === 0) {
          bgColor = store.colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1 ? store.colorPallete.Primary : store.colorPallete.Secondary;
        }
        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
    await context.sync();
  }

  // ------------------------------------------------------------
  // Plain Table 5
  // ------------------------------------------------------------
  else if (base === "Plain Table 5") {
    table.getBorder(Word.BorderLocation.insideVertical).type = Word.BorderType.none;
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = store.colorPallete.Primary;
        if (rowIndex === 0) {
          bgColor = store.colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1 ? store.colorPallete.Primary : store.colorPallete.Secondary;
        }
        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
  }

  // ------------------------------------------------------------
  // Table Grid
  // ------------------------------------------------------------
  else if (base === "Table Grid") {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, i) => {
      const bg = i % 2 === 0 ? store.colorPallete.Header : store.colorPallete.Primary;
      applyColor(row, bg);
      row.cells.items.forEach((cell, cellIndex) => {
        applyBoldIfNeeded(cell, i, cellIndex);
      });
    });
  } else if (base === "Table Grid 2") {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = store.colorPallete.Primary;
        if (cellIndex === 0) {
          // Side header only
          bgColor = store.colorPallete.Header;
        } else {
          // Rest of the cells follow alternating Primary/Secondary logic
          bgColor = rowIndex % 2 === 0 ? store.colorPallete.Primary : store.colorPallete.Secondary;
        }
        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
  } else if (base === "Table Grid 1") {
    table.rows.items.forEach(r => r.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, i) => {
      const bg = i % 2 === 0 ? store.colorPallete.Header : store.colorPallete.Primary;
      applyColor(row, bg);
      row.cells.items.forEach((cell, cellIndex) => {
        applyBoldIfNeeded(cell, i, cellIndex);
      });
    });

    // ✅ ONE full width double border under header row
    const headerRow = table.rows.items[0];
    const bottom = headerRow.getBorder(Word.BorderLocation.bottom);
    bottom.type = Word.BorderType.double;
    bottom.width = 0.5;
    bottom.color = "000000";
    await context.sync();
  }

  // ------------------------------------------------------------
  // Grid Table 4
  // ------------------------------------------------------------
  else if (base.startsWith("Grid Table 4")) {
    const headerRow = table.rows.items[0];
    applyColor(headerRow, store.colorPallete.Header);
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, i) => {
      if (i > 0) {
        const bg = i % 2 === 0 ? store.colorPallete.Secondary : store.colorPallete.Primary;
        applyColor(row, bg);
      }
      row.cells.items.forEach((cell, cellIndex) => {
        applyBoldIfNeeded(cell, i, cellIndex);
      });
    });
  }

  // ------------------------------------------------------------
  // Grid Table 5 Dark
  // ------------------------------------------------------------
  else if (base.startsWith("Grid Table 5 Dark")) {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = store.colorPallete.Primary;
        if (rowIndex === 0 || cellIndex === 0) {
          bgColor = store.colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1 ? store.colorPallete.Primary : store.colorPallete.Secondary;
        }
        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
  }

  // ------------------------------------------------------------
  // List Table 3
  // ------------------------------------------------------------
  else if (base.startsWith("List Table 3")) {
    const headerRow = table.rows.items[0];
    applyColor(headerRow, store.colorPallete.Header);
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, i) => {
      if (i > 0) applyColor(row, store.colorPallete.Primary);
      row.cells.items.forEach((cell, cellIndex) => {
        applyBoldIfNeeded(cell, i, cellIndex);
      });
    });
  }

  // ------------------------------------------------------------
  // List Table 2
  // ------------------------------------------------------------
  else if (base.startsWith("List Table 2")) {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = store.colorPallete.Primary;
        if (rowIndex === 0) {
          bgColor = store.colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1 ? store.colorPallete.Primary : store.colorPallete.Secondary;
        }
        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
  }

  // ------------------------------------------------------------
  // default
  // ------------------------------------------------------------
  else {
    table.rows.items.forEach(row => applyColor(row, store.colorPallete.Primary));
  }
  await context.sync();
}
function mapImagesToComponentObjects(input) {
  if (!input) return [];

  // 1️⃣ Flatten ALL three arrays into a single list
  const flatImages = [...(input.Flowchart || []), ...(input.Graph || []), ...(input.Image || [])];

  // 2️⃣ Map to required structure
  return flatImages.map(img => ({
    Name: img.ImageName,
    DisplayName: img.ImageName,
    EditorValue: img.ImageData,
    UserValue: img.ImageData,
    ComponentKeyDataType: "IMAGE",
    AIFlag: 0,
    IsImage: true
  }));
}
async function svgBase64ToPngBase64(svgBase64) {
  return new Promise((resolve, reject) => {
    try {
      const svgBlob = new Blob([atob(svgBase64.split(',')[1])], {
        type: "image/svg+xml"
      });
      const url = URL.createObjectURL(svgBlob);
      const img = new Image();
      img.crossOrigin = "anonymous";
      img.onload = function () {
        const canvas = document.createElement("canvas");
        canvas.width = img.width || 1200;
        canvas.height = img.height || 800;
        const ctx = canvas.getContext("2d");
        if (!ctx) return reject("Canvas unsupported.");
        ctx.drawImage(img, 0, 0);
        const pngBase64 = canvas.toDataURL("image/png").split(",")[1];
        resolve(pngBase64);
        URL.revokeObjectURL(url);
      };
      img.onerror = () => reject("SVG load failed.");
      img.src = url;
    } catch (err) {
      reject(err);
    }
  });
}
function resolveWordTableStyle(englishStyle) {
  const lang = Office.context.displayLanguage?.toLowerCase().slice(0, 2) || "en";
  return _components_tablestyles__WEBPACK_IMPORTED_MODULE_2__.wordTableStyleLocales[lang]?.[englishStyle] ?? _components_tablestyles__WEBPACK_IMPORTED_MODULE_2__.wordTableStyleLocales.en[englishStyle] ?? 'none';
}

/**
 * Converts an array of HTML table rows into a 2D string grid.
 * Correctly accounts for colspan and rowspan by filling the grid cells.
 */
function parseHtmlTableToGrid(rows) {
  if (rows.length === 0) return [];

  // 1. Calculate max columns considering colspans
  let maxCols = 0;
  rows.forEach(row => {
    let colsInRow = 0;
    Array.from(row.querySelectorAll("td, th")).forEach(cell => {
      colsInRow += parseInt(cell.getAttribute("colspan") || "1", 10);
    });
    if (colsInRow > maxCols) maxCols = colsInRow;
  });
  const rowCount = rows.length;
  const grid = Array.from({
    length: rowCount
  }, () => new Array(maxCols).fill(""));
  const occupied = Array.from({
    length: rowCount
  }, () => new Array(maxCols).fill(false));
  rows.forEach((row, rowIndex) => {
    const cells = Array.from(row.querySelectorAll("td, th"));
    let colIndex = 0;
    cells.forEach(cell => {
      // Find the next available column in the grid
      while (colIndex < maxCols && occupied[rowIndex][colIndex]) {
        colIndex++;
      }
      if (colIndex >= maxCols) return;
      const cellText = Array.from(cell.childNodes).map(node => {
        if (node.nodeType === Node.TEXT_NODE) {
          return node.textContent?.trim() || "";
        } else if (node.nodeType === Node.ELEMENT_NODE) {
          return node.innerText.trim();
        }
        return "";
      }).filter(text => text.length > 0).join(" ");
      const colspan = parseInt(cell.getAttribute("colspan") || "1", 10);
      const rowspan = parseInt(cell.getAttribute("rowspan") || "1", 10);

      // Fill the grid cells covered by this cell's rowspan/colspan
      for (let r = 0; r < rowspan; r++) {
        for (let c = 0; c < colspan; c++) {
          const targetRow = rowIndex + r;
          const targetCol = colIndex + c;
          if (targetRow < rowCount && targetCol < maxCols) {
            // Only put text in the primary cell; others stay empty strings for merging logic
            grid[targetRow][targetCol] = r === 0 && c === 0 ? cellText : "";
            occupied[targetRow][targetCol] = true;
          }
        }
      }
      colIndex += colspan;
    });
  });
  return grid;
}

/**
 * Flips the 2D grid (rows become columns and vice versa).
 */
function transposeGrid(grid) {
  if (grid.length === 0) return [];
  const rows = grid.length;
  const cols = grid[0].length;
  const transposed = Array.from({
    length: cols
  }, () => new Array(rows).fill(""));
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      transposed[c][r] = grid[r][c];
    }
  }
  return transposed;
}
function detectTableCase(data) {
  if (!Array.isArray(data) || data.length === 0) return "UNKNOWN";

  // CASE 1 → every row has exactly 2 elements (key-value)
  if (data.every(row => Array.isArray(row) && row.length === 2)) {
    return "CASE_1";
  }

  // Must have at least header
  if (!Array.isArray(data[0])) return "UNKNOWN";
  const headerLength = data[0].length;

  // all rows same length as header
  const validTable = data.every(row => Array.isArray(row) && row.length === headerLength);
  if (!validTable) return "UNKNOWN";

  // CASE 2 → header + single row
  if (data.length === 2) return "CASE_2";

  // CASE 3 → header + multiple rows
  if (data.length > 2) return "CASE_3";
  return "UNKNOWN";
}

/***/ }),

/***/ "./src/taskpane/draft/draft.api.ts":
/*!*****************************************!*\
  !*** ./src/taskpane/draft/draft.api.ts ***!
  \*****************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   addAiHistory: function() { return /* binding */ addAiHistory; },
/* harmony export */   addGroupKey: function() { return /* binding */ addGroupKey; },
/* harmony export */   checkLoginType: function() { return /* binding */ checkLoginType; },
/* harmony export */   fetchGlossaryTemplate: function() { return /* binding */ fetchGlossaryTemplate; },
/* harmony export */   getAiHistory: function() { return /* binding */ getAiHistory; },
/* harmony export */   getAllClients: function() { return /* binding */ getAllClients; },
/* harmony export */   getAllCustomTables: function() { return /* binding */ getAllCustomTables; },
/* harmony export */   getAllCustomTexts: function() { return /* binding */ getAllCustomTexts; },
/* harmony export */   getAllPromptTemplates: function() { return /* binding */ getAllPromptTemplates; },
/* harmony export */   getGeneralImages: function() { return /* binding */ getGeneralImages; },
/* harmony export */   getPromptTemplateById: function() { return /* binding */ getPromptTemplateById; },
/* harmony export */   getReportById: function() { return /* binding */ getReportById; },
/* harmony export */   getReportHeadImageById: function() { return /* binding */ getReportHeadImageById; },
/* harmony export */   loginUser: function() { return /* binding */ loginUser; },
/* harmony export */   ssoComplete: function() { return /* binding */ ssoComplete; },
/* harmony export */   ssoLogin: function() { return /* binding */ ssoLogin; },
/* harmony export */   updateAiHistory: function() { return /* binding */ updateAiHistory; },
/* harmony export */   updateGroupKey: function() { return /* binding */ updateGroupKey; },
/* harmony export */   updatePromptTemplate: function() { return /* binding */ updatePromptTemplate; }
/* harmony export */ });
/* harmony import */ var _utils_config__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../utils/config */ "./src/taskpane/utils/config.ts");


// api.ts
const baseUrl = {
  toString: () => _utils_config__WEBPACK_IMPORTED_MODULE_0__.CONFIG.dataUrl
};
async function loginUser(organization, username, password) {
  const response = await fetch(`${baseUrl}/api/user/login`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      ClientName: organization,
      Username: username,
      Password: password,
      LoginType: 'ADDIN'
    })
  });
  if (!response.ok) {
    throw new Error('Network response was not ok');
  }
  const data = await response.json();
  return data;
}

// api.ts

async function getReportById(documentID, jwt) {
  const response = await fetch(`${baseUrl}/api/report/id/${documentID}`, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${jwt}`
    }
  });
  if (!response.ok) {
    throw new Error('Network response was not ok');
  }
  const data = await response.json();
  return data;
}
async function getAllClients(userId, jwt) {
  const response = await fetch(`${baseUrl}/api/client/all/${userId}`, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${jwt}`
    }
  });
  if (!response.ok) {
    throw new Error('Network response was not ok');
  }
  const data = await response.json();
  return data;
}
async function getAiHistory(tagId, jwt) {
  const response = await fetch(`${baseUrl}/api/report/ai-history/${tagId}`, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${jwt}`
    }
  });
  if (!response.ok) {
    throw new Error('Network response was not ok');
  }
  const data = await response.json();
  return data;
}
async function updateGroupKey(tag, jwt) {
  const response = await fetch(`${baseUrl}/api/report/head/groupkey`, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(tag)
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  const data = await response.json();
  return data;
}
async function addAiHistory(payload, jwt) {
  const response = await fetch(`${baseUrl}/api/report/ai-history/add`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  const data = await response.json();
  return data;
}
async function updateAiHistory(payload, jwt) {
  const response = await fetch(`${baseUrl}/api/report/ai-history/update`, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  const data = await response.json();
  return data;
}
async function fetchGlossaryTemplate(clientId, bodyText, jwt) {
  const response = await fetch(`${baseUrl}/api/glossary-template/client-id/${clientId}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(bodyText)
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  const data = await response.json();
  return data;
}
async function addGroupKey(payload, jwt) {
  const response = await fetch(`${baseUrl}/api/report/group-key/add`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  const data = await response.json();
  return data;
}
async function getAllPromptTemplates(jwt) {
  const response = await fetch(`${baseUrl}/api/prompt-template/all`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  const data = await response.json();
  return data;
}
async function getPromptTemplateById(id, jwt) {
  const response = await fetch(`${baseUrl}/api/prompt-template/${id}/data`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  const data = await response.json();
  return data;
}
async function updatePromptTemplate(payload, jwt) {
  const response = await fetch(`${baseUrl}/api/groupkey/update-prompt`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  const data = await response.json();
  return data;
}
async function getAllCustomTables(jwt) {
  const response = await fetch(`${baseUrl}/api/custom-table/all`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  const data = await response.json();
  return data;
}
async function getGeneralImages(jwt) {
  const response = await fetch(`${baseUrl}/api/image/general`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  const data = await response.json();
  return data;
}
async function getReportHeadImageById(id, jwt) {
  const response = await fetch(`${baseUrl}/api/image/report-head/${id}`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  const data = await response.json();
  return data;
}
;
async function getAllCustomTexts(jwt) {
  const response = await fetch(`${baseUrl}/api/custom-text/all`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  const data = await response.json();
  return data;
}
async function checkLoginType(organization, username) {
  const response = await fetch(`${baseUrl}/api/user/check-login-type`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      ClientName: organization,
      Username: username,
      LoginType: 'ADDIN'
    })
  });
  if (!response.ok) {
    throw new Error('Network response was not ok');
  }
  const data = await response.json();
  return data;
}
async function ssoLogin(organization, username) {
  const response = await fetch(`${baseUrl}/api/user/sso-login`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      ClientName: organization,
      Username: username,
      LoginType: 'ADDIN'
    })
  });
  if (!response.ok) {
    throw new Error('Network response was not ok');
  }
  const data = await response.json();
  return data;
}
async function ssoComplete(key) {
  const response = await fetch(`${baseUrl}/api/user/sso-complete`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      key: key
    })
  });
  if (!response.ok) {
    throw new Error('Network response was not ok');
  }
  const data = await response.json();
  return data;
}

/***/ }),

/***/ "./src/taskpane/draft/home.ts":
/*!************************************!*\
  !*** ./src/taskpane/draft/home.ts ***!
  \************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   generateCheckboxHistory: function() { return /* binding */ generateCheckboxHistory; },
/* harmony export */   getDateTimeStamp: function() { return /* binding */ getDateTimeStamp; },
/* harmony export */   initializeAIHistoryEvents: function() { return /* binding */ initializeAIHistoryEvents; },
/* harmony export */   insertTagPrompt: function() { return /* binding */ insertTagPrompt; },
/* harmony export */   jumpToNextBookmarkOfTag: function() { return /* binding */ jumpToNextBookmarkOfTag; },
/* harmony export */   loadHomepage: function() { return /* binding */ loadHomepage; },
/* harmony export */   openAITag: function() { return /* binding */ openAITag; },
/* harmony export */   openPromptBuilderModal: function() { return /* binding */ openPromptBuilderModal; },
/* harmony export */   replaceMention: function() { return /* binding */ replaceMention; },
/* harmony export */   setupPromptBuilderUI: function() { return /* binding */ setupPromptBuilderUI; }
/* harmony export */ });
/* harmony import */ var _draft_api__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./draft.api */ "./src/taskpane/draft/draft.api.ts");
/* harmony import */ var _draft_functions__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./draft-functions */ "./src/taskpane/draft/draft-functions.ts");
/* harmony import */ var _taskpane__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../taskpane */ "./src/taskpane/taskpane.ts");
/* harmony import */ var _services_store_service__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../services/store.service */ "./src/taskpane/services/store.service.ts");
/* harmony import */ var _services_ai_service__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../services/ai.service */ "./src/taskpane/services/ai.service.ts");
/* harmony import */ var _components_bodyelements__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../components/bodyelements */ "./src/taskpane/components/bodyelements.ts");
/* harmony import */ var _summary_summary__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../summary/summary */ "./src/taskpane/summary/summary.ts");
/* harmony import */ var _services_summary_service__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../services/summary.service */ "./src/taskpane/services/summary.service.ts");
/* harmony import */ var _summary_summary_api__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../summary/summary.api */ "./src/taskpane/summary/summary.api.ts");
/* harmony import */ var _utils_doc_storage__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../utils/doc-storage */ "./src/taskpane/utils/doc-storage.ts");
/* harmony import */ var _utils_fontawesome_icons__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../utils/fontawesome-icons */ "./src/taskpane/utils/fontawesome-icons.ts");











let preview = '';
function loadHomepage(availableKeys) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
  const searchBoxClass = store.theme === 'Dark' ? 'bg-secondary text-light' : 'bg-white text-dark';
  document.getElementById('app-body').innerHTML = `
    <div class="container pt-3">
        <div class="d-flex justify-content-end px-2">
            <div class="dropdown">
                <button class="btn btn-default dropdown-toggle" type="button" data-bs-toggle="dropdown" aria-expanded="false">
                    Action
                </button>
                <ul class="dropdown-menu">
                    <li>
                        <a class="dropdown-item" href="#" id="add-btn-tag">
                            <i class="fa fa-plus me-2" aria-hidden="true"></i> Add
                        </a>
                    </li>
                    <li>
                        <a class="dropdown-item" href="#" id="apply-btn-tag">
                            <i class="fa-solid fa-circle-check me-2"></i> Apply
                        </a>
                    </li>

                </ul>
            </div>
        </div>

        <div class="form-group px-2 pt-2">
            <input type="text" id="search-box" class="form-control ${searchBoxClass}" placeholder="Search Tags..." autocomplete="off" />
        </div>

        <ul id="suggestion-list" class="list-group mt-2 px-2"></ul>
        
        <div id="tags-in-selected-text" class="mt-2 px-2 selected-text-box d-none">
            <label class="form-label mb-2 fw-bold">Tags in Selected Text</label>
            <div class="tag-panel d-flex flex-wrap gap-2" id="tag-badge-wrapper"></div>
        </div>
    </div>`;
  const searchBox = document.getElementById('search-box');
  const suggestionList = document.getElementById('suggestion-list');
  function updateSuggestions() {
    const searchTerm = searchBox.value.trim().toLowerCase();
    suggestionList.replaceChildren();
    if (searchTerm === '') {
      suggestionList.innerHTML = '';
      return;
    }
    const filteredMentions = availableKeys.filter(m => m.DisplayName.toLowerCase().includes(searchTerm));

    // Split groups
    const nonAITags = filteredMentions.filter(m => m.AIFlag === 0);
    const aiTags = filteredMentions.filter(m => m.AIFlag === 1);

    // Further split non-AI tags into: TEXT + IMAGE
    const propertiesTags = nonAITags.filter(m => m.ComponentKeyDataType === "TEXT" || m.ComponentKeyDataType === "TABLE");
    const imageTags = nonAITags.filter(m => m.ComponentKeyDataType === "IMAGE" && m.IsImage);
    const createSection = (labelText, mentions, isAISection = false, isImageSection = false) => {
      if (mentions.length === 0) return;
      const themeClasses = store.theme === 'Dark' ? {
        itemClass: 'bg-dark text-light list-hover-dark',
        labelClass: 'bg-dark text-light'
      } : {
        itemClass: 'bg-light text-dark list-hover-light',
        labelClass: 'bg-light text-dark'
      };
      const label = document.createElement('li');
      label.className = `list-group-item fw-bold text-secondary ${themeClasses.labelClass}`;
      label.textContent = labelText;
      suggestionList.appendChild(label);
      mentions.forEach(mention => {
        const listItem = document.createElement('li');
        listItem.className = `list-group-item list-group-item-action ${themeClasses.itemClass}`;

        // ICON LOGIC
        let icon = `<i class="fa-solid fa-layer-group text-muted me-2"></i>`; // default (TEXT)
        if (isAISection) icon = (0,_utils_fontawesome_icons__WEBPACK_IMPORTED_MODULE_10__.getIconSvg)(_utils_fontawesome_icons__WEBPACK_IMPORTED_MODULE_10__.faMicrochipAi, 'text-muted me-2');
        if (isImageSection) icon = `<i class="fa-solid fa-image text-muted me-2"></i>`;
        listItem.innerHTML = `${icon} ${mention.DisplayName}`;
        listItem.onclick = () => {
          if (isAISection) {
            const tagId = mention.ID || mention.ReportHeadSummaryTagID;
            if (store.currentChatTagId !== -1 && store.currentChatTagId !== undefined && store.currentChatTagId !== null && String(store.currentChatTagId) === String(tagId)) {
              return;
            }
            (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.confirmSwitchChatHistory)(() => {
              const appBody = document.getElementById('app-body');
              appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';
              generateCheckboxHistory(mention, "AITag").catch(() => appBody.innerHTML = '<div class="text-danger p-2">Error loading data</div>').then(html => {
                appBody.innerHTML = html;
              });
            });
          } else {
            // Properties + Images behave same
            replaceMention(mention, mention.ComponentKeyDataType);
            suggestionList.replaceChildren();
          }
        };
        suggestionList.appendChild(listItem);
      });
    };

    // Render in desired order
    createSection('Properties', propertiesTags);
    createSection('AI Tags', aiTags, true);
    createSection('Images', imageTags, false, true);
  }
  if (store.selectedNames.length > 0) {
    const badgeWrapper = document.getElementById('tags-in-selected-text');
    badgeWrapper.classList.remove('d-none');
    badgeWrapper.classList.add('d-block');
    (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.renderSelectedTags)(store.selectedNames, availableKeys);
  }

  // Add input event listener to the search box
  let debounceTimeout;
  searchBox.addEventListener('input', () => {
    clearTimeout(debounceTimeout);
    debounceTimeout = setTimeout(updateSuggestions, 300); // Delay input handling by 300ms
  });
  document.getElementById('add-btn-tag').addEventListener('click', () => {
    if (!store.isPendingResponse) {
      (0,_taskpane__WEBPACK_IMPORTED_MODULE_2__.addGenAITags)();
    }
  });
  document.getElementById('apply-btn-tag').addEventListener('click', () => {
    if (!store.isPendingResponse) {
      (0,_taskpane__WEBPACK_IMPORTED_MODULE_2__.applyTagFn)();
    }
  });
}
async function replaceMention(word, type) {
  return Word.run(async context => {
    try {
      const selection = context.document.getSelection();
      await context.sync();
      if (!selection) {
        throw new Error('Selection is invalid or not found.');
      }
      let newSelection = selection;
      if (type === 'TABLE') {
        const parser = new DOMParser();
        const doc = parser.parseFromString(word.EditorValue, 'text/html');
        const bodyNodes = Array.from(doc.body.childNodes);
        await context.sync();
        for (const node of bodyNodes) {
          if (node.nodeType === Node.TEXT_NODE) {
            let textContent = node.textContent?.trim();
            if (textContent) {
              textContent = textContent.replace(/\n- /g, "\n• ");
              textContent.split('\n').forEach(line => {
                if (line.trim()) {
                  (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.insertLineWithHeadingStyle)(selection, line);
                }
              });
            }
          } else if (node.nodeType === Node.ELEMENT_NODE) {
            const element = node;
            if (element.tagName.toLowerCase() === 'table') {
              const rows = Array.from(element.querySelectorAll('tr'));
              if (rows.length === 0) {
                selection.insertParagraph("[Empty Table]", Word.InsertLocation.before);
                continue;
              }
              let grid = (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.parseHtmlTableToGrid)(rows);
              const tableCase = (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.detectTableCase)(grid);
              const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
              const base = store.tableStyle.split(" - ")[0].trim();
              if (base === 'Table Grid 2') {
                store.isReversed = true;
              } else {
                store.isReversed = false;
              }
              if (store.isReversed && tableCase !== "CASE_1") {
                grid = (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.transposeGrid)(grid);
              }
              const numRows = grid.length;
              const numCols = grid[0]?.length || 0;
              const paragraph = selection.insertParagraph("", Word.InsertLocation.before);
              await context.sync();
              const table = paragraph.insertTable(numRows, numCols, Word.InsertLocation.after);
              const resolvedTableStyle = (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.resolveWordTableStyle)(store.tableStyle);
              if (resolvedTableStyle !== 'none') {
                table.style = resolvedTableStyle;
              }
              await context.sync();

              // Population/Merging logic
              if (!store.isReversed) {
                if (!store.colorPallete.Customize) {
                  grid.forEach((row, rowIndex) => {
                    row.forEach((cellValue, cellIndex) => {
                      const tableCell = table.getCell(rowIndex, cellIndex);
                      tableCell.value = cellValue;
                      (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.applyCustomTextStyleToCell)(tableCell, store);
                    });
                  });

                  // Vertical merging logic for 1st column (standard view)
                  let lastParamRowIndex = -1;
                  grid.forEach((row, rowIndex) => {
                    if (rowIndex === 0) return; // skip header
                    const firstColText = row[0];
                    if (firstColText) {
                      lastParamRowIndex = rowIndex;
                    } else if (lastParamRowIndex !== -1) {
                      const topCell = table.getCell(lastParamRowIndex, 0);
                      const bottomCell = table.getCell(rowIndex, 0);
                      topCell.merge(bottomCell);
                      try {
                        topCell.verticalAlignment = Word.VerticalAlignment.center;
                        topCell.body.paragraphs.getFirst().alignment = Word.Alignment.center;
                      } catch (e) {}
                    }
                  });
                }
              } else {
                // Manual population for transposed table
                grid.forEach((row, rowIndex) => {
                  row.forEach((cellValue, cellIndex) => {
                    const tableCell = table.getCell(rowIndex, cellIndex);
                    tableCell.value = cellValue;
                    (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.applyCustomTextStyleToCell)(tableCell, store);
                  });
                });
              }

              // Styling logic (always call if Customize, now passing isReversed)
              if (store.colorPallete.Customize) {
                await (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.colorTable)(table, rows, context, store.isReversed);
              }
              newSelection = table.getCell(0, 0); // Set the cursor to the start of the table
            } else {
              let elementText = element.innerText.trim();
              if (elementText) {
                elementText = elementText.replace(/\n- /g, "\n• ");
                elementText.split('\n').forEach(line => {
                  if (line.trim()) {
                    (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.insertLineWithHeadingStyle)(selection, line);
                  }
                });
              }
              newSelection = selection; // If it's not a table, just use the existing selection.
            }
          }
        }
      } else if (type === "IMAGE") {
        let base64Image = word.EditorValue;
        if (base64Image.startsWith("data:image/svg+xml")) {
          // Convert SVG → PNG
          base64Image = await (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.svgBase64ToPngBase64)(base64Image);
        } else if (base64Image.startsWith("data:image")) {
          base64Image = base64Image.split(",")[1]; // strip prefix
        }
        selection.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.replace);
        newSelection = selection;
      } else {
        if (word.EditorValue === '' || word.IsApplied) {
          selection.insertParagraph(`#${word.DisplayName}#`, Word.InsertLocation.before);
        } else {
          let content = (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.removeQuotes)(word.EditorValue);
          let lines = content.split(/\r?\n/); // Handle both \r\n and \n
          lines.forEach(line => {
            selection.insertParagraph(line, Word.InsertLocation.before);
          });
        }
        newSelection = selection; // After inserting the text, set selection to it.
      }

      // Move the cursor to the next line after content insertion
      const nextLineParagraph = selection.insertParagraph("", Word.InsertLocation.after);
      await context.sync();

      // Set the new cursor position after content
      newSelection = nextLineParagraph;
      selection.select(); // Select the new paragraph where the cursor will be
      await context.sync();
    } catch (error) {
      console.error('Detailed error:', error);
    }
  });
}
async function openAITag(tag) {
  tag.ReportHeadAIHistoryList.forEach(historyList => {
    historyList.Response = (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.removeQuotes)(historyList.Response);
    tag.FilteredReportHeadAIHistoryList.unshift(historyList);
  });
}
async function generateCheckboxHistory(tag, type) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
  store.currentChatTagId = tag.ID || tag.ReportHeadSummaryTagID;
  _utils_doc_storage__WEBPACK_IMPORTED_MODULE_9__.DocStorage.setItem("currentChatTagId", String(store.currentChatTagId));
  var skipFetch = false;
  if ((!tag.FilteredReportHeadAIHistoryList || tag.FilteredReportHeadAIHistoryList.length === 0) && !skipFetch) {
    if (type !== 'Summary') {
      await _services_ai_service__WEBPACK_IMPORTED_MODULE_4__.AIService.fetchAIHistory(tag);
    } else {
      await _services_summary_service__WEBPACK_IMPORTED_MODULE_7__.summaryService.fetchSummaryAIHistory(tag);
    }
  }
  const history = tag.FilteredReportHeadAIHistoryList;
  const chat = history.find(item => item.Selected === 1);
  const finalResponse = chat.FormattedResponse ? '\n' + (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.updateEditorFinalTable)(chat.FormattedResponse) : chat.Response;
  tag.ComponentKeyDataType = chat.FormattedResponse ? 'TABLE' : 'TEXT';
  tag.UserValue = finalResponse;
  tag.EditorValue = finalResponse;
  tag.text = finalResponse;
  if (history.length === 0) {
    return '<div>No AI history available.</div>';
  }

  // Check current theme
  const isDark = store.theme === 'Dark';
  const closeBtnClass = isDark ? 'fa-solid fa-circle-xmark bg-dark text-light' : 'fa-solid fa-circle-xmark bg-light text-dark';
  const jumpBtnColorClass = isDark ? 'text-light' : 'text-dark';
  const headerBgClass = isDark ? 'bg-dark text-light' : 'bg-white text-dark';
  const DisplayName = type === 'Summary' ? tag.Name : tag.DisplayName;
  const closeBar = `
    <div class="chat-header sticky-top ${headerBgClass} z-3">
        <div class="d-flex justify-content-between align-items-start px-3 pt-3 pb-1">
            <div class="d-flex align-items-start flex-grow-1" style="max-width: calc(100% - 50px);">
                ${(0,_utils_fontawesome_icons__WEBPACK_IMPORTED_MODULE_10__.getIconSvg)(_utils_fontawesome_icons__WEBPACK_IMPORTED_MODULE_10__.faMicrochipAi, 'text-muted me-2 mt-1', 'font-size: 13px;')}
                <span class="fw-bold" style="font-size: 13px; line-height: 1.4; letter-spacing: 0.3px;">${DisplayName}</span>
            </div>
            <div class="d-flex align-items-center ms-2" style="margin-top: 2px;">
                <button id="jump-to-next-tag" class="btn btn-sm p-0 me-2 border-0 bg-transparent ${jumpBtnColorClass} c-pointer" title="Jump to next replaced instance" style="display: inline-flex; align-items: center; justify-content: center; transition: transform 0.2s ease;">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                        <circle cx="12" cy="12" r="10" />
                        <path d="M 15 14 L 15 12.5 A 2.5 2.5 0 0 0 12.5 10 L 9 10" />
                        <polyline points="11 8 9 10 11 12" />
                    </svg>
                </button>
                <div class="c-pointer d-inline-flex align-items-center justify-content-center" id="close-btn-tag">
                    <i class="${closeBtnClass}" id="close-ai-window" style="font-size: 13px;"></i>
                </div>
            </div>
        </div>
        <hr class="mt-2 mb-1 mx-3">
    </div>
    `;
  const chatBody = `
        <div class="chat-body flex-grow-1 overflow-auto">
            ${(0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.generateChatHistoryHtml)(history)}
        </div>
    `;
  const chatFooterHtml = `
        <div class="d-flex align-items-end justify-content-end chatbox p-2" id="chatFooter">
            ${(0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.chatfooter)(tag)}
        </div>
    `;
  const horizontalLoader = String(tag.Status) === "0" ? '<div class="horizontal-loader"></div>' : '';
  initializeAIHistoryEvents(tag, store.jwt, store.availableKeys, type);
  return `${closeBar}${chatBody}${horizontalLoader}${chatFooterHtml}`;
}
async function setupPromptBuilderUI(container, promptBuilderList) {
  // Static template and field definitions
  let preview = '';
  let templateText = '';

  // Field configs (can be extended)
  let fieldsList = [];

  // Create the form container
  // Create the form container
  container.innerHTML = `
  <div class="form-group mb-3 p-3 pt-0">
    <label class='form-label'><span class="text-danger">*</span> Prompt Builder Template</label>
    <select id="promptBuilderTemplate" class="form-control">
      <option value="" disabled selected>Select a template</option>
    </select>
    <div id="templateError" class="invalid-feedback d-none">Type is required.</div>
  </div>

  <div id="fieldsContainer"></div>

  <div class="form-group mb-3 p-3 pt-0" id="previewContainer" style="display: none;">
    <label class="mb-2">Preview</label>
    <div id="preview" class="form-control"></div>
  </div>

  <div class="d-flex justify-content-between px-3 align-items-center mt-3">
    <span id="resetBtn" class="text-primary fw-bold" style="cursor: pointer;">Reset</span>
    <button id="applyBtn" class="btn btn-primary text-white" disabled>Insert</button>
  </div>
`;

  // Element references
  const templateSelect = container.querySelector('#promptBuilderTemplate');
  const applyBtn = container.querySelector('#applyBtn');
  const resetBtn = container.querySelector('#resetBtn');
  const previewDiv = container.querySelector('#preview');
  const fieldsContainer = container.querySelector('#fieldsContainer');
  const previewContainer = container.querySelector('#previewContainer');
  const templateError = container.querySelector('#templateError');

  // Populate template dropdown
  promptBuilderList.forEach(item => {
    const option = document.createElement('option');
    option.value = item.ID.toString();
    option.textContent = item.Name;
    templateSelect.appendChild(option);
  });
  templateSelect.addEventListener('change', async () => {
    const templateId = templateSelect.value;
    const jwt = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_9__.DocStorage.getItem('token') || '';
    const data = await (0,_draft_api__WEBPACK_IMPORTED_MODULE_0__.getPromptTemplateById)(templateId, jwt);
    if (data.Status && data.Data) {
      fieldsList = data.Data;
      preview = promptBuilderList.find(item => item.ID.toString() === templateId).Template;
      templateText = promptBuilderList.find(item => item.ID.toString() === templateId).Template;
    }
    if (!templateId) {
      templateError.classList.remove('d-none');
      return;
    }
    templateError.classList.add('d-none');
    renderFields();
    updatePreview();
  });
  function renderFields() {
    fieldsContainer.innerHTML = '';
    fieldsList.forEach(field => {
      const div = document.createElement('div');
      div.className = 'form-group mb-3 p-3 pt-0';
      const label = document.createElement('label');
      label.textContent = field.Label;
      div.appendChild(label);
      if (field.Type === 1) {
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'form-control';
        input.id = field.Label;
        input.addEventListener('input', replaceKeywordsManually);
        div.appendChild(input);
      } else if (field.Type === 2) {
        const select = document.createElement('select');
        select.className = 'form-control';
        select.id = field.Label;
        field.PromptTemplateOptionList.forEach(opt => {
          const option = document.createElement('option');
          option.value = opt.Text;
          option.textContent = opt.Option;
          select.appendChild(option);
        });
        select.addEventListener('change', replaceKeywordsManually);
        div.appendChild(select);
      }
      fieldsContainer.appendChild(div);
    });
  }
  function replaceKeywordsManually() {
    const keywordMap = {};
    fieldsList.forEach(field => {
      const id = field.Label;
      const keyword = `#${id}#`;
      let value = '';
      const element = document.getElementById(id);
      if (element) {
        value = element instanceof HTMLInputElement || element instanceof HTMLSelectElement ? element.value : '';
      }
      keywordMap[keyword] = value ? value : keyword;
    });
    let insertValue = templateText;
    for (const [keyword, value] of Object.entries(keywordMap)) {
      insertValue = insertValue.replace(new RegExp(keyword, 'g'), value);
    }
    preview = insertValue;
    previewDiv.textContent = preview;
    previewContainer.style.display = preview ? 'block' : 'none';
    applyBtn.disabled = preview === '';
  }
  function updatePreview() {
    replaceKeywordsManually();
  }
  function resetForm() {
    // Reset only the dynamic field values
    fieldsList.forEach(field => {
      const element = document.getElementById(field.Label);
      if (element) {
        if (element instanceof HTMLInputElement) {
          element.value = '';
        } else if (element instanceof HTMLSelectElement) {
          element.selectedIndex = 0; // optional: reset to first option
        }
      }
    });

    // Clear preview
    previewDiv.textContent = templateText;
    preview = templateText;
  }
  function applyPrompt() {
    if (!preview) return;
    const promptTextarea = document.getElementById('prompt');
    if (promptTextarea) {
      promptTextarea.value = preview;
      (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.switchToAddTag)();
    }
  }
  resetBtn.addEventListener('click', resetForm);
  applyBtn.addEventListener('click', applyPrompt);
}
async function insertTagPrompt(tag, type = "AITag") {
  return Word.run(async context => {
    try {
      const selection = context.document.getSelection();
      await context.sync();
      if (!selection) throw new Error("Invalid selection");

      /* --------------------------------------------------
         1️⃣ Create invisible anchor at cursor
      -------------------------------------------------- */
      const anchorChar = selection.insertText("\u200B",
      // zero-width space
      Word.InsertLocation.replace);
      await context.sync();
      let cursor = anchorChar.getRange();
      let bookmarkStart = null;
      let bookmarkEnd = null;
      const include = r => {
        if (!bookmarkStart) {
          bookmarkStart = r.getRange("Start");
        }
        bookmarkEnd = r.getRange("End");
      };

      /* --------------------------------------------------
         2️⃣ Insert content
      -------------------------------------------------- */
      if (tag.ComponentKeyDataType === "TABLE") {
        const parser = new DOMParser();
        const doc = parser.parseFromString(tag.EditorValue, "text/html");
        const bodyNodes = Array.from(doc.body.childNodes);
        for (const node of bodyNodes) {
          // TEXT NODE
          if (node.nodeType === Node.TEXT_NODE) {
            let txt = node.textContent?.trim();
            if (!txt) continue;
            txt = txt.replace(/\n- /g, "\n• ");
            for (const line of txt.split(/\r?\n/)) {
              if (!line.trim()) {
                const p = cursor.insertParagraph("", Word.InsertLocation.after);
                (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.insertLineWithHeadingStyle)(p, "");
                include(p.getRange());
                cursor = p.getRange();
                continue;
              }
              const p = cursor.insertParagraph("", Word.InsertLocation.after);
              (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.insertLineWithHeadingStyle)(p, line);
              include(p.getRange());
              cursor = p.getRange();
            }
          }

          // ELEMENT NODE
          else if (node.nodeType === Node.ELEMENT_NODE) {
            const el = node;

            // TABLE
            if (el.tagName.toLowerCase() === "table") {
              const rows = Array.from(el.querySelectorAll("tr"));
              if (!rows.length) continue;
              let grid = (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.parseHtmlTableToGrid)(rows);
              const tableCase = (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.detectTableCase)(grid);
              const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
              const base = store.tableStyle.split(" - ")[0].trim();
              if (base === 'Table Grid 2') {
                store.isReversed = true;
              } else {
                store.isReversed = false;
              }
              if (store.isReversed && tableCase !== 'CASE_1') {
                grid = (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.transposeGrid)(grid);
              }
              const numRows = grid.length;
              const numCols = grid[0]?.length || 0;
              const p = cursor.insertParagraph("", Word.InsertLocation.after);
              const table = p.insertTable(numRows, numCols, Word.InsertLocation.after);
              const resolvedTableStyle = (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.resolveWordTableStyle)(store.tableStyle);
              if (resolvedTableStyle !== 'none') {
                table.style = resolvedTableStyle;
              }

              // Population/Merging logic
              if (!store.isReversed) {
                if (!store.colorPallete.Customize) {
                  grid.forEach((rowGrid, rowIndex) => {
                    rowGrid.forEach((cellValue, cellIndex) => {
                      const tableCell = table.getCell(rowIndex, cellIndex);
                      tableCell.value = cellValue;
                      (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.applyCustomTextStyleToCell)(tableCell, store);
                    });
                  });

                  // Vertical merging logic for 1st column (standard view)
                  let lastParamRowIndex = -1;
                  grid.forEach((rowGrid, rowIndex) => {
                    if (rowIndex === 0) return; // skip header
                    const firstColText = rowGrid[0];
                    if (firstColText) {
                      lastParamRowIndex = rowIndex;
                    } else if (lastParamRowIndex !== -1) {
                      const topCell = table.getCell(lastParamRowIndex, 0);
                      const bottomCell = table.getCell(rowIndex, 0);
                      topCell.merge(bottomCell);
                      try {
                        topCell.verticalAlignment = Word.VerticalAlignment.center;
                        topCell.body.paragraphs.getFirst().alignment = Word.Alignment.center;
                      } catch (e) {}
                    }
                  });
                }
              } else {
                // Manual population for transposed table
                grid.forEach((rowGrid, rowIndex) => {
                  rowGrid.forEach((cellValue, cellIndex) => {
                    const tableCell = table.getCell(rowIndex, cellIndex);
                    tableCell.value = cellValue;
                    (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.applyCustomTextStyleToCell)(tableCell, store);
                  });
                });
              }

              // Styling logic
              if (store.colorPallete.Customize) {
                await (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.colorTable)(table, rows, context, store.isReversed);
              }
              include(table.getRange());
              cursor = table.getRange();
            }

            // OTHER HTML ELEMENTS
            else {
              let txt = el.innerText?.trim();
              if (!txt) continue;
              txt = txt.replace(/\n- /g, "\n• ");
              for (const line of txt.split(/\r?\n/)) {
                if (!line.trim()) {
                  const p = cursor.insertParagraph("", Word.InsertLocation.after);
                  (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.insertLineWithHeadingStyle)(p, "");
                  include(p.getRange());
                  cursor = p.getRange();
                  continue;
                }
                const p = cursor.insertParagraph("", Word.InsertLocation.after);
                (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.insertLineWithHeadingStyle)(p, line);
                include(p.getRange());
                cursor = p.getRange();
              }
            }
          }
        }
      }

      // NON-TABLE CONTENT
      else {
        const txt = tag.EditorValue.replace(/\n- /g, "\n• ").trim();
        for (const line of txt.split(/\r?\n/)) {
          if (!line.trim()) {
            const p = cursor.insertParagraph("", Word.InsertLocation.after);
            (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.insertLineWithHeadingStyle)(p, "");
            include(p.getRange());
            cursor = p.getRange();
            continue;
          }
          const p = cursor.insertParagraph("", Word.InsertLocation.after);
          (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.insertLineWithHeadingStyle)(p, line);
          include(p.getRange());
          cursor = p.getRange();
        }
      }
      await context.sync();

      /* --------------------------------------------------
         3️⃣ Create ONE bookmark covering everything
      -------------------------------------------------- */
      if (bookmarkStart && bookmarkEnd) {
        const prefix = type === "Summary" ? "SM" : "ID";
        const bookmarkName = `${prefix}${tag.ID || tag.ReportHeadSummaryTagID}_Split_${getDateTimeStamp()}`;
        bookmarkStart.expandTo(bookmarkEnd).insertBookmark(bookmarkName);
      }

      /* --------------------------------------------------
         4️⃣ Remove invisible anchor
      -------------------------------------------------- */
      anchorChar.delete();
      await context.sync();
      (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.toaster)("Inserted successfully", "success");
    } catch (err) {
      console.error(err);
      (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.toaster)("Something went wrong", "error");
    }
  });
}
function getDateTimeStamp() {
  const d = new Date();
  const pad = n => n.toString().padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_` + `${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}
async function openPromptBuilderModal(tag, type) {
  const container = document.getElementById('confirmation-popup');
  if (!container) return;

  // Show the modal
  container.innerHTML = (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.PromptBuilderModalPopup)();
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
  const promptBuilderList = store.promptBuilderList || [];

  // References to modal elements
  const templateSelect = document.getElementById('promptBuilderTemplatePopup');
  const insertBtn = document.getElementById('prompt-builder-popup-insert');
  const cancelBtn = document.getElementById('prompt-builder-popup-cancel');
  const previewDiv = document.getElementById('previewPopup');
  const fieldsContainer = document.getElementById('fieldsContainerPopup');
  const previewContainer = document.getElementById('previewContainerPopup');
  const templateError = document.getElementById('templateErrorPopup');
  let fieldsList = [];
  let templateText = '';
  let currentPreview = '';

  // Populate template dropdown
  promptBuilderList.forEach(item => {
    const option = document.createElement('option');
    option.value = item.ID.toString();
    option.textContent = item.Name;
    templateSelect.appendChild(option);
  });

  // Close/Cancel functions
  const closeModal = () => {
    container.innerHTML = '';
  };
  cancelBtn.addEventListener('click', closeModal);

  // Template change logic
  templateSelect.addEventListener('change', async () => {
    const templateId = templateSelect.value;
    const jwt = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_9__.DocStorage.getItem('token') || '';
    try {
      const data = await (0,_draft_api__WEBPACK_IMPORTED_MODULE_0__.getPromptTemplateById)(templateId, jwt);
      if (data.Status && data.Data) {
        fieldsList = data.Data;
        const selectedTemplateObj = promptBuilderList.find(item => item.ID.toString() === templateId);
        templateText = selectedTemplateObj ? selectedTemplateObj.Template : '';
        currentPreview = templateText;
      }
      if (!templateId) {
        templateError.classList.remove('d-none');
        return;
      }
      templateError.classList.add('d-none');
      renderFields();
      updatePreview();
    } catch (err) {
      console.error("Failed to load prompt template fields:", err);
      (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.toaster)("Failed to load template fields", "error");
    }
  });
  function renderFields() {
    fieldsContainer.innerHTML = '';
    fieldsList.forEach(field => {
      const div = document.createElement('div');
      div.className = 'form-group mb-3';
      const label = document.createElement('label');
      label.className = 'form-label fw-semibold small';
      label.textContent = field.Label;
      div.appendChild(label);
      if (field.Type === 1) {
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'form-control';
        input.id = `modal-field-${field.Label}`;
        input.addEventListener('input', replaceKeywordsManually);
        div.appendChild(input);
      } else if (field.Type === 2) {
        const select = document.createElement('select');
        select.className = 'form-select';
        select.id = `modal-field-${field.Label}`;
        field.PromptTemplateOptionList.forEach(opt => {
          const option = document.createElement('option');
          option.value = opt.Text;
          option.textContent = opt.Option;
          select.appendChild(option);
        });
        select.addEventListener('change', replaceKeywordsManually);
        div.appendChild(select);
      }
      fieldsContainer.appendChild(div);
    });
  }
  function replaceKeywordsManually() {
    const keywordMap = {};
    fieldsList.forEach(field => {
      const id = `modal-field-${field.Label}`;
      const keyword = `#${field.Label}#`;
      let value = '';
      const element = document.getElementById(id);
      if (element) {
        value = element.value || '';
      }
      keywordMap[keyword] = value ? value : keyword;
    });
    let insertValue = templateText;
    for (const [keyword, value] of Object.entries(keywordMap)) {
      insertValue = insertValue.replace(new RegExp(keyword, 'g'), value);
    }
    currentPreview = insertValue;
    previewDiv.textContent = currentPreview;
    previewContainer.style.display = currentPreview ? 'block' : 'none';
    insertBtn.disabled = currentPreview === '';
  }
  function updatePreview() {
    replaceKeywordsManually();
  }

  // Insert logic: copy generated template text and insert into #chatInput
  insertBtn.addEventListener('click', () => {
    if (!currentPreview) return;
    const chatInput = document.getElementById('chatInput');
    if (chatInput) {
      chatInput.value = currentPreview;
      chatInput.dispatchEvent(new Event('input', {
        bubbles: true
      }));
    }
    closeModal();
  });
}
function initializeAIHistoryEvents(tag, jwt, availableKeys, type) {
  setTimeout(() => {
    tag.FilteredReportHeadAIHistoryList.forEach((chat, index) => {
      // Copy buttons
      if (tag.textareavalue) {
        document.getElementById(`chatInput`).value = tag.textareavalue;
        delete tag.textareavalue;
      }

      // After initializing buttons inside setTimeout
      const chatInput = document.getElementById("chatInput");
      const changeSourceButton = document.getElementById("changeSourceButton");
      if (chatInput && changeSourceButton) {
        // Enabled by default
        changeSourceButton.disabled = false;
      }
      ;
      document.getElementById(`copyPrompt-${index}`)?.addEventListener('click', () => (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.copyText)(chat.Prompt));
      const savePromptele = document.getElementById(`savePrompt-${index}`);
      if (savePromptele) {
        document.getElementById(`savePrompt-${index}`)?.addEventListener('click', () => {
          const container = document.getElementById('confirmation-popup');
          if (container) {
            container.innerHTML = (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.Confirmationpopup)('Do you want to save the current prompt as a global default?');

            // Wait for DOM to update and then attach cancel button listener
            setTimeout(() => {
              document.getElementById('confirmation-popup-cancel')?.addEventListener('click', () => {
                container.innerHTML = '';
              });
              document.getElementById('confirmation-popup-confirm')?.addEventListener('click', async () => {
                try {
                  document.getElementById('confirmation-popup-cancel')?.setAttribute('disabled', 'true');
                  document.getElementById('confirmation-popup-confirm')?.setAttribute('disabled', 'true');
                  let data;
                  if (type === 'Summary') {
                    const payload = {
                      Name: tag.Name,
                      Prompt: chat.Prompt
                    };
                    data = await (0,_summary_summary_api__WEBPACK_IMPORTED_MODULE_8__.updateSummaryTagPrompt)(payload, jwt);
                  } else {
                    let updatedTag = JSON.parse(JSON.stringify(tag));
                    updatedTag.Prompt = chat.Prompt;
                    data = await (0,_draft_api__WEBPACK_IMPORTED_MODULE_0__.updatePromptTemplate)(updatedTag, jwt);
                  }
                  if (data['Status']) {
                    (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.toaster)('Updated Succesfully', 'success');
                    container.innerHTML = '';
                  } else {
                    document.getElementById('confirmation-popup-cancel')?.setAttribute('disabled', 'false');
                    document.getElementById('confirmation-popup-confirm')?.setAttribute('disabled', 'false');
                    (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.toaster)('Something went wrong', 'error');
                  }
                } catch (error) {
                  document.getElementById('confirmation-popup-cancel')?.setAttribute('disabled', 'false');
                  document.getElementById('confirmation-popup-confirm')?.setAttribute('disabled', 'false');
                  (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.toaster)('Something went wrong', 'error');
                }
              });
            }, 0);
          }
        });
      }
      const openRefferance = document.getElementById(`openRefferance-${index}`);
      if (openRefferance) {
        document.getElementById(`openRefferance-${index}`)?.addEventListener('click', () => {
          const container = document.getElementById('confirmation-popup');
          if (container) {
            const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
            const sourceList = type === 'Summary' ? store.sourceSummaryList : store.sourceList;
            const rawSources = type === 'Summary' ? chat.SourceVector : chat.SourceValue;
            const sourceIds = Array.isArray(rawSources) ? rawSources : rawSources ? String(rawSources).split(',') : [];
            const chatSources = sourceIds.map(item => {
              if (type === 'Summary') {
                return sourceList.find(source => String(item) === String(source.VectorID));
              } else {
                return sourceList.find(source => Number(item) === source.VectorID);
              }
            });
            const sources = chatSources.filter(src => !!src);
            const popupData = {
              Data: chat.Evidences,
              Name: type === 'Summary' ? tag.Name : tag.DisplayName,
              UserValue: chat.Response,
              Sources: sources
            };
            container.innerHTML = (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.DataModalPopup)(popupData);

            // Wait for DOM to update and then attach cancel button listener
            setTimeout(() => {
              document.getElementById('confirmation-popup-cancel')?.addEventListener('click', () => {
                container.innerHTML = '';
              });
              document.getElementById('confirmation-popup-confirm')?.addEventListener('click', async () => {
                try {
                  document.getElementById('confirmation-popup-cancel')?.setAttribute('disabled', 'true');
                  document.getElementById('confirmation-popup-confirm')?.setAttribute('disabled', 'true');
                  let updatedTag = JSON.parse(JSON.stringify(tag));
                  updatedTag.Prompt = chat.Prompt;
                  const data = await (0,_draft_api__WEBPACK_IMPORTED_MODULE_0__.updatePromptTemplate)(updatedTag, jwt);
                  if (data['Status']) {
                    (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.toaster)('Updated Succesfully', 'success');
                    container.innerHTML = '';
                  } else {
                    document.getElementById('confirmation-popup-cancel')?.setAttribute('disabled', 'false');
                    document.getElementById('confirmation-popup-confirm')?.setAttribute('disabled', 'false');
                    (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.toaster)('Something went wrong', 'error');
                  }
                } catch (error) {
                  document.getElementById('confirmation-popup-cancel')?.setAttribute('disabled', 'false');
                  document.getElementById('confirmation-popup-confirm')?.setAttribute('disabled', 'false');
                  (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.toaster)('Something went wrong', 'error');
                }
              });
              document.getElementById('datamodel-popup-ok')?.addEventListener('click', async () => {
                container.innerHTML = '';
              });
            }, 0);
          }
        });
      }
      document.getElementById(`copyResponse-${index}`)?.addEventListener('click', () => (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.copyText)(chat.Response));

      // Checkbox logic
      const checkbox = document.getElementById(`checkbox-${index}`);
      if (checkbox) {
        checkbox.addEventListener('change', async event => {
          const isChecked = event.target.checked;

          // Reset all
          tag.FilteredReportHeadAIHistoryList.forEach((_, otherIndex) => {
            const otherCheckbox = document.getElementById(`checkbox-${otherIndex}`);
            const responseContainer = document.getElementById(`responseContainer-${otherIndex}`);
            if (otherCheckbox) otherCheckbox.checked = false;
            if (responseContainer) {
              responseContainer.classList.remove('ai-selected-response');
              responseContainer.classList.add('bg-light');
            }
            tag.FilteredReportHeadAIHistoryList[otherIndex].Selected = 0;
          });

          // Set selected
          if (isChecked) {
            checkbox.checked = true;
            const responseContainer = document.getElementById(`responseContainer-${index}`);
            if (responseContainer) {
              responseContainer.classList.add('ai-selected-response');
              responseContainer.classList.remove('bg-light');
            }
            chat.Selected = 1;
          } else {
            chat.Selected = 0;
          }
          try {
            const data = type === 'Summary' ? await (0,_summary_summary_api__WEBPACK_IMPORTED_MODULE_8__.updateSummaryHistory)(chat, jwt) : await (0,_draft_api__WEBPACK_IMPORTED_MODULE_0__.updateAiHistory)(chat, jwt);
            if (data['Data']) {
              tag.ReportHeadAIHistoryList = JSON.parse(JSON.stringify(data['Data']));
              tag.FilteredReportHeadAIHistoryList = [];
              tag.ReportHeadAIHistoryList.forEach(historyList => {
                historyList.Response = (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.removeQuotes)(historyList.Response);
                tag.FilteredReportHeadAIHistoryList.unshift(historyList);
              });
              const finalResponse = chat.FormattedResponse ? '\n' + (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.updateEditorFinalTable)(chat.FormattedResponse) : chat.Response;
              tag.ComponentKeyDataType = chat.FormattedResponse ? 'TABLE' : 'TEXT';
              tag.UserValue = finalResponse;
              tag.EditorValue = finalResponse;
              tag.text = finalResponse;
              const currentlySelected = tag.FilteredReportHeadAIHistoryList.some(item => item.Selected === 1);
              tag.IsApplied = !currentlySelected;
              if (type === 'Summary') {
                const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
                store.summaryTagList.forEach(currentTag => {
                  const currentId = currentTag.ID || currentTag.ReportHeadSummaryTagID;
                  const tagId = tag.ID || tag.ReportHeadSummaryTagID;
                  if (currentId === tagId) {
                    const isTable = chat.FormattedResponse !== '';
                    const finalResponse = chat.FormattedResponse ? '\n' + (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.updateEditorFinalTable)(chat.FormattedResponse) : chat.Response;
                    currentTag.ComponentKeyDataType = isTable ? 'TABLE' : 'TEXT';
                    currentTag.UserValue = finalResponse;
                    currentTag.EditorValue = finalResponse;
                    currentTag.text = finalResponse;
                    currentTag.IsApplied = tag.IsApplied;
                  }
                });
              } else {
                availableKeys.forEach(currentTag => {
                  if (currentTag.ID === tag.ID) {
                    const isTable = chat.FormattedResponse !== '';
                    const finalResponse = chat.FormattedResponse ? '\n' + (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.updateEditorFinalTable)(chat.FormattedResponse) : chat.Response;
                    currentTag.ComponentKeyDataType = isTable ? 'TABLE' : 'TEXT';
                    currentTag.UserValue = finalResponse;
                    currentTag.EditorValue = finalResponse;
                    currentTag.text = finalResponse;
                    currentTag.IsApplied = tag.IsApplied;
                  }
                });
                const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
                store.aiTagList.forEach(currentTag => {
                  if (currentTag.ID === tag.ID) {
                    const isTable = chat.FormattedResponse !== '';
                    const finalResponse = chat.FormattedResponse ? '\n' + (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.updateEditorFinalTable)(chat.FormattedResponse) : chat.Response;
                    currentTag.ComponentKeyDataType = isTable ? 'TABLE' : 'TEXT';
                    currentTag.UserValue = finalResponse;
                    currentTag.EditorValue = finalResponse;
                    currentTag.text = finalResponse;
                    currentTag.IsApplied = tag.IsApplied;
                  }
                });
              }
            }
          } catch (err) {
            console.error('Failed to update AI history:', err);
          }
        });
      }
    });

    // Jump to next bookmark button
    document.getElementById(`jump-to-next-tag`)?.addEventListener('click', async () => {
      await jumpToNextBookmarkOfTag(tag, type);
    });

    // Close button
    document.getElementById(`close-btn-tag`)?.addEventListener('click', () => {
      (0,_draft_functions__WEBPACK_IMPORTED_MODULE_1__.confirmSwitchChatHistory)(() => {
        const store = _services_store_service__WEBPACK_IMPORTED_MODULE_3__.StoreService.getInstance();
        store.currentChatTagId = -1;
        _utils_doc_storage__WEBPACK_IMPORTED_MODULE_9__.DocStorage.setItem("currentChatTagId", "-1");
        if (store.mode === "Home") {
          loadHomepage(availableKeys);
        } else if (store.mode === "Summary") {
          (0,_summary_summary__WEBPACK_IMPORTED_MODULE_6__.loadSummarypage)(availableKeys);
        }
      });
    });

    // Button: Prompt Builder
    document.getElementById(`promptBuilderButton`)?.addEventListener('click', () => {
      openPromptBuilderModal(tag, type);
    });

    // Button: Insert Tag
    document.getElementById(`insertTagButton`)?.addEventListener('click', () => {
      if (!tag.IsApplied) {
        insertTagPrompt(tag, type);
      }
    });

    // Button: Send Prompt
    document.getElementById(`sendPromptButton`)?.addEventListener('click', () => {
      const textareaValue = document.getElementById(`chatInput`).value;
      _services_ai_service__WEBPACK_IMPORTED_MODULE_4__.AIService.sendPrompt(tag, textareaValue, type);
    });

    // Button: Change Source
    document.getElementById(`changeSourceButton`)?.addEventListener('click', () => {
      const textareaValue = document.getElementById(`chatInput`).value;
      tag.textareavalue = textareaValue;
      (0,_taskpane__WEBPACK_IMPORTED_MODULE_2__.createMultiSelectDropdown)(tag, type);
    });

    // Mention dropdown
    (0,_taskpane__WEBPACK_IMPORTED_MODULE_2__.mentionDropdownFn)(`chatInput`, `mention-dropdown`, 'edit');
  }, 0);
}
async function jumpToNextBookmarkOfTag(tag, type) {
  return Word.run(async context => {
    const selection = context.document.getSelection();
    const bodyRange = context.document.body.getRange();
    const bookmarks = bodyRange.getBookmarks();
    await context.sync();
    const bookmarkNames = bookmarks.value || [];
    const prefix = type === "Summary" ? "SM" : "ID";
    const tagId = tag.ID || tag.ReportHeadSummaryTagID;
    const targetPrefix = `${prefix}${tagId}_Split_`;

    // Filter relevant bookmarks (case-insensitive start match)
    const relevantNames = bookmarkNames.filter(name => name.toUpperCase().startsWith(targetPrefix.toUpperCase()));
    if (relevantNames.length === 0) {
      (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.toaster)("No replaced instances found for this tag in the document.", "info");
      return;
    }

    // Get range and compare relation for each
    const items = relevantNames.map(name => {
      const r = context.document.getBookmarkRangeOrNullObject(name);
      r.load("isNullObject");
      const rel = r.compareLocationWith(selection);
      return {
        name,
        range: r,
        rel
      };
    });
    await context.sync();
    const validItems = items.filter(item => !item.range.isNullObject);
    if (validItems.length === 0) {
      (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.toaster)("No active instances found for this tag.", "info");
      return;
    }

    // Find the first bookmark that is positioned after the current selection
    let nextIndex = validItems.findIndex(item => {
      const val = String(item.rel.value).toLowerCase();
      return val === "after" || val === "adjacentafter";
    });

    // If none is after, wrap around to the first one
    if (nextIndex === -1) {
      nextIndex = 0;
    }
    const targetItem = validItems[nextIndex];
    targetItem.range.select();
    await context.sync();
    (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_5__.toaster)(`Jumped to instance ${nextIndex + 1} of ${validItems.length}`, "success");
  });
}

/***/ }),

/***/ "./src/taskpane/services/ai.service.ts":
/*!*********************************************!*\
  !*** ./src/taskpane/services/ai.service.ts ***!
  \*********************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   AIService: function() { return /* binding */ AIService; }
/* harmony export */ });
/* harmony import */ var _store_service__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./store.service */ "./src/taskpane/services/store.service.ts");
/* harmony import */ var _draft_draft_api__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../draft/draft.api */ "./src/taskpane/draft/draft.api.ts");
/* harmony import */ var _draft_draft_functions__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../draft/draft-functions */ "./src/taskpane/draft/draft-functions.ts");
/* harmony import */ var _draft_home__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../draft/home */ "./src/taskpane/draft/home.ts");
/* harmony import */ var _summary_summary_api__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../summary/summary.api */ "./src/taskpane/summary/summary.api.ts");





class AIService {
  static async fetchAIHistory(tag) {
    const store = _store_service__WEBPACK_IMPORTED_MODULE_0__.StoreService.getInstance();
    try {
      const data = await (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_1__.getAiHistory)(tag.ID, store.jwt);
      if (data.Status && data.Data) {
        tag.ReportHeadAIHistoryList = data['Data'] || [];
        tag.FilteredReportHeadAIHistoryList = [];
        tag.SourceValueID = tag.ReportHeadAIHistoryList[0].SourceValue;
        const selectedSources = store.sourceList.filter(list => tag.SourceValueID.includes(String(list.VectorID)));
        tag.SourceName = selectedSources.map(item => item.SourceName);
        tag.Sources = [...tag.SourceName];
        tag.TempSourceValue = selectedSources.map(item => item.VectorID ? String(item.VectorID) : item.SourceValue);
        tag.ReportHeadAIHistoryList.forEach(historyList => {
          historyList.Response = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_2__.removeQuotes)(historyList.Response);
          tag.FilteredReportHeadAIHistoryList.unshift(historyList);
        });
        return tag.FilteredReportHeadAIHistoryList;
      } else {
        console.warn("No AI history available.");
        return [];
      }
    } catch (error) {
      console.error('Error fetching AI history:', error);
      return [];
    }
  }
  static async sendPrompt(tag, prompt, type = "AITag") {
    const store = _store_service__WEBPACK_IMPORTED_MODULE_0__.StoreService.getInstance();
    if (prompt !== '' && !store.isTagUpdating) {
      store.isTagUpdating = true;

      // UI Updates via direct DOM or another UI service? 
      // For now, keeping direct DOM manipulation as in original code, but cleaner would be callbacks.
      const iconelement = document.getElementById(`sendPromptButton`);
      if (iconelement) iconelement.innerHTML = `<i class="fa fa-spinner fa-spin text-white"></i>`;
      let payload;
      const documentInstruction = store.documentInstruction || store.dataList?.DocumentInstruction || store.dataList?.DocumentInstructions || '';
      if (type === 'Summary') {
        payload = {
          ReportHeadID: store.dataList.ID,
          ReportHeadSummaryTagID: tag.ID,
          Prompt: prompt,
          Response: "",
          Selected: 1,
          SourceVector: tag.TempSourceValue ? tag.TempSourceValue.join(",") : "",
          Name: tag.Name,
          DocumentInstruction: documentInstruction
        };
      } else {
        payload = {
          ReportHeadID: tag.FilteredReportHeadAIHistoryList[0].ReportHeadID,
          DocumentID: store.dataList.NCTID,
          DocumentType: store.dataList.DocumentType,
          TextSetting: store.dataList.TextSetting,
          DocumentTemplate: store.dataList.ReportTemplate,
          ReportHeadGroupKeyID: tag.FilteredReportHeadAIHistoryList[0].ReportHeadGroupKeyID,
          ThreadID: tag.ThreadID,
          AssistantID: store.dataList.AssistantID,
          Container: store.dataList.Container,
          GroupName: store.GroupName,
          Prompt: prompt,
          PromptType: 1,
          Response: '',
          VectorID: store.dataList.VectorID,
          Selected: 0,
          ID: 0,
          SourceValue: tag.TempSourceValue ? tag.TempSourceValue : [],
          DocumentInstruction: documentInstruction
        };
      }
      try {
        store.isPendingResponse = true;
        const data = type === 'Summary' ? await (0,_summary_summary_api__WEBPACK_IMPORTED_MODULE_4__.addSummaryHistory)(payload, store.jwt) : await (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_1__.addAiHistory)(payload, store.jwt);
        if (data['Data'] && data['Data'] !== 'false') {
          tag.ReportHeadAIHistoryList = JSON.parse(JSON.stringify(data['Data']));
          tag.FilteredReportHeadAIHistoryList = [];
          tag.ReportHeadAIHistoryList.forEach(historyList => {
            historyList.Response = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_2__.removeQuotes)(historyList.Response);
            tag.FilteredReportHeadAIHistoryList.unshift(historyList);
          });
          const chat = tag.ReportHeadAIHistoryList[0];
          const finalResponse = chat.FormattedResponse ? '\n' + (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_2__.updateEditorFinalTable)(chat.FormattedResponse) : chat.Response;
          tag.ComponentKeyDataType = chat.FormattedResponse ? 'TABLE' : 'TEXT';
          tag.UserValue = finalResponse;
          tag.EditorValue = finalResponse;
          tag.text = finalResponse;

          // Update lists in Store
          if (type === 'Summary') {
            store.summaryTagList.forEach(currentTag => {
              const currentId = currentTag.ID || currentTag.ReportHeadSummaryTagID;
              const tagId = tag.ID || tag.ReportHeadSummaryTagID;
              if (currentId === tagId) {
                AIService.updateTagWithChat(currentTag, chat, tag.IsApplied);
              }
            });
          } else {
            store.aiTagList.forEach(currentTag => {
              if (currentTag.ID === tag.ID) {
                AIService.updateTagWithChat(currentTag, chat, tag.IsApplied);
              }
            });
            store.availableKeys.forEach(currentTag => {
              if (currentTag.ID === tag.ID) {
                AIService.updateTagWithChat(currentTag, chat, tag.IsApplied);
              }
            });
          }
          const appbody = document.getElementById('app-body');
          if (appbody) appbody.innerHTML = await (0,_draft_home__WEBPACK_IMPORTED_MODULE_3__.generateCheckboxHistory)(tag, type);
          store.isPendingResponse = false;
        }
        if (iconelement) iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`;
        const chatInput = document.getElementById(`chatInput`);
        if (chatInput) chatInput.value = '';
        store.isTagUpdating = false;
        store.isPendingResponse = false;
      } catch (error) {
        if (iconelement) iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`;
        store.isTagUpdating = false;
        store.isPendingResponse = false;
        console.error('Error sending AI prompt:', error);
      }
    } else {
      console.error('No empty prompt allowed or tag updating');
    }
  }
  static updateTagWithChat(currentTag, chat, isApplied) {
    const isTable = chat.FormattedResponse !== '';
    const finalResponse = chat.FormattedResponse ? '\n' + (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_2__.updateEditorFinalTable)(chat.FormattedResponse) : chat.Response;
    currentTag.ComponentKeyDataType = isTable ? 'TABLE' : 'TEXT';
    currentTag.UserValue = finalResponse;
    currentTag.EditorValue = finalResponse;
    currentTag.text = finalResponse;
    currentTag.IsApplied = isApplied;
  }
}

/***/ }),

/***/ "./src/taskpane/services/auth.service.ts":
/*!***********************************************!*\
  !*** ./src/taskpane/services/auth.service.ts ***!
  \***********************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   AuthService: function() { return /* binding */ AuthService; }
/* harmony export */ });
/* harmony import */ var _utils_config__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../utils/config */ "./src/taskpane/utils/config.ts");
/* harmony import */ var _draft_draft_api__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../draft/draft.api */ "./src/taskpane/draft/draft.api.ts");
/* harmony import */ var _store_service__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./store.service */ "./src/taskpane/services/store.service.ts");
/* harmony import */ var _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/doc-storage */ "./src/taskpane/utils/doc-storage.ts");




class AuthService {
  static TOKEN_KEY = 'user_token';
  static USER_ROLE_KEY = 'userRole';
  static STYLE_KEY = 'tableStyle';
  static PALETTE_KEY = 'colorPallete';
  static TEXT_STYLE_KEY = 'defaultTextStyle';
  static getStoredToken() {
    return _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.getItem('token');
  }
  static restoreSession() {
    const sessionToken = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.getItem('token');
    if (sessionToken) {
      // Check expiry (24 hours)
      const tokenLastUpdated = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.getItem('tokenLastUpdated');
      const now = new Date();
      if (tokenLastUpdated) {
        const lastUpdatedTime = new Date(tokenLastUpdated).getTime();
        const differenceInHours = (now.getTime() - lastUpdatedTime) / (1000 * 60 * 60);
        if (differenceInHours >= 24) {
          console.log("JWT token expired (passed 24 hours). Logging out.");
          this.logout();
          return null;
        } else {
          // Update timestamp to now, resetting the 24-hour expiration window
          _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.setItem('tokenLastUpdated', now.toISOString());
        }
      } else {
        // Initialize timestamp for legacy sessions
        _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.setItem('tokenLastUpdated', now.toISOString());
      }
      return {
        jwt: sessionToken,
        userRole: JSON.parse(_utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.getItem(this.USER_ROLE_KEY) || '{}'),
        tableStyle: _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.getItem(this.STYLE_KEY),
        colorPallete: JSON.parse(_utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.getItem(this.PALETTE_KEY) || 'null'),
        userId: _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.getItem('userId'),
        defaultTextStyle: _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.getItem(this.TEXT_STYLE_KEY)
      };
    }
    return null;
  }
  static async login(organization, username, password) {
    try {
      console.log(`Logging in to ${_utils_config__WEBPACK_IMPORTED_MODULE_0__.CONFIG.dataUrl}`);
      const data = await (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_1__.loginUser)(organization, username, password);
      if (data.Status === true && data['Data']) {
        if (data['Data'].ResponseStatus) {
          const jwt = data.Data.Token;
          const userRole = data.Data.UserRole;
          const userId = data.Data.ID;

          // Store interactions in Local Storage for persistence (scoped to document)
          _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.setItem('token', jwt);
          _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.setItem(this.USER_ROLE_KEY, JSON.stringify(userRole));
          _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.setItem('userId', userId);
          _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.setItem('tokenLastUpdated', new Date().toISOString());

          // Sync with StoreService and persist
          const store = _store_service__WEBPACK_IMPORTED_MODULE_2__.StoreService.getInstance();
          store.clearStorage();
          store.jwt = jwt;
          store.UserRole = userRole;
          store.userId = userId;
          store.saveToStorage();
          return {
            success: true,
            data: {
              token: jwt,
              userRole: userRole,
              userId: userId,
              raw: data.Data
            }
          };
        } else {
          return {
            success: false,
            message: "An error occurred during login. Please try again."
          };
        }
      } else {
        return {
          success: false,
          message: "An error occurred during login. Please try again."
        };
      }
    } catch (error) {
      console.error('Error during login:', error);
      return {
        success: false,
        message: "An error occurred during login. Please try again."
      };
    }
  }
  static async checkLoginType(organization, username) {
    try {
      const data = await (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_1__.checkLoginType)(organization, username);
      if (data && data.Status === true && data.Data) {
        return data.Data.AuthenticationType || 'TrialAssure';
      }
      return 'TrialAssure';
    } catch (error) {
      console.error('Error during checkLoginType:', error);
      return 'TrialAssure';
    }
  }
  static async ssoLogin(organization, username) {
    try {
      const data = await (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_1__.ssoLogin)(organization, username);
      if (data && data.Status === true && data.Data && data.Data.RedirectUrl) {
        return {
          success: true,
          redirectUrl: data.Data.RedirectUrl
        };
      } else {
        return {
          success: false,
          message: data && data.Message || 'Something went wrong during SSO login'
        };
      }
    } catch (error) {
      console.error('Error during ssoLogin:', error);
      return {
        success: false,
        message: 'Connection Lost'
      };
    }
  }
  static async ssoComplete(key) {
    try {
      const data = await (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_1__.ssoComplete)(key);
      if (data && data.Status === true && data.Data) {
        const jwt = data.Data.Token;
        const userRole = data.Data.UserRole;
        const userId = data.Data.ID;

        // Store interactions in Local Storage for persistence (scoped to document)
        _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.setItem('token', jwt);
        _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.setItem(this.USER_ROLE_KEY, JSON.stringify(userRole));
        _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.setItem('userId', userId);
        _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.setItem('tokenLastUpdated', new Date().toISOString());

        // Sync with StoreService and persist
        const store = _store_service__WEBPACK_IMPORTED_MODULE_2__.StoreService.getInstance();
        store.clearStorage();
        store.jwt = jwt;
        store.UserRole = userRole;
        store.userId = userId;
        store.saveToStorage();
        return {
          success: true,
          data: {
            token: jwt,
            userRole: userRole,
            userId: userId,
            raw: data.Data
          }
        };
      } else {
        return {
          success: false,
          message: data && data.Message || 'SSO complete failed'
        };
      }
    } catch (error) {
      console.error('Error during ssoComplete:', error);
      return {
        success: false,
        message: 'Connection Lost'
      };
    }
  }
  static logout() {
    _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.removeItem('token');
    _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.removeItem(this.USER_ROLE_KEY);
    _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.removeItem('userId');
    _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.removeItem('tokenLastUpdated');
    const store = _store_service__WEBPACK_IMPORTED_MODULE_2__.StoreService.getInstance();
    store.clearStorage();
    console.log("Logged out");
  }
}

/***/ }),

/***/ "./src/taskpane/services/document.service.ts":
/*!***************************************************!*\
  !*** ./src/taskpane/services/document.service.ts ***!
  \***************************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   DocumentService: function() { return /* binding */ DocumentService; }
/* harmony export */ });
/* harmony import */ var _store_service__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./store.service */ "./src/taskpane/services/store.service.ts");
/* harmony import */ var _draft_draft_api__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../draft/draft.api */ "./src/taskpane/draft/draft.api.ts");
/* harmony import */ var _draft_draft_functions__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../draft/draft-functions */ "./src/taskpane/draft/draft-functions.ts");
/* global Word */



class DocumentService {
  static transformDocumentName(value) {
    if (!value || value.trim() === '') return value;
    const parts = value.split('_');
    if (parts.length <= 1) return value;
    return parts.slice(1).join('_').replace(/%20/g, ' ').replace(/%25/g, '%');
  }
  static async loadReportData(documentId, jwt, userId) {
    try {
      console.log(`Fetching report data for ID: ${documentId}`);
      const data = await (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_1__.getReportById)(documentId, jwt);
      if (!data.Status || !data.Data) {
        throw new Error("Failed to fetch report data");
      }
      const dataList = data.Data;
      const documentInstruction = dataList.DocumentInstruction || dataList.DocumentInstructions || '';
      _store_service__WEBPACK_IMPORTED_MODULE_0__.StoreService.getInstance().documentInstruction = documentInstruction;

      // Basic processing
      if (!dataList.SourceTypeList) dataList.SourceTypeList = [];
      const sourceList = dataList.SourceTypeList.filter(item => item.SourceValue !== '' && item.AIFlag === 1).map(item => ({
        ...item,
        SourceName: decodeURIComponent(DocumentService.transformDocumentName(item.SourceValue))
      }));
      const clientId = dataList.ClientID;
      const aiGroup = dataList.Group.find(element => element.DisplayName === 'AIGroup');
      const groupName = aiGroup ? aiGroup.Name : '';
      const aiTagList = aiGroup ? aiGroup.GroupKey : [];

      // Image processing - Deferred to background (getImages)

      // Available Keys filtering
      let availableKeys = dataList.GroupKeyAll.filter(element => element.ComponentKeyDataType === 'TABLE' || element.ComponentKeyDataType === 'TEXT');
      const imageList = [];

      // Apply transformations to keys (updateEditorFinalTable)
      const processKey = key => {
        if (key.AIFlag === 1) {
          const regex = /<TableStart>([\s\S]*?)<TableEnd>/gi;
          if (regex.exec(key.EditorValue) !== null) {
            key.EditorValue = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_2__.updateEditorFinalTable)(key.EditorValue);
            key.UserValue = key.EditorValue;
            key.InitialTable = true;
            key.ComponentKeyDataType = 'TABLE';
          }
        }
      };
      availableKeys.forEach(processKey);
      aiTagList.forEach(processKey);
      return {
        dataList,
        documentInstruction,
        availableKeys,
        sourceList,
        clientId,
        groupName,
        aiTagList,
        imageList,
        clientList: [],
        // Will be filled if needed or fetched separately
        promptBuilderList: [] // Fetched via loadPromptTemplates
      };
    } catch (error) {
      console.error("Error in loadReportData:", error);
      throw error;
    }
  }
  static async retrieveDocumentProperties() {
    try {
      return await Word.run(async context => {
        const properties = context.document.properties.customProperties;
        properties.load("items");
        await context.sync();
        const property = properties.items.find(prop => prop.key === 'DocumentID');
        const orgName = properties.items.find(prop => prop.key === 'Organization');
        const environment = properties.items.find(prop => prop.key === 'Environment');
        const url = properties.items.find(prop => prop.key === 'URL');
        if (property && orgName) {
          return {
            documentID: property.value,
            organizationName: orgName.value,
            environment: environment ? environment.value : 'unknown',
            // Default to 'Production' if not set
            URL: url ? url.value : ''
          };
        } else {
          return null;
        }
      });
    } catch (error) {
      console.error("Error retrieving document properties", error);
      throw error;
    }
  }
  static async insertText(text) {
    await Word.run(async context => {
      const body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  }
}

/***/ }),

/***/ "./src/taskpane/services/store.service.ts":
/*!************************************************!*\
  !*** ./src/taskpane/services/store.service.ts ***!
  \************************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   StoreService: function() { return /* binding */ StoreService; }
/* harmony export */ });
/* harmony import */ var _utils_doc_storage__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../utils/doc-storage */ "./src/taskpane/utils/doc-storage.ts");

class StoreService {
  static STORAGE_KEY = 'link_ai_store_v1';

  // State Variables
  jwt = '';
  UserRole = {};
  documentID = '';
  organizationName = '';
  documentInstruction = '';
  aiTagList = [];
  summaryTagList = [];
  imageList = [];
  initialised = true;
  availableKeys = [];
  promptBuilderList = [];
  glossaryName = '';
  isGlossaryActive = false;
  GroupName = '';
  layTerms = [];
  dataList = [];
  isTagUpdating = false;
  capturedFormatting = {};
  emptyFormat = false;
  isNoFormatTextAvailable = false;
  clientId = '0';
  userId = 0;
  clientList = [];
  currentYear = new Date().getFullYear();
  selectedNames = [];
  isPendingResponse = false;
  theme = 'Light';
  mode = 'Home';
  tableStyle = 'Plain Table 5';
  defaultTextStyle = 'Normal';
  customizedTextStyle = null;
  colorPallete = {
    "Header": '#FFFFFF',
    "Primary": '#FFFFFF',
    "Secondary": '#FFFFFF',
    "Customize": true,
    "IsHeaderBold": true,
    "IsSideHeaderBold": false
  };
  environment = '';
  customTableStyle = [];
  customizedStyles = [];
  customTextStylesLoaded = false;
  currentChatTagId = -1;
  isReversed = false;
  reprocessingTagIds = {};
  constructor() {
    // Do NOT load from storage here — documentID is not yet known.
    // Call initForDocument(docId) after retrieving the document properties.
  }
  static getInstance() {
    if (!StoreService.instance) {
      StoreService.instance = new StoreService();
    }
    return StoreService.instance;
  }

  /**
   * Must be called once, after documentID is known (from document properties).
   * Scopes all localStorage I/O to this document and rehydrates saved state.
   */
  initForDocument(docId, environment) {
    this.documentID = docId;
    this.environment = environment || '';
    const storageId = environment ? `${docId}_${environment}` : docId;
    (0,_utils_doc_storage__WEBPACK_IMPORTED_MODULE_0__.setDocStorageId)(storageId);
    this.loadFromStorage();
  }

  /**
   * Persist current critical state to localStorage (scoped to this document).
   */
  saveToStorage() {
    const dataToSave = {
      jwt: this.jwt,
      UserRole: this.UserRole,
      documentID: this.documentID,
      organizationName: this.organizationName,
      documentInstruction: this.documentInstruction,
      theme: this.theme,
      mode: this.mode,
      userId: this.userId,
      tableStyle: this.tableStyle,
      colorPallete: this.colorPallete,
      clientId: this.clientId,
      defaultTextStyle: this.defaultTextStyle,
      customizedTextStyle: this.customizedTextStyle
    };
    _utils_doc_storage__WEBPACK_IMPORTED_MODULE_0__.DocStorage.setItem(StoreService.STORAGE_KEY, JSON.stringify(dataToSave));
  }

  /**
   * Rehydrate state from localStorage (scoped to this document).
   */
  loadFromStorage() {
    try {
      const stored = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_0__.DocStorage.getItem(StoreService.STORAGE_KEY);
      if (stored) {
        const data = JSON.parse(stored);
        if (data.jwt) this.jwt = data.jwt;
        if (data.UserRole) this.UserRole = data.UserRole;
        if (data.organizationName) this.organizationName = data.organizationName;
        if (data.documentInstruction) this.documentInstruction = data.documentInstruction;
        if (data.theme) this.theme = data.theme;
        if (data.mode) this.mode = data.mode;
        if (data.userId) this.userId = data.userId;
        if (data.tableStyle) this.tableStyle = data.tableStyle;
        if (data.colorPallete) this.colorPallete = data.colorPallete;
        if (data.clientId) this.clientId = data.clientId;
        if (data.defaultTextStyle) this.defaultTextStyle = data.defaultTextStyle;
        if (data.customizedTextStyle) this.customizedTextStyle = data.customizedTextStyle;
      }
    } catch (error) {
      console.error("Failed to load state from storage", error);
    }
  }

  /**
   * Clear stored state for this document.
   */
  clearStorage() {
    _utils_doc_storage__WEBPACK_IMPORTED_MODULE_0__.DocStorage.removeItem(StoreService.STORAGE_KEY);
    // Reset local variables
    this.jwt = '';
    this.UserRole = {};
    this.userId = 0;
    this.documentInstruction = '';
    this.isReversed = false;
    this.aiTagList = [];
    this.summaryTagList = [];
    this.imageList = [];
    this.dataList = [];
    this.selectedNames = [];
    this.currentChatTagId = -1;
    this.reprocessingTagIds = {};
    this.isPendingResponse = false;
    this.isTagUpdating = false;
    this.tableStyle = 'Plain Table 5';
    this.defaultTextStyle = 'Normal';
    this.customizedTextStyle = null;
    this.colorPallete = {
      "Header": '#FFFFFF',
      "Primary": '#FFFFFF',
      "Secondary": '#FFFFFF',
      "Customize": true,
      "IsHeaderBold": true,
      "IsSideHeaderBold": false
    };
    this.mode = 'Home';
    this.customizedStyles = [];
    this.customTextStylesLoaded = false;
  }
}

/***/ }),

/***/ "./src/taskpane/services/summary.service.ts":
/*!**************************************************!*\
  !*** ./src/taskpane/services/summary.service.ts ***!
  \**************************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   summaryService: function() { return /* binding */ summaryService; }
/* harmony export */ });
/* harmony import */ var _draft_draft_functions__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../draft/draft-functions */ "./src/taskpane/draft/draft-functions.ts");
/* harmony import */ var _summary_summary_api__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../summary/summary.api */ "./src/taskpane/summary/summary.api.ts");
/* harmony import */ var _store_service__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./store.service */ "./src/taskpane/services/store.service.ts");



class summaryService {
  static async fetchSummaryAIHistory(tag) {
    const store = _store_service__WEBPACK_IMPORTED_MODULE_2__.StoreService.getInstance();
    try {
      const data = await (0,_summary_summary_api__WEBPACK_IMPORTED_MODULE_1__.getSummaryTagHistory)(tag.ID || tag.ReportHeadSummaryTagID, store.jwt);
      if (data.Status && data.Data && data.Data.length > 0) {
        tag.ReportHeadAIHistoryList = data['Data'] || [];
        tag.FilteredReportHeadAIHistoryList = [];
        const latestHistory = tag.ReportHeadAIHistoryList[0];
        const rawSources = latestHistory.SourceVector || latestHistory.SourceValue || '';
        const sourceIds = Array.isArray(rawSources) ? rawSources.map(String) : String(rawSources).split(',').map(s => s.trim()).filter(Boolean);
        const selectedSources = store.sourceSummaryList.filter(list => sourceIds.includes(String(list.VectorID)));
        tag.SourceName = selectedSources.map(item => item.FileName || item.SourceName);
        tag.Sources = [...tag.SourceName];
        tag.TempSourceValue = selectedSources.map(item => item.VectorID ? String(item.VectorID) : item.SourceValue);
        tag.ReportHeadAIHistoryList.forEach(historyList => {
          historyList.Response = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_0__.removeQuotes)(historyList.Response);
          tag.FilteredReportHeadAIHistoryList.unshift(historyList);
        });
        return tag.FilteredReportHeadAIHistoryList;
      } else {
        console.warn("No Summary AI history available.");
        return [];
      }
    } catch (error) {
      console.error('Error fetching Summary AI history:', error);
      return [];
    }
  }
}

/***/ }),

/***/ "./src/taskpane/services/ui.service.ts":
/*!*********************************************!*\
  !*** ./src/taskpane/services/ui.service.ts ***!
  \*********************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   UIService: function() { return /* binding */ UIService; }
/* harmony export */ });
class UIService {
  static showNotification(message, type = 'info') {
    console.log(`[${type.toUpperCase()}] ${message}`);
    // Simple toastr wrapper, assuming toastr function is available globally or via import if we were stricter
    // For this refactor, we keep it compatible with existing 'toaster' global if possible, or implement simple DOM manipulation
    const toastContainer = document.getElementById('toastr'); // Hypothetical, in reality the old app used 'toaster()'
    if (typeof window.toaster === 'function') {
      window.toaster(message, type);
    } else {
      // Fallback
      alert(message);
    }
  }
  static toggleLoader(loading) {
    const loader = document.getElementById('page-loader');
    if (loader) loader.style.display = loading ? 'flex' : 'none';
  }
  static renderLoginPage(storedUrl, handleLoginCallback, themeToggleCallback, onUserBlurCallback) {
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
                <div class="mb-3" id="password-container">
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
    if (onUserBlurCallback) {
      document.getElementById('organization')?.addEventListener('blur', onUserBlurCallback);
      document.getElementById('username')?.addEventListener('blur', onUserBlurCallback);
    }
  }
  static applyTheme(theme) {
    const isDark = theme === 'Dark';
    const isLight = theme === 'Light';
    const safeApplyClass = (selector, darkClasses, lightClasses) => {
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
    safeApplyClass('.list-group-item-action', 'bg-dark text-light list-hover-dark', 'bg-light text-dark list-hover-light');
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
  static attachDashboardEvents(handlers) {
    document.getElementById('home')?.addEventListener('click', handlers.onHome);
    document.getElementById('summary-mode')?.addEventListener('click', handlers.onSummary);
    document.getElementById('glossary')?.addEventListener('click', handlers.onGlossary);
    document.getElementById('define-formatting')?.addEventListener('click', handlers.onFormat);
    document.getElementById('removeFormatting')?.addEventListener('click', handlers.onRemoveFormat);
    document.getElementById('theme-toggle')?.addEventListener('click', handlers.onThemeToggle);
    document.getElementById('logout')?.addEventListener('click', handlers.onLogout);
    document.getElementById('predefined-table')?.addEventListener('click', handlers.onPredefinedTable);
    document.getElementById('customized-table')?.addEventListener('click', handlers.onCustomizedTable);
    document.getElementById('default-text-style')?.addEventListener('click', handlers.onDefaultTextStyle);
    document.getElementById('customized-style')?.addEventListener('click', handlers.onCustomizedStyle);

    // Toggle sticky elements to static when mode dropdown is active
    const modeDropdownToggle = document.getElementById('modeDropdown');
    if (modeDropdownToggle) {
      const dropdownParent = modeDropdownToggle.parentElement;
      if (dropdownParent) {
        dropdownParent.addEventListener('show.bs.dropdown', () => {
          document.querySelectorAll('.sticky-top, .chat-header, .accordion-header').forEach(el => {
            el.classList.add('sticky-static');
          });
        });
        dropdownParent.addEventListener('hide.bs.dropdown', () => {
          document.querySelectorAll('.sticky-top, .chat-header, .accordion-header').forEach(el => {
            el.classList.remove('sticky-static');
          });
        });
      }
    }
  }
}

/***/ }),

/***/ "./src/taskpane/summary/summary.api.ts":
/*!*********************************************!*\
  !*** ./src/taskpane/summary/summary.api.ts ***!
  \*********************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   activateSummaryMode: function() { return /* binding */ activateSummaryMode; },
/* harmony export */   addSummaryHistory: function() { return /* binding */ addSummaryHistory; },
/* harmony export */   addSummaryTag: function() { return /* binding */ addSummaryTag; },
/* harmony export */   getSummaryTagHistory: function() { return /* binding */ getSummaryTagHistory; },
/* harmony export */   getSummaryTagStatus: function() { return /* binding */ getSummaryTagStatus; },
/* harmony export */   getSummaryTagsByReportHeadId: function() { return /* binding */ getSummaryTagsByReportHeadId; },
/* harmony export */   refreshSummaryMode: function() { return /* binding */ refreshSummaryMode; },
/* harmony export */   updateSummaryHistory: function() { return /* binding */ updateSummaryHistory; },
/* harmony export */   updateSummaryTagPrompt: function() { return /* binding */ updateSummaryTagPrompt; }
/* harmony export */ });
/* harmony import */ var _utils_config__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../utils/config */ "./src/taskpane/utils/config.ts");


// api.ts
const baseUrl = {
  toString: () => _utils_config__WEBPACK_IMPORTED_MODULE_0__.CONFIG.dataUrl
};
async function getSummaryTagsByReportHeadId(reportHeadId, jwt) {
  const response = await fetch(`${baseUrl}/api/summarytag/reportHead/${reportHeadId}`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  return await response.json();
}
async function activateSummaryMode(payload, jwt) {
  const response = await fetch(`${baseUrl}/api/report/activate-summarymode`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  return await response.json(); // if API returns JSON
}
async function refreshSummaryMode(payload, jwt) {
  const response = await fetch(`${baseUrl}/api/report/refresh-summarymode`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  return await response.json();
}
async function getSummaryTagHistory(reportHeadSummaryTagID, jwt) {
  const response = await fetch(`${baseUrl}/api/summarytag/history/${reportHeadSummaryTagID}`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  return await response.json();
}
async function getSummaryTagStatus(reportHeadSummaryTagID, jwt) {
  const response = await fetch(`${baseUrl}/api/summarytag/status/${reportHeadSummaryTagID}`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    }
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  return await response.json();
}
async function addSummaryHistory(payload, jwt) {
  const response = await fetch(`${baseUrl}/api/report/summary-history/add`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  return await response.json();
}
async function updateSummaryHistory(payload, jwt) {
  const response = await fetch(`${baseUrl}/api/report/summary-history/update`, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  return await response.json();
}
async function addSummaryTag(payload, jwt) {
  const response = await fetch(`${baseUrl}/api/report/summary-tag/add`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  return await response.json();
}
async function updateSummaryTagPrompt(payload, jwt) {
  const response = await fetch(`${baseUrl}/api/summarytag/update-prompt`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${jwt}`
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) {
    throw new Error('Network response was not ok.');
  }
  return await response.json();
}

/***/ }),

/***/ "./src/taskpane/summary/summary.ts":
/*!*****************************************!*\
  !*** ./src/taskpane/summary/summary.ts ***!
  \*****************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   getWordAsBase64: function() { return /* binding */ getWordAsBase64; },
/* harmony export */   loadSummarypage: function() { return /* binding */ loadSummarypage; },
/* harmony export */   summarySelectedNames: function() { return /* binding */ summarySelectedNames; }
/* harmony export */ });
/* harmony import */ var _taskpane__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../taskpane */ "./src/taskpane/taskpane.ts");
/* harmony import */ var _summary_api__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./summary.api */ "./src/taskpane/summary/summary.api.ts");
/* harmony import */ var _services_store_service__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../services/store.service */ "./src/taskpane/services/store.service.ts");
/* harmony import */ var _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/doc-storage */ "./src/taskpane/utils/doc-storage.ts");
/* harmony import */ var _components_bodyelements__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../components/bodyelements */ "./src/taskpane/components/bodyelements.ts");
/* harmony import */ var _draft_draft_functions__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../draft/draft-functions */ "./src/taskpane/draft/draft-functions.ts");






var summarySelectedNames = [];
let currentSummaryInstance = 0;
async function loadSummarypage(availableKeys) {
  const instanceId = ++currentSummaryInstance;
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_2__.StoreService.getInstance();
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
              <a class="dropdown-item" href="#" id="apply-btn-tag">
                <i class="fa-solid fa-circle-check me-2"></i> Apply
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
  const searchBox = document.getElementById('search-box');
  const list = document.getElementById('summary-tag-list');
  const pageButtons = document.getElementById('page-buttons');
  const pageCountLabel = document.getElementById('page-count-label');
  const btnFirst = document.getElementById('page-first');
  const btnPrev = document.getElementById('page-prev');
  const btnNext = document.getElementById('page-next');
  const btnLast = document.getElementById('page-last');
  const addBtn = document.getElementById('add-btn-tag');
  const applyBtn = document.getElementById('apply-btn-tag');
  const reanalyzeBtn = document.getElementById('reanalyze-draft');
  function disableActionButtons(disabled) {
    if (addBtn) addBtn.classList.toggle("disabled", disabled);
    if (applyBtn) applyBtn.classList.toggle("disabled", disabled);
  }
  function setReanalyzeButtonState(enabled) {
    if (reanalyzeBtn) reanalyzeBtn.classList.toggle("disabled", !enabled);
  }

  // ✅ Local state for this instance
  let isSummaryLoading = false;
  let allSummaryTags = [];
  let currentSummaryStatus = 0;

  // ✅ helper to normalize API response into string[]
  function normalizeNames(res) {
    let names = res?.Data ?? res?.SelectedNames ?? res?.selectedNames ?? res?.data ?? res ?? [];
    if (typeof names === "string") {
      return names.split(",").map(x => x.trim()).filter(Boolean);
    }
    if (Array.isArray(names) && names.length > 0 && typeof names[0] === "object") {
      return names.map(x => x.DisplayName || x.Name || x.TagName).filter(Boolean);
    }
    return Array.isArray(names) ? names : [];
  }
  let summaryEmptyMessage = "no available summary tags";

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
        list.innerHTML = `<div class="p-3 text-muted">${summaryEmptyMessage}</div>`;
      }
      return;
    }
    pageItems.forEach(tag => {
      const row = document.createElement('button');
      row.type = 'button';
      const themeClasses = store.theme === 'Dark' ? `bg-dark text-light list-hover-dark` : `bg-light text-dark list-hover-light`;
      row.className = `list-group-item list-group-item-action d-flex justify-content-between align-items-center ${themeClasses}`;
      let tagStatus = tag.Status === undefined || tag.Status === null ? "1" : String(tag.Status);
      if (store.reprocessingTagIds[String(tag.ID || tag.ReportHeadSummaryTagID)]) {
        tagStatus = "0";
      }
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
      row.onclick = async () => {
        const tagId = tag.ID || tag.ReportHeadSummaryTagID;
        if (store.currentChatTagId !== -1 && store.currentChatTagId !== undefined && store.currentChatTagId !== null && String(store.currentChatTagId) === String(tagId)) {
          return;
        }
        (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_5__.confirmSwitchChatHistory)(async () => {
          try {
            const appBody = document.getElementById('app-body');
            appBody.innerHTML = `
            <div id="button-container">
              <div class="loader" id="loader"></div>
            </div>
          `;
            store.currentChatTagId = tag.ID || tag.ReportHeadSummaryTagID;
            _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.setItem("currentChatTagId", String(store.currentChatTagId));
            const {
              generateCheckboxHistory
            } = await Promise.resolve(/*! import() */).then(__webpack_require__.bind(__webpack_require__, /*! ../draft/home */ "./src/taskpane/draft/home.ts"));
            const html = await generateCheckboxHistory(tag, "Summary");
            appBody.innerHTML = html;
          } catch {
            document.getElementById('app-body').innerHTML = '<div class="text-danger p-2">Error loading data</div>';
          }
        });
      };
      list.appendChild(row);
      if (tagStatus === "2") {
        const infoBtn = row.querySelector(`#reprocess-${tag.ID || tag.ReportHeadSummaryTagID}`);
        if (infoBtn) {
          infoBtn.addEventListener('click', e => {
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
    filtered = term ? allSummaryTags.filter(t => (t.Name || '').toLowerCase().includes(term)) : allSummaryTags;
    currentPage = 1;
    renderAll();
  }
  let debounceTimeout;
  searchBox.addEventListener('input', () => {
    clearTimeout(debounceTimeout);
    debounceTimeout = setTimeout(applySearch, 250);
  });
  btnFirst.addEventListener('click', () => {
    currentPage = 1;
    renderAll();
  });
  btnPrev.addEventListener('click', () => {
    currentPage = Math.max(1, currentPage - 1);
    renderAll();
  });
  btnNext.addEventListener('click', () => {
    currentPage = Math.min(getTotalPages(), currentPage + 1);
    renderAll();
  });
  btnLast.addEventListener('click', () => {
    currentPage = getTotalPages();
    renderAll();
  });

  // ✅ Action button wiring
  addBtn?.addEventListener('click', () => {
    if (!isSummaryLoading) (0,_taskpane__WEBPACK_IMPORTED_MODULE_0__.addGenAITags)();
  });
  applyBtn?.addEventListener('click', () => {
    if (!isSummaryLoading) (0,_taskpane__WEBPACK_IMPORTED_MODULE_0__.applyTagFn)();
  });

  // --------------------------
  // ✅ NEW FLOW: First call GET tags API and read SummaryTagGenerated from it
  // --------------------------
  let hasActivated = false;
  function deduplicateSummarySources(sources) {
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
      const getRes = await (0,_summary_api__WEBPACK_IMPORTED_MODULE_1__.getSummaryTagsByReportHeadId)(store.documentID, store.jwt);

      // ✅ 2) check Data.SummaryTagGenerated from GET API response
      currentSummaryStatus = getRes?.Data?.SummaryTagGenerated;

      // ✅ 3) Update status from store if manually reprocessing
      allSummaryTags = getRes?.Data?.SummaryTags || [];
      allSummaryTags.forEach(tag => {
        if (store.reprocessingTagIds[String(tag.ID || tag.ReportHeadSummaryTagID)]) {
          tag.Status = "0";
        }
      });
      store.summaryTagList = allSummaryTags;
      store.sourceSummaryList = deduplicateSummarySources(getRes?.Data?.SummarySources || []);
      filtered = allSummaryTags;
      renderAll();
      const summaryStatus = currentSummaryStatus;

      // Check if any tag is still processing (status 0) or manually reprocessing
      const hasProcessingTags = allSummaryTags.some(t => String(t.Status) === "0" || store.reprocessingTagIds[String(t.ID || t.ReportHeadSummaryTagID)]);
      if (summaryStatus === 2 && !hasProcessingTags) {
        setReanalyzeButtonState(true);
        if (allSummaryTags && allSummaryTags.length > 0) {
          summarySelectedNames = normalizeNames(allSummaryTags);
        }
        isSummaryLoading = false;
        renderAll();
        return;
      }

      // status 0 -> activate once, then poll until 2
      if (summaryStatus === 0) {
        if (!hasActivated) {
          try {
            hasActivated = true;
            const base64Data = await getWordAsBase64();
            const payload = {
              ReportHeadID: store.documentID,
              ActiveDocument: base64Data
            };
            await (0,_summary_api__WEBPACK_IMPORTED_MODULE_1__.activateSummaryMode)(payload, store.jwt);
          } catch (activationErr) {
            console.error("Activation failed:", activationErr);
            summaryEmptyMessage = "no summary tags available";
            isSummaryLoading = false;
            renderAll();
            return;
          }
        }
        await pollSummaryUntilDone();
      }

      // status 1 OR (status 2 but tags still processing) -> poll until 2
      if (summaryStatus === 1 || hasProcessingTags) {
        await pollSummaryUntilDone();
      }
    } catch (err) {
      console.error("Summary load failed:", err);
      if (list && instanceId === currentSummaryInstance) {
        list.innerHTML = `<div class="p-3 text-danger">Failed to load Summary mode</div>`;
      }
    } finally {
      if (instanceId === currentSummaryInstance) {
        isSummaryLoading = false;
        setReanalyzeButtonState(true);
        disableActionButtons(false);
        renderAll();
      }
    }
  }

  // ✅ Reprocess a single tag
  async function reprocessTag(tag) {
    try {
      tag.Status = "0"; // Show spinner immediately
      renderAll();
      const jwt = store.jwt;
      const historyRes = await (0,_summary_api__WEBPACK_IMPORTED_MODULE_1__.getSummaryTagHistory)(tag.ID || tag.ReportHeadSummaryTagID, jwt);
      const history = historyRes?.Data || [];
      if (history.length === 0) {
        (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_4__.toaster)("No history found to reprocess", "error");
        tag.Status = "2";
        renderAll();
        return;
      }

      // Latest history is usually index 0 after unshift, but let's check or just take first in raw list
      const lastHistory = history[0];
      const payload = {
        ReportHeadID: Number(store.documentID),
        ReportHeadSummaryTagID: tag.ID || tag.ReportHeadSummaryTagID,
        Prompt: lastHistory.Prompt,
        Response: "",
        Selected: 1,
        SourceVector: lastHistory.SourceVector ? lastHistory.SourceVector : '',
        Name: tag.Name
      };
      await (0,_summary_api__WEBPACK_IMPORTED_MODULE_1__.addSummaryHistory)(payload, jwt);
      store.reprocessingTagIds[String(tag.ID || tag.ReportHeadSummaryTagID)] = true;

      // Start polling if not already running (status 1)
      if (currentSummaryStatus !== 1) {
        currentSummaryStatus = 1;
        pollSummaryUntilDone();
      }
    } catch (err) {
      console.error("Reprocess failed:", err);
      (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_4__.toaster)("Reprocess failed", "error");
      tag.Status = "2";
      renderAll();
    }
  }

  // ✅ Ping API every 10 sec until status becomes 2 and all tags processed
  async function pollSummaryUntilDone() {
    try {
      while (instanceId === currentSummaryInstance) {
        const statusRes = await (0,_summary_api__WEBPACK_IMPORTED_MODULE_1__.getSummaryTagStatus)(store.documentID, store.jwt);
        if (instanceId !== currentSummaryInstance) break;
        const data = statusRes?.Data;
        const status = data?.SummaryTagGenerated;
        const tagStatuses = data?.SummaryTagStatus || [];
        currentSummaryStatus = status;

        // Update individual tag statuses in our local list if they exist in the response
        if (tagStatuses && tagStatuses.length > 0) {
          for (const ts of tagStatuses) {
            // Use for...of for await inside loop
            const matchingTag = allSummaryTags.find(t => (t.ReportHeadSummaryTagID || t.ID) === ts.ReportHeadSummaryTagID);
            if (matchingTag) {
              const oldStatus = String(matchingTag.Status);
              matchingTag.Status = String(ts.Status);
              if (String(ts.Status) !== "2") {
                delete store.reprocessingTagIds[String(ts.ReportHeadSummaryTagID)];
              }

              // Auto-refresh chat if currently viewing this tag and status changed from 0
              if (store.currentChatTagId === ts.ReportHeadSummaryTagID && oldStatus === "0" && String(ts.Status) !== "0") {
                const {
                  generateCheckboxHistory
                } = await Promise.resolve(/*! import() */).then(__webpack_require__.bind(__webpack_require__, /*! ../draft/home */ "./src/taskpane/draft/home.ts"));
                generateCheckboxHistory(matchingTag, "Summary").then(html => {
                  const appBody = document.getElementById('app-body');
                  if (appBody && store.currentChatTagId === ts.ReportHeadSummaryTagID) appBody.innerHTML = html;
                });
              }
            }
          }
          renderAll();
        }

        // Check if all tags are processed (status != "0")
        // We consider it done only if status is 2 and there are no tags with Status "0"
        const allTagsProcessed = tagStatuses.length > 0 ? tagStatuses.every(t => String(t.Status) !== "0") : true;
        if (status === 2 && allTagsProcessed) {
          setReanalyzeButtonState(true);
          disableActionButtons(false);
          isSummaryLoading = false;

          // after done -> fetch tags again and render
          const getRes2 = await (0,_summary_api__WEBPACK_IMPORTED_MODULE_1__.getSummaryTagsByReportHeadId)(store.documentID, store.jwt);
          if (instanceId !== currentSummaryInstance) break;
          allSummaryTags = getRes2?.Data?.SummaryTags || [];
          store.summaryTagList = allSummaryTags;
          filtered = allSummaryTags;
          if (allSummaryTags && allSummaryTags.length > 0) {
            summarySelectedNames = normalizeNames(allSummaryTags);
          }
          renderAll();
          break;
        }
        await new Promise(r => setTimeout(r, 10000));
      }
    } catch (err) {
      console.error("Polling failed:", err);
      // restart polling after a delay if still active instance
      if (instanceId === currentSummaryInstance) {
        setTimeout(pollSummaryUntilDone, 10000);
      }
    }
  }

  // ✅ Reanalyze wiring (same logic but triggers refresh + poll)
  reanalyzeBtn?.addEventListener('click', async () => {
    if (isSummaryLoading || reanalyzeBtn.classList.contains("disabled")) return;
    const popupContainer = document.getElementById("confirmation-popup");
    if (!popupContainer) return;
    popupContainer.innerHTML = (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_4__.Confirmationpopup)("Do you want to refresh all summary tags of the draft?");

    // Change button labels to No/Yes as requested
    const cancelBtn = document.getElementById("confirmation-popup-cancel");
    const confirmBtn = document.getElementById("confirmation-popup-confirm");
    if (cancelBtn) cancelBtn.innerText = "No";
    if (confirmBtn) confirmBtn.innerText = "Yes";
    const handleAction = async refresh => {
      popupContainer.innerHTML = ""; // Close popup
      if (refresh === null) {
        // Just cancel
        return;
      }
      try {
        setReanalyzeButtonState(false);
        disableActionButtons(true);
        isSummaryLoading = true;
        currentSummaryStatus = 1;

        // Start individual tag spinners immediately
        allSummaryTags.forEach(t => {
          t.Status = "0";
          store.reprocessingTagIds[String(t.ID || t.ReportHeadSummaryTagID)] = true;
        });
        renderRows();
        const base64Data = await getWordAsBase64();
        const payload = {
          ReportHeadID: Number(store.documentID),
          RefreshSummaryTag: refresh,
          ActiveDocument: base64Data
        };
        const res = await (0,_summary_api__WEBPACK_IMPORTED_MODULE_1__.refreshSummaryMode)(payload, store.jwt);
        if (instanceId !== currentSummaryInstance) return;

        // Immediately update state and UI from response
        if (res?.Data) {
          allSummaryTags = res.Data.SummaryTags || [];
          store.summaryTagList = allSummaryTags;
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
      } finally {
        if (instanceId === currentSummaryInstance) {
          setReanalyzeButtonState(true);
          disableActionButtons(false);
          isSummaryLoading = false;
        }
      }
    };
    document.getElementById("confirmation-popup-confirm")?.addEventListener("click", () => handleAction(true));
    document.getElementById("confirmation-popup-cancel")?.addEventListener("click", () => handleAction(false));
  });

  // ✅ Final: run new logic
  await firstLoadAndRender();

  // Reopen last active Summary Tag if applicable
  const savedTagId = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_3__.DocStorage.getItem("currentChatTagId");
  if (savedTagId && savedTagId !== "-1") {
    store.currentChatTagId = Number(savedTagId);
    const activeTag = allSummaryTags.find(t => (t.ID || t.ReportHeadSummaryTagID) === store.currentChatTagId);
    if (activeTag) {
      const appBody = document.getElementById('app-body');
      appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';
      const {
        generateCheckboxHistory
      } = await Promise.resolve(/*! import() */).then(__webpack_require__.bind(__webpack_require__, /*! ../draft/home */ "./src/taskpane/draft/home.ts"));
      generateCheckboxHistory(activeTag, "Summary").catch(() => appBody.innerHTML = '<div class="text-danger p-2">Error loading data</div>').then(html => {
        if (html) {
          appBody.innerHTML = html;
        }
      });
    }
  }
}

// ------------------ Base64 utils (unchanged) ------------------

function getWordAsBase64() {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, {
      sliceSize: 1024 * 1024
    }, result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(result.error);
        return;
      }
      const file = result.value;
      const slices = [];
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
    });
  });
}
function uint8ToBase64(bytes) {
  let binary = '';
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return window.btoa(binary);
}

/***/ }),

/***/ "./src/taskpane/taskpane.ts":
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.ts ***!
  \**********************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   addGenAITags: function() { return /* binding */ addGenAITags; },
/* harmony export */   applyAITagFn: function() { return /* binding */ applyAITagFn; },
/* harmony export */   applySummaryTagFn: function() { return /* binding */ applySummaryTagFn; },
/* harmony export */   applyTagFn: function() { return /* binding */ applyTagFn; },
/* harmony export */   applyglossary: function() { return /* binding */ applyglossary; },
/* harmony export */   checkGlossary: function() { return /* binding */ checkGlossary; },
/* harmony export */   createMultiSelectDropdown: function() { return /* binding */ createMultiSelectDropdown; },
/* harmony export */   customizeCustomStyle: function() { return /* binding */ customizeCustomStyle; },
/* harmony export */   customizeTable: function() { return /* binding */ customizeTable; },
/* harmony export */   customizeTextStyle: function() { return /* binding */ customizeTextStyle; },
/* harmony export */   formatOptionsDisplay: function() { return /* binding */ formatOptionsDisplay; },
/* harmony export */   getDocumentParagraphStyles: function() { return /* binding */ getDocumentParagraphStyles; },
/* harmony export */   getDocumentStyleDetails: function() { return /* binding */ getDocumentStyleDetails; },
/* harmony export */   mentionDropdownFn: function() { return /* binding */ mentionDropdownFn; },
/* harmony export */   normalizeBlankLines: function() { return /* binding */ normalizeBlankLines; },
/* harmony export */   removeMatchingContentControls: function() { return /* binding */ removeMatchingContentControls; },
/* harmony export */   removeTrailingEmptyParagraphs: function() { return /* binding */ removeTrailingEmptyParagraphs; }
/* harmony export */ });
/* harmony import */ var _utils_config__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./utils/config */ "./src/taskpane/utils/config.ts");
/* harmony import */ var _services_auth_service__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./services/auth.service */ "./src/taskpane/services/auth.service.ts");
/* harmony import */ var _services_document_service__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./services/document.service */ "./src/taskpane/services/document.service.ts");
/* harmony import */ var _services_ui_service__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./services/ui.service */ "./src/taskpane/services/ui.service.ts");
/* harmony import */ var _services_store_service__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./services/store.service */ "./src/taskpane/services/store.service.ts");
/* harmony import */ var _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./utils/doc-storage */ "./src/taskpane/utils/doc-storage.ts");
/* harmony import */ var _draft_home__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./draft/home */ "./src/taskpane/draft/home.ts");
/* harmony import */ var _draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./draft/draft-functions */ "./src/taskpane/draft/draft-functions.ts");
/* harmony import */ var _components_bodyelements__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./components/bodyelements */ "./src/taskpane/components/bodyelements.ts");
/* harmony import */ var _draft_draft_api__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ./draft/draft.api */ "./src/taskpane/draft/draft.api.ts");
/* harmony import */ var _components_tablestyles__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ./components/tablestyles */ "./src/taskpane/components/tablestyles.ts");
/* harmony import */ var _components_customstyles__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ./components/customstyles */ "./src/taskpane/components/customstyles.ts");
/* harmony import */ var _summary_summary__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ./summary/summary */ "./src/taskpane/summary/summary.ts");
/* harmony import */ var _summary_summary_api__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! ./summary/summary.api */ "./src/taskpane/summary/summary.api.ts");
// Imports






// Note: GlossaryService import removed if unused or moved

// Restoration of variables needed by the rest of the file (Legacy Support - check if needed)








Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("footer").innerText = `© ${new Date().getFullYear()} - TrialAssure LINK AI Assistant ${_utils_config__WEBPACK_IMPORTED_MODULE_0__.CONFIG.version}`;

    // 1. Check for SSO start parameter (inside dialog)
    const ssoStart = getQueryParam('ssoStart');
    if (ssoStart === 'true') {
      const targetUrl = getQueryParam('redirectUrl');
      if (targetUrl) {
        window.location.href = targetUrl;
        return;
      }
    }

    // 2. Check for SSO key inside the dialog
    const ssoKey = getQueryParam('key');
    if (ssoKey && typeof Office !== 'undefined' && Office.context && Office.context.ui && typeof Office.context.ui.messageParent === 'function') {
      Office.context.ui.messageParent(JSON.stringify({
        key: ssoKey
      }));
      return;
    }

    // Initialize Services
    // AuthService.init(); // if needed

    // Retrieve Properties via Service
    _services_document_service__WEBPACK_IMPORTED_MODULE_2__.DocumentService.retrieveDocumentProperties().then(async props => {
      if (props) {
        if (!_utils_config__WEBPACK_IMPORTED_MODULE_0__.CONFIG.environment.includes(props.environment) && props.environment !== 'unknown') {
          document.getElementById('app-body').innerHTML = `
        <p class="px-3 text-center">The document is not exported from this environment</p>`;
          console.log(`Custom property "documentID" not found.`);
        } else {
          // Update local state for legacy compatibility
          // documentID = props.documentID; // Moved to Store
          // organizationName = props.organizationName; // Moved to Store
          const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
          store.initForDocument(props.documentID, props.environment);
          store.organizationName = props.organizationName;
          store.environment = props.environment;
          if (props.URL) {
            _utils_config__WEBPACK_IMPORTED_MODULE_0__.CONFIG.dataUrl = props.URL;
          }

          // Check for SSO key first
          if (ssoKey) {
            _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(true);
            _services_auth_service__WEBPACK_IMPORTED_MODULE_1__.AuthService.ssoComplete(ssoKey).then(async ssoResult => {
              // Clear query params to clean up the URL
              const url = new URL(window.location.href);
              url.searchParams.delete('key');
              window.history.replaceState({}, document.title, url.toString());
              if (ssoResult.success) {
                const data = ssoResult.data;
                store.jwt = data.token;
                store.UserRole = data.userRole;
                store.userId = data.userId;
                store.saveToStorage();

                // Preserve legacy logic for style restoring
                const style = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('tableStyle');
                if (style) store.tableStyle = style;
                const defaultTextStyle = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('defaultTextStyle');
                if (defaultTextStyle) store.defaultTextStyle = defaultTextStyle;
                const localPallete = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('colorPallete');
                if (localPallete) store.colorPallete = JSON.parse(localPallete);
                _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
                (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)('You are successfully logged in', 'success');
                await displayMenu();
                window.location.hash = '#/dashboard';
              } else {
                _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
                (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)(ssoResult.message || 'SSO Login failed', 'error');
                loadLoginPage();
              }
            }).catch(err => {
              console.error("SSO completion error:", err);
              _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
              loadLoginPage();
            });
          } else {
            // Check Session
            const session = _services_auth_service__WEBPACK_IMPORTED_MODULE_1__.AuthService.restoreSession();
            if (session) {
              // Restore session state
              store.jwt = session.jwt;
              store.UserRole = session.userRole;
              if (session.tableStyle) store.tableStyle = session.tableStyle;
              if (session.colorPallete) store.colorPallete = session.colorPallete;
              if (session.defaultTextStyle) store.defaultTextStyle = session.defaultTextStyle;

              // Handle custom text style null properties fallback check
              if (store.customizedTextStyle && store.customizedTextStyle.properties === null) {
                getDocumentParagraphStyles().then(availableStyles => {
                  const styleNameInWord = availableStyles.find(s => s.toLowerCase() === store.customizedTextStyle.name.toLowerCase() || s.toLowerCase() === store.customizedTextStyle.id.toLowerCase());
                  if (styleNameInWord) {
                    store.defaultTextStyle = styleNameInWord;
                    _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("defaultTextStyle", styleNameInWord);
                    store.saveToStorage();
                  }
                }).catch(e => console.error("Error restoring custom style fallback on load:", e));
              }
              window.location.hash = '#/dashboard';
              (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)('You are successfully logged in', 'success');
              await displayMenu(); // Trigger legacy menu display
            } else {
              loadLoginPage();
            }
          }
        }
      } else {
        document.getElementById('app-body').innerHTML = `
        <p class="px-3 text-center">Export a document from the LINK AI application to use this functionality.</p>`;
        console.log(`Custom property "documentID" not found.`);
      }
    }).catch(err => {
      console.error("Failed to initialize", err);
    });

    // Theme is restored inside initForDocument → loadFromStorage, nothing extra needed here.

    // Setup UI
    setupEventHandlers();
  }
});
function setupEventHandlers() {
  // Porting event listeners
  document.getElementById("login-btn")?.addEventListener("click", handleLogin);
}
async function login() {
  const sessionToken = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('token');
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  if (sessionToken) {
    store.UserRole = JSON.parse(_utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('userRole')) || '';
    store.jwt = sessionToken;
    window.location.hash = '#/dashboard';
    const style = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('tableStyle');
    if (style) {
      store.tableStyle = style;
    }
    const defaultTextStyle = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('defaultTextStyle');
    if (defaultTextStyle) {
      store.defaultTextStyle = defaultTextStyle;
    }
    const localPallete = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('colorPallete');
    if (localPallete) {
      store.colorPallete = JSON.parse(localPallete);
    }

    // Handle custom text style null properties fallback check
    if (store.customizedTextStyle && store.customizedTextStyle.properties === null) {
      getDocumentParagraphStyles().then(availableStyles => {
        const styleNameInWord = availableStyles.find(s => s.toLowerCase() === store.customizedTextStyle.name.toLowerCase() || s.toLowerCase() === store.customizedTextStyle.id.toLowerCase());
        if (styleNameInWord) {
          store.defaultTextStyle = styleNameInWord;
          _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("defaultTextStyle", styleNameInWord);
          store.saveToStorage();
        }
      }).catch(e => console.error("Error restoring custom style fallback in login:", e));
    }
  } else {
    loadLoginPage();
  }
}
let currentAuthType = 'TrialAssure';
async function handleUserBlur() {
  const orgEl = document.getElementById('organization');
  const userEl = document.getElementById('username');
  if (!orgEl || !userEl) return;
  const org = orgEl.value;
  const username = userEl.value;
  if (org && username) {
    const detectedType = await _services_auth_service__WEBPACK_IMPORTED_MODULE_1__.AuthService.checkLoginType(org, username);
    currentAuthType = detectedType;
    const passwordContainer = document.getElementById('password-container');
    const passwordInput = document.getElementById('password');
    if (detectedType === 'AzureAD') {
      if (passwordContainer) {
        passwordContainer.style.display = 'none';
      }
      if (passwordInput) {
        passwordInput.removeAttribute('required');
        passwordInput.value = '';
      }
    } else {
      if (passwordContainer) {
        passwordContainer.style.display = 'block';
      }
      if (passwordInput) {
        passwordInput.setAttribute('required', 'true');
      }
    }
  }
}
function loadLoginPage() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.renderLoginPage(_utils_config__WEBPACK_IMPORTED_MODULE_0__.CONFIG.storeUrl, handleLogin, () => {
    store.theme = store.theme === 'Light' ? 'Dark' : 'Light';
    _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.applyTheme(store.theme);
    _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem('theme', store.theme);
  }, handleUserBlur);
}
async function handleLogin(event) {
  event.preventDefault();
  _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(true);
  try {
    const organizationInput = document.getElementById('organization').value;
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;
    const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
    const targetOrg = (store.organizationName || '').toLowerCase().trim();
    const enteredOrg = (organizationInput || '').toLowerCase().trim();
    if (enteredOrg === targetOrg && targetOrg !== '') {
      if (currentAuthType === 'AzureAD') {
        const result = await _services_auth_service__WEBPACK_IMPORTED_MODULE_1__.AuthService.ssoLogin(organizationInput, username);
        if (result.success && result.redirectUrl) {
          const localRedirectUrl = window.location.origin + window.location.pathname + "?ssoStart=true&redirectUrl=" + encodeURIComponent(result.redirectUrl);
          Office.context.ui.displayDialogAsync(localRedirectUrl, {
            height: 60,
            width: 35,
            displayInIframe: false
          }, asyncResult => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
              showLoginError("Could not open login window: " + asyncResult.error.message);
              return;
            }
            const dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, async args => {
              dialog.close();
              _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(true);
              if (args.message) {
                try {
                  const parsed = JSON.parse(args.message);
                  if (parsed.key) {
                    const ssoResult = await _services_auth_service__WEBPACK_IMPORTED_MODULE_1__.AuthService.ssoComplete(parsed.key);
                    if (ssoResult.success) {
                      const data = ssoResult.data;
                      store.jwt = data.token;
                      store.UserRole = data.userRole;
                      store.userId = data.userId;
                      store.saveToStorage();

                      // Preserve legacy logic for style restoring
                      const style = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('tableStyle');
                      if (style) store.tableStyle = style;
                      const defaultTextStyle = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('defaultTextStyle');
                      if (defaultTextStyle) store.defaultTextStyle = defaultTextStyle;
                      const localPallete = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('colorPallete');
                      if (localPallete) store.colorPallete = JSON.parse(localPallete);
                      _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
                      (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)('You are successfully logged in', 'success');
                      await displayMenu();
                      window.location.hash = '#/dashboard';
                    } else {
                      _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
                      showLoginError(ssoResult.message || "SSO complete failed");
                    }
                  } else {
                    _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
                    showLoginError("SSO login failed: no key received.");
                  }
                } catch (e) {
                  console.error("SSO message parsing error:", e);
                  _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
                  showLoginError("SSO login failed: invalid message data.");
                }
              } else {
                _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
                showLoginError("SSO login failed: empty response.");
              }
            });
            dialog.addEventHandler(Office.EventType.DialogEventReceived, args => {
              _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
              console.log("SSO dialog event: ", args);
            });
          });
        } else {
          showLoginError(result.message || "SSO Login failed");
        }
      } else {
        // Use AuthService
        const result = await _services_auth_service__WEBPACK_IMPORTED_MODULE_1__.AuthService.login(organizationInput, username, password);
        if (result.success) {
          const data = result.data;
          store.jwt = data.token;
          store.UserRole = data.userRole;
          store.userId = data.userId;
          store.saveToStorage();

          // Preserve legacy logic for style restoring
          const style = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('tableStyle');
          if (style) store.tableStyle = style;
          const defaultTextStyle = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('defaultTextStyle');
          if (defaultTextStyle) store.defaultTextStyle = defaultTextStyle;
          const localPallete = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('colorPallete');
          if (localPallete) store.colorPallete = JSON.parse(localPallete);
          (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)('You are successfully logged in', 'success');
          await displayMenu();
          window.location.hash = '#/dashboard';
        } else {
          showLoginError(result.message || "Login failed");
        }
      }
    } else {
      showLoginError("The organization specified is not associated with this document");
    }
  } catch (error) {
    console.error("Login process error:", error);
    showLoginError("An unexpected error occurred. Please try again.");
  } finally {
    _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
  }
}
function handleSsoMessage(event) {
  if (!event.data) {
    return;
  }
  let parsedData = event.data;

  // If the payload was sent as a string, parse it into an object
  if (typeof event.data === 'string') {
    try {
      parsedData = JSON.parse(event.data);
    } catch (e) {
      // Not a JSON string (could be internal webpack/chrome messages), ignore it
      return;
    }
  }

  // Verify if this message is our login payload (should contain a Token and a User ID/Username)
  const hasToken = parsedData.Token !== undefined || parsedData.token !== undefined;
  const hasUser = parsedData.UserID !== undefined || parsedData.userId !== undefined || parsedData.ID !== undefined || parsedData.Username !== undefined;
  if (hasToken && hasUser) {
    window.removeEventListener('message', handleSsoMessage);
    const token = parsedData.Token || parsedData.token;
    const userRole = parsedData.UserRole || parsedData.userRole;
    const userId = parsedData.ID || parsedData.UserID || parsedData.userId;

    // Store properties
    _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem('token', token);
    if (userRole) {
      _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem('userRole', typeof userRole === 'string' ? userRole : JSON.stringify(userRole));
    }
    _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem('userId', userId);
    _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem('tokenLastUpdated', new Date().toISOString());

    // Sync with StoreService and persist
    const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
    store.clearStorage();
    store.jwt = token;
    if (userRole) {
      store.UserRole = typeof userRole === 'string' ? JSON.parse(userRole) : userRole;
    }
    store.userId = userId;
    store.saveToStorage();

    // Preserve legacy logic for style restoring
    const style = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('tableStyle');
    if (style) store.tableStyle = style;
    const defaultTextStyle = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('defaultTextStyle');
    if (defaultTextStyle) store.defaultTextStyle = defaultTextStyle;
    const localPallete = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('colorPallete');
    if (localPallete) store.colorPallete = JSON.parse(localPallete);
    _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
    (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)('You are successfully logged in', 'success');
    displayMenu().then(() => {
      window.location.hash = '#/dashboard';
    });
  }
}
function getQueryParam(name) {
  const searchParams = new URLSearchParams(window.location.search);
  let value = searchParams.get(name);
  if (!value && window.location.hash.includes('?')) {
    const hashQuery = window.location.hash.split('?')[1];
    const hashParams = new URLSearchParams(hashQuery);
    value = hashParams.get(name);
  }
  return value;
}
function showLoginError(message) {
  loadLoginPage(); // Reload the form UI
  const errorDiv = document.getElementById('login-error');
  if (errorDiv) {
    errorDiv.style.display = 'block';
    errorDiv.textContent = message;
  }
}
async function displayMenu() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  store.userId = Number(_utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('userId'));
  // document.getElementById('aitag').addEventListener('click', redirectAI);
  await fetchDocument('Init');
}
async function getTableStyle() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  const tableStyleObj = await (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_9__.getAllCustomTables)(store.jwt);
  store.customTableStyle = tableStyleObj['Data'];
  const selectedTable = store.customTableStyle.find(style => style.ID === store.dataList.TableCustomizationID);
  if (selectedTable) {
    _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("CustomStyle", selectedTable ? selectedTable.Name : '');
    store.colorPallete = {
      "Header": selectedTable.Setting.HeaderColor,
      "Primary": selectedTable.Setting.PrimaryColor,
      "Secondary": selectedTable.Setting.SecondaryColor,
      "Customize": true,
      "IsHeaderBold": selectedTable.Setting.IsHeaderBold,
      "IsSideHeaderBold": selectedTable.Setting.IsSideHeaderBold
    };
    store.tableStyle = selectedTable.Setting.BaseStyle;
  }
}
async function getCustomTextStyles() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  try {
    const textStyleObj = await (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_9__.getAllCustomTexts)(store.jwt);
    if (textStyleObj && textStyleObj.Status && Array.isArray(textStyleObj.Data)) {
      store.customizedStyles = textStyleObj.Data.map(_components_customstyles__WEBPACK_IMPORTED_MODULE_11__.mapApiStyleToCustomStyle);
      const selectedTextStyle = store.customizedStyles.find(style => style.ID === store.dataList.TextCustomizationID || style.id === store.dataList.TextCustomizationID);
      if (selectedTextStyle) {
        if (selectedTextStyle.properties === null) {
          try {
            const availableStyles = await getDocumentParagraphStyles();
            const styleNameInWord = availableStyles.find(s => s.toLowerCase() === selectedTextStyle.name.toLowerCase() || s.toLowerCase() === selectedTextStyle.id.toLowerCase());
            if (styleNameInWord) {
              store.defaultTextStyle = styleNameInWord;
              _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("defaultTextStyle", styleNameInWord);
            }
          } catch (e) {
            console.error("Error checking matching Word style:", e);
          }
        }
        store.customizedTextStyle = selectedTextStyle;
        _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("customTextStyleId", selectedTextStyle.id);
        _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("customTextStyle", JSON.stringify(selectedTextStyle));
      }
      store.saveToStorage();
    } else {
      console.warn("No custom text styles found in API.");
      store.customizedStyles = [];
    }
  } catch (error) {
    console.error("Failed to load custom text styles:", error);
    store.customizedStyles = [];
  } finally {
    store.customTextStylesLoaded = true;
  }
}
async function fetchDocument(action) {
  _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(true);
  try {
    const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
    const userId = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('userId') || '0';
    const reportData = await _services_document_service__WEBPACK_IMPORTED_MODULE_2__.DocumentService.loadReportData(store.documentID, store.jwt, userId);

    // Assign to store
    store.dataList = reportData.dataList;
    store.documentInstruction = reportData.documentInstruction || store.dataList?.DocumentInstruction || store.dataList?.DocumentInstructions || '';
    await getTableStyle();
    await getCustomTextStyles();
    await loadPromptTemplates();
    store.availableKeys = reportData.availableKeys;
    store.sourceList = reportData.sourceList;
    store.clientId = reportData.clientId;
    // Global Assignment
    store.aiTagList = reportData.aiTagList;
    store.imageList = reportData.imageList;
    store.clientList = reportData.clientList;
    // promptBuilderList = reportData.promptBuilderList;

    // Handle Side Effects
    if (action === 'AIpanel' || action === 'Refresh' || action === 'Init') {
      if (store.mode === "Home") (0,_draft_home__WEBPACK_IMPORTED_MODULE_6__.loadHomepage)(store.availableKeys);
      if (store.mode === "Summary") (0,_summary_summary__WEBPACK_IMPORTED_MODULE_12__.loadSummarypage)(store.availableKeys);
    }

    // Render navigation header
    const logoHeaderEl = document.getElementById('logo-header');
    if (logoHeaderEl) {
      logoHeaderEl.innerHTML = (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.logoheader)(_utils_config__WEBPACK_IMPORTED_MODULE_0__.CONFIG.storeUrl);
    }
    (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.switchModeIcon)();
    _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);

    // Fetch images in background
    getImages();

    // Event Wiring
    _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.attachDashboardEvents({
      onHome: () => {
        (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.confirmSwitchChatHistory)(async () => {
          if (!store.isPendingResponse) {
            if (store.isGlossaryActive) await removeMatchingContentControls();
            store.currentChatTagId = -1;
            _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("currentChatTagId", "-1");
            (0,_draft_home__WEBPACK_IMPORTED_MODULE_6__.loadHomepage)(store.availableKeys);
          }
          store.mode = 'Home';
          (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.switchModeIcon)();
        });
      },
      onSummary: () => {
        (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.confirmSwitchChatHistory)(async () => {
          if (!store.isPendingResponse) {
            if (store.isGlossaryActive) await removeMatchingContentControls();
            store.currentChatTagId = -1;
            _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("currentChatTagId", "-1");
            (0,_summary_summary__WEBPACK_IMPORTED_MODULE_12__.loadSummarypage)(store.availableKeys);
          }
          store.mode = 'Summary';
          (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.switchModeIcon)();
        });
      },
      onGlossary: () => {
        (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.confirmSwitchChatHistory)(() => {
          if (store.emptyFormat) fetchGlossary();
        });
      },
      onFormat: () => {
        (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.confirmSwitchChatHistory)(() => {
          if (!store.isPendingResponse) formatOptionsDisplay();
        });
      },
      onRemoveFormat: () => {
        (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.confirmSwitchChatHistory)(() => {
          if (Object.keys(store.capturedFormatting).length > 0) removeOptionsConfirmation();
        });
      },
      onThemeToggle: () => {
        store.theme = store.theme === 'Light' ? 'Dark' : 'Light';
        _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.applyTheme(store.theme);
        _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem('theme', store.theme);
      },
      onLogout: () => {
        (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.confirmSwitchChatHistory)(async () => {
          if (!store.isPendingResponse) {
            if (store.isGlossaryActive) await removeMatchingContentControls();
            logout();
          }
        });
      },
      onPredefinedTable: () => {
        if (!store.isPendingResponse) {
          customizeTable('Pre');
        }
      },
      onCustomizedTable: () => {
        if (!store.isPendingResponse) {
          customizeTable('Custom');
        }
      },
      onDefaultTextStyle: () => {
        if (!store.isPendingResponse) {
          customizeTextStyle();
        }
      },
      onCustomizedStyle: () => {
        if (!store.isPendingResponse) {
          customizeCustomStyle();
        }
      }
    });

    // Register selection change handler for tag detection
    if (action === 'Init') {
      if (store.jwt) {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handleSelectionChange);
      }
    }
    _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
  } catch (error) {
    console.error("Error loading document data", error);
    _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.showNotification("Error loading data", "error");
    _services_ui_service__WEBPACK_IMPORTED_MODULE_3__.UIService.toggleLoader(false);
  }
}
async function formatOptionsDisplay() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  if (!store.isTagUpdating) {
    // Check if isTagUpdating is false
    if (store.isGlossaryActive) {
      await removeMatchingContentControls();
    }
    const htmlBody = `
      <div class="container pt-3">
        <div class="card">
          <div class="card-header">
               <!-- Buttons for Capture and Empty Format -->
            <div class="d-flex justify-content-end">
              <button id="capture-format-btn" class="btn btn-primary bg-primary-clr"><i class="fa fa-border-style me-1"></i>  Capture Format</button>
            </div>
            <!-- <h5 class="card-title">Formatting Options</h5> -->
          </div>
          <div class="card-body">
          <div class="formating-checkbox">
               <input type="checkbox" id="empty-format-checkbox" class="form-check-input">
              <label for="empty-format-checkbox" class="form-check-label empty-format-checkbox-label" style="flex: 1;">
                   Skip ignoring and removing format-based text
              </label>
            </div>

            <!-- Section to display captured formatting -->
            <div id="format-details">
              <h5 class="my-3">Selected Formatting:</h5>
              <ul id="format-list" class="list-unstyled"></ul>
            </div>
          </div>
        </div>
      </div>
    `;
    document.getElementById('app-body').innerHTML = htmlBody;
    if (Object.keys(store.capturedFormatting).length === 0) {
      const formatDetails = document.getElementById("format-details");
      formatDetails.style.display = 'none';
      // The object is not empty
    }
    const glossaryBtn = document.getElementById('glossary');
    if (!glossaryBtn.classList.contains('disabled-link')) {
      glossaryBtn.classList.add('disabled-link');
    }
    if (store.emptyFormat) {
      clearCapturedFormatting();
    } else {
      if (store.capturedFormatting.Bold === null || store.capturedFormatting.Bold === undefined || store.capturedFormatting.Underline === 'Mixed' || store.capturedFormatting.Underline === undefined || store.capturedFormatting.Size === null || store.capturedFormatting.Size === undefined || store.capturedFormatting["Font Name"] === null || store.capturedFormatting["Font Name"] === undefined || store.capturedFormatting["Background Color"] === '' || store.capturedFormatting["Background Color"] === undefined || store.capturedFormatting["Text Color"] === '' || store.capturedFormatting["Text Color"] === undefined) {
        const formatList = document.getElementById("format-list");
        formatList.innerHTML = "<p>Multiple style values found. Try again</p>";
        const removeFormatBtn = document.getElementById('removeFormatting');
        if (!removeFormatBtn.classList.contains('disabled-link')) {
          removeFormatBtn.classList.add('disabled-link');
        }
      } else {
        const removeFormatBtn = document.getElementById('removeFormatting');
        removeFormatBtn.classList.remove('disabled-link');
        displayCapturedFormatting();
      }
    }
    // Event listeners for the buttons

    document.getElementById("capture-format-btn").addEventListener("click", captureFormatting);
    const emptyFormatCheckbox = document.getElementById("empty-format-checkbox");
    if (store.isNoFormatTextAvailable) {
      emptyFormatCheckbox.checked = true;
      clearCapturedFormatting();
    }
    emptyFormatCheckbox.addEventListener("change", () => {
      if (emptyFormatCheckbox.checked) {
        store.isNoFormatTextAvailable = true;
        clearCapturedFormatting();
      } else {
        const CaptureBtn = document.getElementById('capture-format-btn');
        CaptureBtn.disabled = false;
        store.isNoFormatTextAvailable = false;
        store.emptyFormat = false;
        const glossaryBtn = document.getElementById('glossary');
        if (!glossaryBtn.classList.contains('disabled-link')) {
          glossaryBtn.classList.add('disabled-link');
        }
      }
    });
  }
}
function displayCapturedFormatting() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  store.emptyFormat = false;
  const formatList = document.getElementById("format-list");
  formatList.innerHTML = ""; // Clear the list before adding new items

  for (const [key, value] of Object.entries(store.capturedFormatting)) {
    if ((key === "Text Color" || key === "Background Color") && value) {
      formatList.innerHTML += `
        <li><strong>${key}:</strong>${value}
          <span style="display:inline-block;width:15px;height:15px;background-color:${value};border:1px solid black;"></span>
        </li>
      `;
    } else {
      formatList.innerHTML += `<li><strong>${key}:</strong> ${value}</li>`;
    }
  }
}
function clearCapturedFormatting() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  store.capturedFormatting = {}; // Clear the captured formatting object
  const formatDetails = document.getElementById("format-details");
  formatDetails.style.display = 'none';
  // formatList.innerHTML = `<li>No formatting selected.</li>`;
  store.emptyFormat = true;
  const glossaryBtn = document.getElementById('glossary');
  glossaryBtn.classList.remove('disabled-link');
  const CaptureBtn = document.getElementById('capture-format-btn');
  CaptureBtn.disabled = true;
  const removeFormatBtn = document.getElementById('removeFormatting');
  if (!removeFormatBtn.classList.contains('disabled-link')) {
    removeFormatBtn.classList.add('disabled-link');
  }
  console.log("Captured formatting cleared.");
}
async function captureFormatting() {
  try {
    await Word.run(async context => {
      const selection = context.document.getSelection();
      const font = selection.font;
      font.load(["bold", "italic", "underline", "size", "highlightColor", "name", 'color']);
      await context.sync();
      const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
      store.capturedFormatting = {
        Bold: font.bold,
        Italic: font.italic,
        Underline: font.underline,
        Size: font.size,
        "Background Color": font.highlightColor,
        "Font Name": font.name,
        'Text Color': font.color
      };
      const formatDetails = document.getElementById("format-details");
      formatDetails.style.display = 'block';
      if (store.capturedFormatting.Bold === null || store.capturedFormatting.Underline === 'Mixed' || store.capturedFormatting.Size === null || store.capturedFormatting["Font Name"] === null || store.capturedFormatting["Background Color"] === '' || store.capturedFormatting["Text Color"] === '') {
        const formatList = document.getElementById("format-list");
        formatList.innerHTML = "<p>Multiple style values found. Try again</p>";
        const removeFormatBtn = document.getElementById('removeFormatting');
        if (!removeFormatBtn.classList.contains('disabled-link')) {
          removeFormatBtn.classList.add('disabled-link');
        }
      } else {
        const removeFormatBtn = document.getElementById('removeFormatting');
        removeFormatBtn.classList.remove('disabled-link');
        displayCapturedFormatting();
      }
    });
  } catch (error) {
    console.error("Error capturing formatting:", error);
  }
}
async function removeOptionsConfirmation() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  if (!store.isTagUpdating) {
    if (store.isGlossaryActive) {
      await removeMatchingContentControls();
    } // Check if isTagUpdating is false
    const htmlBody = `
      <div class="container pt-3">
        <div class="card">
          <div class="card-header">
            <h5 class="card-title">Are you sure you want to remove formatted text ?</h5>
          </div>
          <div class="card-body">
          <div id="format-details">
              <h5>Selected Formatting:</h5>
              <ul id="format-list" class="list-unstyled mb-3"></ul>
              <small class="text-secondary font-italic" id="warning-rem-fmt"></small>
             
            </div>
               <!-- Buttons for Capture and Empty Format -->

            <div class="mt-3 d-flex justify-content-between">
              <span id="change-ft-btn" class="fw-bold text-primary my-auto c-pointer">Cancel</span>
              <button id="clear-ft-btn" class="btn btn-primary px-3"><i class="fa fa-check-circle me-2"></i>Yes</button>

            </div>

            
          </div>
        </div>
      </div>
    `;
    document.getElementById('app-body').innerHTML = htmlBody;
    displayCapturedFormatting();
    if (store.capturedFormatting['Background Color'] === null && store.capturedFormatting['Text Color'] === '#000000') {
      const warningEle = document.getElementById('warning-rem-fmt').innerHTML = 'Warning : The captured formatting is broad. This might result in unintended text removal throughout the document. Proceed?';
    }

    // Event listeners for the buttons
    document.getElementById("clear-ft-btn").addEventListener("click", removeFormattedText);
    document.getElementById("change-ft-btn").addEventListener("click", formatOptionsDisplay);
  }
}
async function removeFormattedText() {
  try {
    await Word.run(async context => {
      const iconelement = document.getElementById(`clear-ft-btn`);
      iconelement.innerHTML = `<i class="fa fa-spinner fa-spin text-white me-2"></i>Yes`;
      const clrBtn = document.getElementById('clear-ft-btn');
      clrBtn.disabled = true;
      const changeBtn = document.getElementById('change-ft-btn');
      changeBtn.disabled = true;
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items"); // Load paragraphs from the body

      await context.sync();
      const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();

      // Iterate through each paragraph in the document body
      for (const paragraph of paragraphs.items) {
        // Check if the paragraph contains text
        if (paragraph.text.trim() !== "") {
          const textRanges = paragraph.split([" "], true, true); // Split paragraph into individual words/segments
          textRanges.load("items, font");
          await context.sync();
          for (const range of textRanges.items) {
            const font = range.font;
            font.load(["bold", "italic", "underline", "size", "highlightColor", "name", "color"]);
            await context.sync();

            // Check if the text range matches the captured formatting
            if (font.highlightColor === store.capturedFormatting['Background Color'] && font.color === store.capturedFormatting['Text Color'] && font.bold === store.capturedFormatting['Bold'] && font.italic === store.capturedFormatting['Italic'] && font.size === store.capturedFormatting['Size'] && font.underline === store.capturedFormatting['Underline'] && font.name === store.capturedFormatting['Font Name']) {
              // Clear the range whether it's a full word or part of a word
              font.highlightColor = "#FFFFFF"; // Set new background color
              font.color = "#000000"; // Set new text color
              font.bold = false; // Reset bold if needed
              font.italic = false; // Reset italic if needed
              font.underline = "None";
              paragraph.insertText(" ", Word.InsertLocation.replace);
            }
          }
        }
      }
      await context.sync();
      store.capturedFormatting = {}; // Clear the captured formatting object
      const formatDetails = document.getElementById("format-details");
      formatDetails.style.display = 'none';
      // formatList.innerHTML = `<li>No formatting selected.</li>`;
      store.emptyFormat = true;
      store.isNoFormatTextAvailable = true;
      const glossaryBtn = document.getElementById('glossary');
      glossaryBtn.classList.remove('disabled-link');
      formatOptionsDisplay();
    });
  } catch (error) {
    console.error("Error removing formatted text:", error);
  }
}

// fetchAIHistory and sendPrompt moved to AIService
// Left empty or removed to prevent errors if these were exported.
// Since we removed exports at the top, we can now remove the functions.

// Your existing copyText function

async function logout() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  if (store.isGlossaryActive) {
    await removeMatchingContentControls();
  }
  _services_auth_service__WEBPACK_IMPORTED_MODULE_1__.AuthService.logout();
  _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.clearAll();
  sessionStorage.clear();
  window.location.hash = '#/new';
  store.initialised = true;
  document.getElementById('logo-header').innerHTML = ``;
  login();
}
async function applyTagFn() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  return Word.run(async context => {
    try {
      const body = context.document.body;
      context.load(body, 'text');
      await context.sync();
      if (store.mode === 'Summary') {
        await applySummaryTagFn(body, context);
      } else {
        await applyAITagFn(body, context);
        await applyImageTagFn(body, context);
      }
    } catch (err) {
      (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)("Something went wrong", "error");
      console.error("Error during tag application:", err);
      if (store.mode === 'Summary') {
        (0,_summary_summary__WEBPACK_IMPORTED_MODULE_12__.loadSummarypage)(store.availableKeys);
      } else {
        (0,_draft_home__WEBPACK_IMPORTED_MODULE_6__.loadHomepage)(store.availableKeys);
      }
    }
  });
}
async function applyImageTagFn(body, context) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  for (let i = 0; i < store.imageList.length; i++) {
    const tag = store.imageList[i];
    const searchResults = body.search(`$${tag.DisplayName}$`, {
      matchCase: false,
      matchWholeWord: false
    });
    context.load(searchResults, 'items');
    await context.sync();
    for (const item of searchResults.items) {
      if (tag.EditorValue !== "") {
        let base64Image = tag.EditorValue;

        // Clean base64
        if (!base64Image) continue;

        // Convert SVG → PNG
        if (base64Image.startsWith("data:image/svg+xml")) {
          base64Image = await (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.svgBase64ToPngBase64)(base64Image);
        }
        // Already PNG/JPEG → strip data prefix
        else if (base64Image.startsWith("data:image")) {
          base64Image = base64Image.split(",")[1];
        }
        const imageRange = item.getRange();
        imageRange.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.replace);
        await context.sync();
      }
    }
  }
  await context.sync();
  (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)("AI tag application completed!", "success");
  (0,_draft_home__WEBPACK_IMPORTED_MODULE_6__.loadHomepage)(store.availableKeys);
}
async function applySummaryTagFn(body, context) {
  document.getElementById('app-body').innerHTML = `
  <div id="button-container">
    <div class="loader" id="loader"></div>
    <div id="highlighted-text"></div>
  </div>`;
  (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)("Please wait... applying Summary tags", "info");
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  for (const tag of store.summaryTagList) {
    tag.EditorValue = removeQuotes(tag.Response);
    if (!tag.Response) continue;
    const results = body.search(`#${tag.Name}#`, {
      matchCase: false,
      matchWholeWord: false
    });
    context.load(results, "items");
    await context.sync();
    for (const item of results.items) {
      /* --------------------------------------------------
         1️⃣ Anchor correctly (NO invisible chars)
      -------------------------------------------------- */
      const anchor = item.getRange("Start");

      // Remove placeholder text completely
      item.delete();
      await context.sync();
      let cursor = anchor;
      let bookmarkStart = null;
      let bookmarkEnd = null;
      const include = r => {
        if (!bookmarkStart) {
          bookmarkStart = r.getRange("Start");
        }
        bookmarkEnd = r.getRange("End");
      };

      /* --------------------------------------------------
         2️⃣ Insert content forward from anchor
      -------------------------------------------------- */

      // TABLE CONTENT
      if (tag.ComponentKeyDataType === "TABLE") {
        const parser = new DOMParser();
        const doc = parser.parseFromString(tag.EditorValue, "text/html");
        const nodes = Array.from(doc.body.childNodes);
        for (const node of nodes) {
          // TEXT NODE
          if (node.nodeType === Node.TEXT_NODE) {
            let txt = node.textContent?.trim();
            if (!txt) continue;
            txt = txt.replace(/\n- /g, "\n• ");
            for (const line of txt.split(/\r?\n/)) {
              if (!line.trim()) {
                const p = cursor.insertParagraph("", Word.InsertLocation.after);
                (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.insertLineWithHeadingStyle)(p, "");
                include(p.getRange());
                cursor = p.getRange();
                continue;
              }
              const p = cursor.insertParagraph("", Word.InsertLocation.after);
              (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.insertLineWithHeadingStyle)(p, line);
              include(p.getRange());
              cursor = p.getRange();
            }
          }

          // ELEMENT NODE
          else if (node.nodeType === Node.ELEMENT_NODE) {
            const el = node;

            // TABLE
            if (el.tagName.toLowerCase() === "table") {
              const rows = Array.from(el.querySelectorAll("tr"));
              if (!rows.length) continue;
              let grid = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.parseHtmlTableToGrid)(rows);
              const tableCase = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.detectTableCase)(grid);
              const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
              const base = store.tableStyle.split(" - ")[0].trim();
              if (base === 'Table Grid 2') {
                store.isReversed = true;
              } else {
                store.isReversed = false;
              }
              if (store.isReversed && tableCase !== "CASE_1") {
                grid = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.transposeGrid)(grid);
              }
              const numRows = grid.length;
              const numCols = grid[0]?.length || 0;
              const p = cursor.insertParagraph("", Word.InsertLocation.after);
              const table = p.insertTable(numRows, numCols, Word.InsertLocation.after);
              const resolvedTableStyle = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.resolveWordTableStyle)(store.tableStyle);
              if (resolvedTableStyle !== 'none') {
                table.style = resolvedTableStyle;
              }

              // Population/Merging logic
              if (!store.isReversed) {
                if (!store.colorPallete.Customize) {
                  grid.forEach((rowGrid, rowIndex) => {
                    rowGrid.forEach((cellValue, cellIndex) => {
                      const tableCell = table.getCell(rowIndex, cellIndex);
                      tableCell.value = cellValue;
                      (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.applyCustomTextStyleToCell)(tableCell, store);
                    });
                  });

                  // Vertical merging logic for 1st column (standard view)
                  let lastParamRowIndex = -1;
                  grid.forEach((rowGrid, rowIndex) => {
                    if (rowIndex === 0) return; // skip header
                    const firstColText = rowGrid[0];
                    if (firstColText) {
                      lastParamRowIndex = rowIndex;
                    } else if (lastParamRowIndex !== -1) {
                      const topCell = table.getCell(lastParamRowIndex, 0);
                      const bottomCell = table.getCell(rowIndex, 0);
                      topCell.merge(bottomCell);
                      try {
                        topCell.verticalAlignment = Word.VerticalAlignment.center;
                        topCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
                      } catch (e) {}
                    }
                  });
                }
              } else {
                // Manual population for transposed table
                grid.forEach((rowGrid, rowIndex) => {
                  rowGrid.forEach((cellValue, cellIndex) => {
                    const tableCell = table.getCell(rowIndex, cellIndex);
                    tableCell.value = cellValue;
                    (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.applyCustomTextStyleToCell)(tableCell, store);
                  });
                });
              }

              // Styling logic
              if (store.colorPallete.Customize) {
                await (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.colorTable)(table, rows, context, store.isReversed);
              }
              include(table.getRange());
              cursor = table.getRange();
            }

            // OTHER ELEMENTS
            else {
              let txt = el.innerText?.trim();
              if (!txt) continue;
              txt = txt.replace(/\n- /g, "\n• ");
              for (const line of txt.split(/\r?\n/)) {
                if (!line.trim()) {
                  const p = cursor.insertParagraph("", Word.InsertLocation.after);
                  (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.insertLineWithHeadingStyle)(p, "");
                  include(p.getRange());
                  cursor = p.getRange();
                  continue;
                }
                const p = cursor.insertParagraph("", Word.InsertLocation.after);
                (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.insertLineWithHeadingStyle)(p, line);
                include(p.getRange());
                cursor = p.getRange();
              }
            }
          }
        }
      }

      // TEXT CONTENT
      else {
        const txt = tag.EditorValue.replace(/\n- /g, "\n• ").trim();
        for (const line of txt.split(/\r?\n/)) {
          if (!line.trim()) {
            const p = cursor.insertParagraph("", Word.InsertLocation.after);
            (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.insertLineWithHeadingStyle)(p, "");
            include(p.getRange());
            cursor = p.getRange();
            continue;
          }
          const p = cursor.insertParagraph("", Word.InsertLocation.after);
          (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.insertLineWithHeadingStyle)(p, line);
          include(p.getRange());
          cursor = p.getRange();
        }
      }
      await context.sync();

      /* --------------------------------------------------
         3️⃣ Create SINGLE bookmark
      -------------------------------------------------- */
      if (bookmarkStart && bookmarkEnd) {
        const bookmarkName = `SM${tag.ID || tag.ReportHeadSummaryTagID}_Split_${(0,_draft_home__WEBPACK_IMPORTED_MODULE_6__.getDateTimeStamp)()}`;
        bookmarkStart.expandTo(bookmarkEnd).insertBookmark(bookmarkName);
      }
    }
  }
  await context.sync();
  (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)("Summary tag application completed!", "success");
  (0,_summary_summary__WEBPACK_IMPORTED_MODULE_12__.loadSummarypage)(store.availableKeys);
}
async function applyAITagFn(body, context) {
  document.getElementById('app-body').innerHTML = `
  <div id="button-container">
    <div class="loader" id="loader"></div>
    <div id="highlighted-text"></div>
  </div>`;
  (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)("Please wait... applying AI tags", "info");
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  for (const tag of store.aiTagList) {
    tag.EditorValue = removeQuotes(tag.EditorValue);
    if (!tag.EditorValue || tag.IsApplied) continue;
    const results = body.search(`#${tag.DisplayName}#`, {
      matchCase: false,
      matchWholeWord: false
    });
    context.load(results, "items");
    await context.sync();
    for (const item of results.items) {
      /* --------------------------------------------------
         1️⃣ Anchor correctly (NO invisible chars)
      -------------------------------------------------- */
      const anchor = item.getRange("Start");

      // Remove placeholder text completely
      item.delete();
      await context.sync();
      let cursor = anchor;
      let bookmarkStart = null;
      let bookmarkEnd = null;
      const include = r => {
        if (!bookmarkStart) {
          bookmarkStart = r.getRange("Start");
        }
        bookmarkEnd = r.getRange("End");
      };

      /* --------------------------------------------------
         2️⃣ Insert content forward from anchor
      -------------------------------------------------- */

      // TABLE CONTENT
      if (tag.ComponentKeyDataType === "TABLE") {
        const parser = new DOMParser();
        const doc = parser.parseFromString(tag.EditorValue, "text/html");
        const nodes = Array.from(doc.body.childNodes);
        for (const node of nodes) {
          // TEXT NODE
          if (node.nodeType === Node.TEXT_NODE) {
            let txt = node.textContent?.trim();
            if (!txt) continue;
            txt = txt.replace(/\n- /g, "\n• ");
            for (const line of txt.split(/\r?\n/)) {
              if (!line.trim()) {
                const p = cursor.insertParagraph("", Word.InsertLocation.after);
                (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.insertLineWithHeadingStyle)(p, "");
                include(p.getRange());
                cursor = p.getRange();
                continue;
              }
              const p = cursor.insertParagraph("", Word.InsertLocation.after);
              (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.insertLineWithHeadingStyle)(p, line);
              include(p.getRange());
              cursor = p.getRange();
            }
          }

          // ELEMENT NODE
          else if (node.nodeType === Node.ELEMENT_NODE) {
            const el = node;

            // TABLE
            if (el.tagName.toLowerCase() === "table") {
              const rows = Array.from(el.querySelectorAll("tr"));
              if (!rows.length) continue;
              let grid = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.parseHtmlTableToGrid)(rows);
              const tableCase = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.detectTableCase)(grid);
              const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
              const base = store.tableStyle.split(" - ")[0].trim();
              if (base === 'Table Grid 2') {
                store.isReversed = true;
              } else {
                store.isReversed = false;
              }
              if (store.isReversed && tableCase !== "CASE_1") {
                grid = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.transposeGrid)(grid);
              }
              const numRows = grid.length;
              const numCols = grid[0]?.length || 0;
              const p = cursor.insertParagraph("", Word.InsertLocation.after);
              const table = p.insertTable(numRows, numCols, Word.InsertLocation.after);
              const resolvedTableStyle = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.resolveWordTableStyle)(store.tableStyle);
              if (resolvedTableStyle !== 'none') {
                table.style = resolvedTableStyle;
              }

              // Population/Merging logic
              if (!store.isReversed) {
                if (!store.colorPallete.Customize) {
                  grid.forEach((rowGrid, rowIndex) => {
                    rowGrid.forEach((cellValue, cellIndex) => {
                      const tableCell = table.getCell(rowIndex, cellIndex);
                      tableCell.value = cellValue;
                      (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.applyCustomTextStyleToCell)(tableCell, store);
                    });
                  });

                  // Vertical merging logic for 1st column (standard view)
                  let lastParamRowIndex = -1;
                  grid.forEach((rowGrid, rowIndex) => {
                    if (rowIndex === 0) return; // skip header
                    const firstColText = rowGrid[0];
                    if (firstColText) {
                      lastParamRowIndex = rowIndex;
                    } else if (lastParamRowIndex !== -1) {
                      const topCell = table.getCell(lastParamRowIndex, 0);
                      const bottomCell = table.getCell(rowIndex, 0);
                      topCell.merge(bottomCell);
                      try {
                        topCell.verticalAlignment = Word.VerticalAlignment.center;
                        topCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
                      } catch (e) {}
                    }
                  });
                }
              } else {
                // Manual population for transposed table
                grid.forEach((rowGrid, rowIndex) => {
                  rowGrid.forEach((cellValue, cellIndex) => {
                    const tableCell = table.getCell(rowIndex, cellIndex);
                    tableCell.value = cellValue;
                    (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.applyCustomTextStyleToCell)(tableCell, store);
                  });
                });
              }

              // Styling logic
              if (store.colorPallete.Customize) {
                await (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.colorTable)(table, rows, context, store.isReversed);
              }
              include(table.getRange());
              cursor = table.getRange();
            }

            // OTHER ELEMENTS
            else {
              let txt = el.innerText?.trim();
              if (!txt) continue;
              txt = txt.replace(/\n- /g, "\n• ");
              for (const line of txt.split(/\r?\n/)) {
                if (!line.trim()) {
                  const p = cursor.insertParagraph("", Word.InsertLocation.after);
                  (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.insertLineWithHeadingStyle)(p, "");
                  include(p.getRange());
                  cursor = p.getRange();
                  continue;
                }
                const p = cursor.insertParagraph("", Word.InsertLocation.after);
                (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.insertLineWithHeadingStyle)(p, line);
                include(p.getRange());
                cursor = p.getRange();
              }
            }
          }
        }
      }

      // IMAGE CONTENT
      else if (tag.ComponentKeyDataType === "IMAGE") {
        let base64 = tag.EditorValue;
        if (base64.startsWith("data:image/svg+xml")) {
          base64 = await (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.svgBase64ToPngBase64)(base64);
        } else if (base64.startsWith("data:image")) {
          base64 = base64.split(",")[1];
        }
        const pic = cursor.insertInlinePictureFromBase64(base64, Word.InsertLocation.after);
        include(pic.getRange());
        cursor = pic.getRange();
      }

      // TEXT CONTENT
      else {
        const txt = tag.EditorValue.replace(/\n- /g, "\n• ").trim();
        for (const line of txt.split(/\r?\n/)) {
          if (!line.trim()) {
            const p = cursor.insertParagraph("", Word.InsertLocation.after);
            (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.insertLineWithHeadingStyle)(p, "");
            include(p.getRange());
            cursor = p.getRange();
            continue;
          }
          const p = cursor.insertParagraph("", Word.InsertLocation.after);
          (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.insertLineWithHeadingStyle)(p, line);
          include(p.getRange());
          cursor = p.getRange();
        }
      }
      await context.sync();

      /* --------------------------------------------------
         3️⃣ Create SINGLE bookmark
      -------------------------------------------------- */
      if (bookmarkStart && bookmarkEnd) {
        const bookmarkName = `ID${tag.ID}_Split_${(0,_draft_home__WEBPACK_IMPORTED_MODULE_6__.getDateTimeStamp)()}`;
        bookmarkStart.expandTo(bookmarkEnd).insertBookmark(bookmarkName);
      }
    }
  }
}
async function normalizeBlankLines(context) {
  const paragraphs = context.document.body.paragraphs;
  context.load(paragraphs, "items, text");
  await context.sync();
  let previousWasEmpty = false;
  for (const p of paragraphs.items) {
    const isEmpty = !p.text || p.text.trim() === "";
    if (isEmpty && previousWasEmpty) {
      p.delete();
    }
    previousWasEmpty = isEmpty;
  }
  await context.sync();
}
async function removeTrailingEmptyParagraphs(context) {
  const paragraphs = context.document.body.paragraphs;
  context.load(paragraphs, "items, text");
  await context.sync();
  for (let i = paragraphs.items.length - 1; i >= 0; i--) {
    const p = paragraphs.items[i];
    if (!p.text || p.text.trim() === "") {
      p.delete();
    } else {
      break; // stop once real content is found
    }
  }
  await context.sync();
}
async function fetchGlossary() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  if (!store.isTagUpdating) {
    document.getElementById('app-body').innerHTML = `
  <div id="button-container">

          <div class="loader" id="loader"></div>

        <div id="highlighted-text"></div>`;
    loadGlossary();
  }
}
function loadGlossary() {
  document.getElementById('app-body').innerHTML = `
        <div id="button-container">
          <button class="btn btn-secondary me-2 mark-glossary btn-sm" id="applyglossary">Apply Glossary</button>
        </div>
  `;
  document.getElementById('applyglossary').addEventListener('click', applyglossary);
}
async function applyglossary() {
  document.getElementById('app-body').innerHTML = `
  <div id="button-container">

          <div class="loader" id="loader"></div>

        <div id="highlighted-text"></div>`;
  try {
    await Word.run(async context => {
      const body = context.document.body;
      body.load("text");
      await context.sync(); // Sync to get the text content

      const bodyText = {
        "Content": body.text.replace(/[\n\r]/g, ' ')
      };
      try {
        const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
        const data = await (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_9__.fetchGlossaryTemplate)(store.dataList?.ClientID, bodyText, store.jwt);
        store.layTerms = data.Data;
        if (data.Data.length > 0) {
          store.glossaryName = data.Data[0].GlossaryTemplate;
          loadGlossary();
        } else {
          document.getElementById('app-body').innerHTML = `
            <p class="text-center">Data not available</p>
          `;
        }
      } catch (error) {
        console.error('Error fetching glossary data:', error);
      }
      // Sort terms by length (longest first)
      const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
      store.layTerms.sort((a, b) => b.ClinicalTerm.length - a.ClinicalTerm.length);
      const processedTerms = new Set(); // Track added larger terms

      // Filter out smaller terms if they are included in a larger term
      const filteredTerms = store.layTerms.filter(term => {
        for (const biggerTerm of Array.from(processedTerms)) {
          if (typeof biggerTerm === 'string' && biggerTerm.includes(term.ClinicalTerm.toLowerCase())) {
            console.log(`Skipping "${term.ClinicalTerm}" because it's part of "${biggerTerm}"`);
            return false; // Exclude this smaller term
          }
        }
        processedTerms.add(term.ClinicalTerm.toLowerCase());
        return true;
      });
      store.filteredGlossaryTerm = filteredTerms;
      await removeMatchingContentControls();
      const foundRanges = new Map(); // Track words already processed

      const searchPromises = Array.from(store.filteredGlossaryTerm).map(term => {
        const searchResults = body.search(term.ClinicalTerm, {
          matchCase: false,
          matchWholeWord: false
        });
        searchResults.load("items");
        return searchResults;
      });
      await context.sync();
      for (const searchResults of searchPromises) {
        for (const range of searchResults.items) {
          if (!range || !range.text) {
            console.log("Invalid range. Skipping...");
            continue;
          }

          // Load existing content controls inside this range
          const font = range.font;
          font.load(["bold", "italic", "underline", "size", "highlightColor", "name", 'color']);
          range.load("contentControls");
          await context.sync();
          const existingControl = range.contentControls.items.length > 0;
          if (existingControl) {
            console.log(`Skipping "${range.text}" because it already has a content control.`);
            continue; // Skip if content control is already present
          }
          // Check if we've already processed this term at this range
          if (foundRanges.has(range.text)) {
            console.log(`Skipping duplicate occurrence of "${range.text}"`);
            continue;
          }
          // Mark this word as processed
          foundRanges.set(range.text, true);
          // Remove existing content controls if any
          if (range.contentControls && range.contentControls.items.length > 0) {
            console.log(`Removing existing content control from: "${range.text}"`);
            for (const control of range.contentControls.items) {
              control.delete(false); // 'false' keeps the text, only removes the control
            }
            await context.sync(); // Ensure deletion is applied before adding a new one
          }
          try {
            // Insert a new content control
            const contentControl = range.insertContentControl();
            contentControl.title = `${range.text}`;
            if (font.highlightColor !== null) {
              contentControl.tag = `${font.highlightColor}`;
            }
            contentControl.font.highlightColor = "yellow"; // Highlight the control
            contentControl.appearance = Word.ContentControlAppearance.boundingBox;
            await context.sync();
          } catch (error) {
            console.error(`Error inserting content control for term "${range.text}":`, error);
          }
        }
      }
      // document.getElementById('glossarycheck').style.display='block';
      store.isGlossaryActive = true;
      document.getElementById('app-body').innerHTML = `
      <div id="button-container">
        <button class="btn btn-secondary me-2 clear-glossary btn-sm" id="clearGlossary">Clear Glossary</button>
      </div>

      <div id="highlighted-text"></div>
      <div class="d-flex justify-content-center box-loader">
       <div class="loader" id="loader"></div></div>
      
`;
      const displayElement = document.getElementById('loader');
      displayElement.style.display = 'none';
      await context.sync();
      document.getElementById('clearGlossary').addEventListener('click', removeMatchingContentControls);
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handleSelectionChange);
    });

    // Optional: Notify user of completion
    console.log('Glossary applied successfully');
  } catch (error) {
    console.error('Error applying glossary:', error);
    // Optional: Notify user of error
    console.log('Error applying glossary. Please try again.');
  }
}
async function handleSelectionChange() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  if (!store.jwt) {
    return;
  }

  // Handle glossary mode
  if (store.isGlossaryActive) {
    await checkGlossary();
  }

  // Handle Home or Summary mode - detect bookmarks/tags in selection
  if (store.mode === 'Home' || store.mode === 'Summary') {
    await logBookmarksInSelection();
  }
}
async function checkGlossary() {
  try {
    await Word.run(async context => {
      const selection = context.document.getSelection();
      selection.load("text, font.highlightColor");
      await context.sync();
      if (selection.text) {
        const loader = document.getElementById('loader');
        if (loader) {
          loader.style.display = 'block';
        }
        const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
        const searchPromises = store.layTerms.map(term => {
          const searchResults = selection.search(term.ClinicalTerm, {
            matchCase: false,
            matchWholeWord: false
          });
          searchResults.load("items");
          return searchResults;
        });
        await context.sync();
        const selectedWords = [];
        for (const searchResults of searchPromises) {
          for (const range of searchResults.items) {
            const font = range.font;
            font.load(["bold", "italic", "underline", "size", "highlightColor", "name", "color"]);
            await context.sync();
            if (font.highlightColor !== store.capturedFormatting['Background Color'] || font.color !== store.capturedFormatting['Text Color'] || font.bold !== store.capturedFormatting['Bold'] || font.italic !== store.capturedFormatting['Italic'] || font.size !== store.capturedFormatting['Size'] || font.underline !== store.capturedFormatting['Underline'] || font.name !== store.capturedFormatting['Font Name']) {
              selectedWords.push(range.text);
            }
          }
        }
        // searchPromises.forEach(searchResults => {
        //   searchResults.items.forEach(item => {
        //   });
        // });
        displayHighlightedText(selectedWords);
        await context.sync();

        // const highlightColor = selection.font.highlightColor;

        // if (highlightColor === "red") {
        //   displayHighlightedText(selection.text);
        // } else {
        //   console.log('Selected text is not highlighted.');
        // }
      } else {
        console.log('No text is selected.');
      }
    });
  } catch (error) {
    console.error('Error displaying glossary:', error);
  }
}
function displayHighlightedText(words) {
  const displayElement = document.getElementById('highlighted-text');
  if (displayElement) {
    displayElement.innerHTML = ''; // Clear previous content
    const loader = document.getElementById('loader');
    loader.style.display = 'block';
    // Group lay terms by their clinical term
    const groupedTerms = {};
    const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
    words.forEach(word => {
      store.layTerms.forEach(term => {
        if (term.ClinicalTerm.toLowerCase() === word.toLowerCase()) {
          if (!groupedTerms[term.ClinicalTerm]) {
            groupedTerms[term.ClinicalTerm] = [];
          }
          if (!groupedTerms[term.ClinicalTerm].includes(term.LayTerm)) {
            groupedTerms[term.ClinicalTerm].push(term.LayTerm);
          }
        }
      });
    });

    // Create a box for each clinical term
    Object.keys(groupedTerms).forEach(clinicalTerm => {
      // Create the main box for the clinical term
      const mainBox = document.createElement('div');
      mainBox.className = 'box'; // Add box class for styling

      // Create a heading for the clinical term
      const heading = document.createElement('h3');
      heading.textContent = `${clinicalTerm} (${store.glossaryName})`;
      mainBox.appendChild(heading);

      // Create sub-boxes for each lay term
      groupedTerms[clinicalTerm].forEach(layTerm => {
        const subBox = document.createElement('div');
        subBox.className = 'sub-box'; // Add class for sub-box styling
        subBox.textContent = layTerm;

        // Add click event listener to replace ClinicalTerm with LayTerm
        subBox.addEventListener('click', async () => {
          await replaceClinicalTerm(clinicalTerm, layTerm);

          // Remove the main box containing the clicked sub-box
          mainBox.remove();
        });
        mainBox.appendChild(subBox);
      });
      displayElement.appendChild(mainBox);
    });
    loader.style.display = 'none';
  }
}
async function replaceClinicalTerm(clinicalTerm, layTerm) {
  const displayElement = document.getElementById('loader');
  displayElement.style.display = 'block';
  try {
    await Word.run(async context => {
      // Get the current selection
      const selection = context.document.getSelection();
      selection.load('text');
      await context.sync();
      if (selection.text.toLowerCase().includes(clinicalTerm.toLowerCase())) {
        // Search for the clinicalTerm in the document
        const searchResults = selection.search(clinicalTerm, {
          matchCase: false,
          matchWholeWord: false
        });
        searchResults.load('items');
        await context.sync();

        // Replace each occurrence of the clinicalTerm with the layTerm
        for (const item of searchResults.items) {
          // Load the font properties
          item.font.load(['bold', 'italic', 'underline', 'color', 'highlightColor', 'size', 'name']);
          await context.sync(); // Ensure the properties are loaded before accessing them

          // Insert the layTerm while keeping the formatting
          item.insertText(layTerm, Word.InsertLocation.replace);

          // Apply the original formatting to the new text
          item.font.bold = item.font.bold;
          item.font.italic = item.font.italic;
          item.font.underline = item.font.underline;
          item.font.color = item.font.color;
          item.font.highlightColor = '#c7c7c7';
          item.font.size = item.font.size;
          item.font.name = item.font.name;
        }
        await context.sync();
        displayElement.style.display = 'none';
        console.log(`Replaced '${clinicalTerm}' with '${layTerm}' and preserved the original formatting.`);
      } else {
        displayElement.style.display = 'none';
        console.log(`Selected text does not contain '${clinicalTerm}'.`);
      }
    });
  } catch (error) {
    displayElement.style.display = 'none';
    console.error('Error replacing term:', error);
  }
}
async function removeMatchingContentControls() {
  try {
    await Word.run(async context => {
      document.getElementById('app-body').innerHTML = `
      <div id="button-container">
        <div class="loader" id="loader"></div>
        <div id="highlighted-text"></div>`;
      const body = context.document.body;

      // Load all content controls
      const contentControls = body.contentControls;
      contentControls.load("items");
      await context.sync();
      if (contentControls.items.length === 0) {
        console.log("No content controls found.");
        return;
      }
      for (const control of contentControls.items) {
        const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
        if (control.title && store.filteredGlossaryTerm.some(term => term.ClinicalTerm.toLowerCase() === control.title.toLowerCase())) {
          const range = control.getRange();
          range.load("text");
          await context.sync();
          if (control.tag && /^#[0-9A-Fa-f]{6}$/.test(control.tag)) {
            range.font.highlightColor = control.tag;
          } else {
            range.font.highlightColor = null;
          }
          await context.sync();
          control.delete(true);
        }
      }
      document.getElementById('app-body').innerHTML = `
      <div id="button-container">
        <button class="btn btn-secondary me-2 mark-glossary btn-sm" id="applyglossary">Apply Glossary</button>
      </div>
      `;
      await context.sync();
      const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
      store.isGlossaryActive = false;
      document.getElementById('applyglossary').addEventListener('click', applyglossary);
    });
  } catch (error) {
    console.error("Error removing content controls:", error);
  }
}
async function addGenAITags() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  if (!store.isTagUpdating) {
    if (store.isGlossaryActive) {
      await removeMatchingContentControls();
    }
    let selectedClient = store.clientList.filter(item => item.ID === store.clientId);

    // Build Primary Source List
    let sourceTypeList = [...Array.from(new Map(store.dataList.SourceTypeList.filter(item => item.VectorID > 0).map(item => [item.SourceTypeID, {
      Name: item.SourceType,
      ID: item.SourceTypeID
    }])).values())];
    let sourceOptions = sourceTypeList.map(src => {
      return `
        <li class="source-dropdown-item dropdown-item p-2" style="cursor: pointer;">
          <div class="form-check">
            <input class="form-check-input" type="checkbox" value="${src.ID}" id="source${src.ID}">
            <label class="form-check-label text-prewrap" for="source${src.ID}">${src.Name}</label>
          </div>
        </li>`;
    }).join("");
    let sponsorOptions = store.clientList.map(client => {
      const isSelectedClient = selectedClient.some(selected => selected.ID === client.ID);
      return `
        <li class="sponsor-dropdown-item dropdown-item p-2" style="cursor: pointer;">
          <div class="form-check">
            <input class="form-check-input" type="checkbox" value="${client.ID}" id="sponsor${client.ID}" ${isSelectedClient ? 'checked disabled' : ''}>
            <label class="form-check-label text-prewrap" for="sponsor${client.ID}">${client.Name}</label>
          </div>
        </li>`;
    }).join("");
    document.getElementById('app-body').innerHTML = _components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.navTabs;

    // Inject modal
    document.getElementById('add-tag-body').innerHTML = (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.addtagbody)(sponsorOptions, sourceOptions, store.mode === "Summary");
    const promptTemplateElement = document.getElementById('add-prompt-template');
    (0,_draft_home__WEBPACK_IMPORTED_MODULE_6__.setupPromptBuilderUI)(promptTemplateElement, store.promptBuilderList);
    document.getElementById('tag-tab').addEventListener('click', () => (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.switchToAddTag)());
    document.getElementById('prompt-tab').addEventListener('click', () => (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.switchToPromptBuilder)());
    mentionDropdownFn('prompt', 'mention-dropdown', 'add');
    const form = document.getElementById('genai-form');
    const nameField = document.getElementById('name');
    const descriptionField = document.getElementById('description');
    const promptField = document.getElementById('prompt');
    // const primarySourceField = document.getElementById('primarySource');

    const saveGloballyCheckbox = document.getElementById('saveGlobally');
    const availableForAllCheckbox = document.getElementById('isAvailableForAll');
    const sponsorDropdownButton = document.getElementById('sponsorDropdown');
    const sponsorDropdownItems = document.querySelectorAll('.sponsor-dropdown-item .form-check-input');
    const sourceDropdownButton = document.getElementById('sourceDropdown');
    const sourceDropdownItems = document.querySelectorAll('.source-dropdown-item .form-check-input');
    const isSummaryMode = store.mode === "Summary";
    document.getElementById('cancel-btn-gen-ai').addEventListener('click', () => {
      const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
      if (!store.isPendingResponse) (0,_draft_home__WEBPACK_IMPORTED_MODULE_6__.loadHomepage)(store.availableKeys);
    });
    if (form && nameField && promptField && sponsorDropdownItems.length > 0 && (isSummaryMode || sourceDropdownItems.length > 0)) {
      const updateSponsorSelectAllState = () => {
        const selectAllCb = document.getElementById('sponsorSelectAll');
        if (!selectAllCb) return;
        const individualSponsors = Array.from(sponsorDropdownItems).filter(cb => cb.id !== 'sponsorSelectAll');
        if (individualSponsors.length === 0) {
          selectAllCb.checked = false;
          return;
        }
        const allChecked = individualSponsors.every(cb => cb.checked);
        selectAllCb.checked = allChecked;
      };
      const updateSourceSelectAllState = () => {
        const selectAllCb = document.getElementById('sourceSelectAll');
        if (!selectAllCb) return;
        const individualSources = Array.from(sourceDropdownItems).filter(cb => cb.id !== 'sourceSelectAll');
        if (individualSources.length === 0) {
          selectAllCb.checked = false;
          return;
        }
        const allChecked = individualSources.every(cb => cb.checked);
        selectAllCb.checked = allChecked;
      };
      const updateSponsorDropdownLabel = () => {
        if (availableForAllCheckbox.checked) {
          sponsorDropdownButton.textContent = store.clientList.map(x => x.Name).join(", ");
        } else {
          const selectedNames = Array.from(sponsorDropdownItems).filter(cb => cb.checked && cb.id !== 'sponsorSelectAll').map(cb => cb.parentElement.textContent.trim());
          sponsorDropdownButton.textContent = selectedNames.length ? selectedNames.join(", ") : "Select Sponsors";
        }
        updateSponsorSelectAllState();
      };
      const updateSourceDropdownLabel = () => {
        const selectedNames = Array.from(sourceDropdownItems).filter(cb => cb.checked && cb.id !== 'sourceSelectAll').map(cb => cb.parentElement.querySelector('label').textContent.trim());
        const labelSpan = document.getElementById('sourceDropdownLabel');
        if (labelSpan) {
          labelSpan.textContent = selectedNames.length ? selectedNames.join(", ") : "Select Source Types";
        }
        updateSourceSelectAllState();
      };

      // Submit Handler
      form.addEventListener('submit', async e => {
        e.preventDefault();
        form.querySelectorAll('.is-invalid').forEach(i => i.classList.remove('is-invalid'));
        let valid = true;
        if (!nameField.value.trim()) {
          nameField.classList.add('is-invalid');
          valid = false;
        }
        if (!promptField.value.trim()) {
          promptField.classList.add('is-invalid');
          valid = false;
        }

        // SOURCE VALIDATION
        let selectedPrimarySources = [];
        selectedPrimarySources = Array.from(sourceDropdownItems).filter(cb_node => cb_node.checked && cb_node.id !== 'sourceSelectAll').map(cb_node => cb_node.value);
        if (!selectedPrimarySources.length && !isSummaryMode) {
          document.getElementById("primarySourceError").style.display = "block";
          valid = false;
        } else {
          document.getElementById("primarySourceError").style.display = "none";
        }
        if (!valid) return;
        const selectedSponsors = Array.from(sponsorDropdownItems).filter(cb_node => cb_node.checked && cb_node.id !== 'sponsorSelectAll').map(cb_node => store.clientList.find(c => c.ID == cb_node.value));
        const selectedSources = Array.from(sourceDropdownItems).filter(cb => cb.checked && cb.id !== 'sourceSelectAll').map(cb => sourceTypeList.find(s => s.ID == cb.value));
        const isAvailableForAll = availableForAllCheckbox.checked;
        const isSaveGlobally = saveGloballyCheckbox.checked;
        const aigroup = store.dataList.Group.find(el => el.DisplayName === 'AIGroup');
        const formData = {
          DisplayName: nameField.value.trim(),
          Prompt: promptField.value.trim(),
          Description: descriptionField.value.trim(),
          GroupKeyClient: selectedSponsors,
          AllClient: isAvailableForAll ? 1 : 0,
          SaveGlobally: isSaveGlobally,
          UserDefined: '1',
          ComponentKeyDataTypeID: '1',
          ComponentKeyDataAccessID: '3',
          AIFlag: 1,
          DocumentTypeID: store.dataList.DocumentTypeID,
          ReportHeadID: store.dataList.ID,
          // MULTI SELECT SOURCE TYPE
          SourceTypeID: selectedPrimarySources.join(","),
          SummaryTagClient: isSummaryMode ? selectedSponsors.map(s => ({
            ClientID: s.ID,
            Client: s.Name
          })) : [],
          ReportHeadGroupID: aigroup.ID,
          ReportHeadSourceID: 0
        };
        await createTextGenTag(formData);
      });
      const checkAndDisableSponsors = () => {
        sponsorDropdownItems.forEach(cb => {
          if (!cb.disabled) {
            cb.checked = true;
            cb.disabled = true;
          }
        });
        updateSponsorDropdownLabel();
      };
      const enableSponsors = () => {
        sponsorDropdownItems.forEach(cb_node => {
          const cb = cb_node;
          const isSelectedClient = selectedClient.some(sel => sel.ID === parseInt(cb.value));
          if (!isSelectedClient) cb.disabled = false;
        });
        updateSponsorDropdownLabel();
      };
      saveGloballyCheckbox.addEventListener('change', function () {
        if (!store.isPendingResponse) {
          if (this.checked) {
            availableForAllCheckbox.disabled = false;
            sponsorDropdownButton.disabled = false;
          } else {
            enableSponsors();
            availableForAllCheckbox.checked = false;
            availableForAllCheckbox.disabled = true;
            sponsorDropdownButton.disabled = true;
            sponsorDropdownItems.forEach(cb => {
              if (!cb.disabled) {
                cb.checked = false;
                cb.disabled = false;
              }
            });
            updateSponsorDropdownLabel();
          }
        }
      });
      availableForAllCheckbox.addEventListener('change', function () {
        if (!store.isPendingResponse) {
          this.checked ? checkAndDisableSponsors() : enableSponsors();
        }
      });
      saveGloballyCheckbox.checked = true;
      availableForAllCheckbox.checked = true;
      // ✔ Trigger its logic so sponsor dropdown activates properly
      saveGloballyCheckbox.dispatchEvent(new Event("change"));
      availableForAllCheckbox.dispatchEvent(new Event("change"));
      document.querySelectorAll('.sponsor-dropdown-item').forEach(item => {
        item.addEventListener('click', function (e) {
          e.stopPropagation();
          const checkbox = this.querySelector('.form-check-input');
          if (!checkbox) return;
          const target = e.target;
          if (target !== checkbox && target.tagName !== 'LABEL') {
            if (!checkbox.disabled) {
              checkbox.checked = !checkbox.checked;
            }
          }
          if (checkbox.id === 'sponsorSelectAll') {
            const isChecked = checkbox.checked;
            sponsorDropdownItems.forEach(cb_node => {
              const cb = cb_node;
              if (!cb.disabled) cb.checked = isChecked;
            });
          }
          updateSponsorDropdownLabel();
        });
      });
      document.querySelectorAll('.source-dropdown-item').forEach(item => {
        item.addEventListener('click', function (e) {
          e.stopPropagation();
          const checkbox = this.querySelector('.form-check-input');
          if (!checkbox) return;
          const target = e.target;
          if (target !== checkbox && target.tagName !== 'LABEL') {
            if (!checkbox.disabled) {
              checkbox.checked = !checkbox.checked;
            }
          }
          if (checkbox.id === 'sourceSelectAll') {
            const isChecked = checkbox.checked;
            sourceDropdownItems.forEach(cb => {
              const cbInput = cb;
              if (!cbInput.disabled) {
                cbInput.checked = isChecked;
              }
            });
          }
          updateSourceDropdownLabel();
          const selectedCount = Array.from(sourceDropdownItems).filter(cb => cb.checked).length;
          if (selectedCount === 0 && !isSummaryMode) {
            document.getElementById("primarySourceError").style.display = "block";
          } else {
            document.getElementById("primarySourceError").style.display = "none";
          }
        });
      });
      updateSponsorDropdownLabel();
      updateSourceDropdownLabel();
      [nameField, promptField].forEach(field => {
        field.addEventListener('input', function () {
          const input = this;
          if (input.classList.contains('is-invalid') && input.value.trim()) {
            input.classList.remove('is-invalid');
          }
        });
      });
    } else {
      console.error("Required elements missing.");
    }
  }
}
async function customizeTable(type) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  const container = document.getElementById("confirmation-popup");
  if (!container) return;
  const customStyleName = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem("CustomStyle") || "";
  const defaultStyle = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem("DefaultStyle") || store.tableStyle;
  let styleObj = type === "Custom" ? customStyleName : defaultStyle;
  container.innerHTML = (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.customizeTablePopup)(styleObj, type);
  const cancelBtn = document.getElementById("confirmation-popup-cancel");
  const okBtn = document.getElementById("confirmation-popup-confirm");
  const dropdown = document.getElementById("confirmation-popup-dropdown");
  const tablePreview = document.getElementById("confirmation-popup-table-preview");
  const applyStyle = () => {
    if (!dropdown || !tablePreview) return;
    let styleObj;
    if (type === "Custom") {
      styleObj = store.customTableStyle.find(s => s.Name === dropdown.value);
    } else {
      styleObj = _components_tablestyles__WEBPACK_IMPORTED_MODULE_10__.wordTableStyles.find(s => s.style === dropdown.value);
    }
    if (styleObj && type === 'Pre') {
      // Clear existing styles
      Array.from(tablePreview.rows).forEach(row => {
        Array.from(row.cells).forEach(cell => cell.removeAttribute("style"));
      });
      if (styleObj.tableClass) tablePreview.style.cssText = styleObj.tableClass;
      if (styleObj.headerClass) {
        const thead = tablePreview.querySelector("thead");
        if (thead) {
          Array.from(thead.rows).forEach(row => {
            Array.from(row.cells).forEach(cell => {
              cell.style.cssText = styleObj.headerClass;
            });
          });
        }
      }
      if (styleObj.sideHeader && styleObj.rowClass) {
        Array.from(tablePreview.rows).forEach((row, index) => {
          Array.from(row.cells).forEach((cell, cellIndex) => {
            if (cellIndex === 0 && index !== 0) {
              cell.style.cssText = "font-weight:bold;";
            }
          });
        });
      }
      if (styleObj.format === "empty" && styleObj.rowClass) {
        Array.from(tablePreview.rows).forEach((row, index) => {
          if (index % 2 === 1) row.style.cssText = styleObj.rowClass;
        });
      } else if (styleObj.format === "partial" && styleObj.rowClass) {
        Array.from(tablePreview.rows).forEach((row, index) => {
          Array.from(row.cells).forEach((cell, cellIndex) => {
            if (cellIndex === 0) {
              cell.style.cssText = styleObj.tableClass + "font-weight:bold;";
            } else if (index % 2 === 1) {
              cell.style.cssText = styleObj.rowClass;
            }
          });
        });
      } else if (styleObj.format === "full") {
        Array.from(tablePreview.rows).forEach((row, index) => {
          Array.from(row.cells).forEach((cell, cellIndex) => {
            const headerClass = index === 0 ? styleObj.headerClass : "";
            if (cellIndex === 0 && styleObj.sideHeader) {
              cell.style.cssText = styleObj.tableClass + "font-weight:bold;" + headerClass;
            } else {
              cell.style.cssText = styleObj.tableClass + headerClass;
            }
          });
        });
      }
    } else if (styleObj) {
      tablePreview.innerHTML = styleObj.Preview;
    } else {
      tablePreview.innerHTML = '';
    }
  };

  // Initial preview
  applyStyle();
  dropdown?.addEventListener("change", applyStyle);
  if (cancelBtn) cancelBtn.addEventListener("click", () => container.innerHTML = "");
  if (okBtn && dropdown) {
    okBtn.addEventListener("click", () => {
      if (type === "Custom") {
        const styleObj = store.customTableStyle.find(s => s.Name === dropdown.value);
        store.colorPallete.Header = styleObj.Setting.HeaderColor;
        store.colorPallete.Primary = styleObj.Setting.PrimaryColor;
        store.colorPallete.Secondary = styleObj.Setting.SecondaryColor;
        store.colorPallete.Customize = true;
        store.colorPallete.IsSideHeaderBold = styleObj.Setting.IsSideHeaderBold;
        store.colorPallete.IsHeaderBold = styleObj.Setting.IsHeaderBold;
        _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("CustomStyle", styleObj.Name);
        store.tableStyle = styleObj.Setting.BaseStyle; // stores full object as 
      } else {
        store.colorPallete.Customize = false;
        store.tableStyle = dropdown.value; // normal style string
        _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("DefaultStyle", store.tableStyle);
      }
      _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("colorPallete", JSON.stringify(store.colorPallete));
      _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("tableStyle", store.tableStyle);
      container.innerHTML = "";
    });
  }
}
async function getDocumentParagraphStyles() {
  return Word.run(async context => {
    try {
      const styles = context.document.getStyles();
      styles.load("items/nameLocal,items/type");
      await context.sync();
      const paragraphStyles = styles.items.filter(style => style.type === "Paragraph" || Word.StyleType && style.type === Word.StyleType.paragraph).map(style => style.nameLocal).sort((a, b) => a.localeCompare(b));
      return paragraphStyles.length > 0 ? paragraphStyles : ["Normal", "Body Text", "No Spacing"];
    } catch (error) {
      console.error("Failed to load document styles:", error);
      return ["Normal", "Body Text", "No Spacing"];
    }
  });
}
async function getDocumentStyleDetails(styleName) {
  return Word.run(async context => {
    try {
      const styles = context.document.getStyles();
      const style = styles.getByName(styleName);
      style.load("font/bold,font/italic,font/underline,font/name,font/size,font/color");
      await context.sync();
      const fontUnderline = style.font.underline;
      const hasUnderline = fontUnderline && fontUnderline.toLowerCase() !== "none";
      return {
        bold: style.font.bold || false,
        italic: style.font.italic || false,
        underline: hasUnderline,
        fontFamily: style.font.name || "Calibri",
        size: style.font.size ? `${style.font.size}pt` : "11pt",
        fontColor: style.font.color || "#000000",
        backgroundColor: "transparent"
      };
    } catch (error) {
      console.error(`Failed to load style details for ${styleName}:`, error);
      return {
        bold: false,
        italic: false,
        underline: false,
        fontFamily: "Calibri",
        size: "11pt",
        fontColor: "#000000",
        backgroundColor: "transparent"
      };
    }
  });
}
async function customizeTextStyle() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  const container = document.getElementById("confirmation-popup");
  if (!container) return;
  const popupClass = store.theme === 'Dark' ? 'bg-dark text-light' : 'bg-light text-dark';

  // Show a loading dialog first
  container.innerHTML = `
    <div class="modal show d-block" tabindex="-1">
      <div class="modal-dialog">
        <div class="modal-content ${popupClass}">
          <div class="modal-body text-center p-4">
            <i class="fa fa-spinner fa-spin fa-2x mb-2 text-primary"></i>
            <div>Fetching styles from active document...</div>
          </div>
        </div>
      </div>
    </div>
  `;
  const availableStyles = await getDocumentParagraphStyles();
  const currentStyle = store.defaultTextStyle || "Normal";
  container.innerHTML = (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.customizeTextStylePopup)(currentStyle, availableStyles);
  const cancelBtn = document.getElementById("text-style-popup-cancel");
  const okBtn = document.getElementById("text-style-popup-confirm");
  const dropdown = document.getElementById("text-style-dropdown");
  const preview = document.getElementById("text-style-preview");
  const updatePreview = async () => {
    if (!dropdown || !preview) return;
    const styleName = dropdown.value;
    preview.textContent = "Loading style details...";
    const details = await getDocumentStyleDetails(styleName);
    preview.style.fontWeight = details.bold ? "bold" : "normal";
    preview.style.fontStyle = details.italic ? "italic" : "normal";
    preview.style.textDecoration = details.underline ? "underline" : "none";
    preview.style.fontFamily = details.fontFamily;
    preview.style.fontSize = details.size;
    preview.style.color = details.fontColor;
    preview.style.backgroundColor = details.backgroundColor;
    preview.textContent = `${styleName} Preview`;
  };

  // Initial preview load
  await updatePreview();
  dropdown?.addEventListener("change", updatePreview);
  if (cancelBtn) {
    cancelBtn.addEventListener("click", () => {
      container.innerHTML = "";
    });
  }
  if (okBtn && dropdown) {
    okBtn.addEventListener("click", () => {
      const finalStyle = dropdown.value;
      store.defaultTextStyle = finalStyle;
      _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("defaultTextStyle", store.defaultTextStyle);

      // Clear customized style when default text style is selected
      store.customizedTextStyle = null;
      _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.removeItem("customTextStyleId");
      _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.removeItem("customTextStyle");
      store.saveToStorage();
      (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)("Default text style saved successfully", "success");
      container.innerHTML = "";
    });
  }
}
async function customizeCustomStyle() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  const container = document.getElementById("confirmation-popup");
  if (!container) return;
  if (!store.customTextStylesLoaded) {
    const popupClass = store.theme === 'Dark' ? 'bg-dark text-light' : 'bg-light text-dark';
    container.innerHTML = `
      <div class="modal show d-block" tabindex="-1">
        <div class="modal-dialog">
          <div class="modal-content ${popupClass}">
            <div class="modal-body text-center p-4">
              <i class="fa fa-spinner fa-spin fa-2x mb-2 text-primary"></i>
              <div>Fetching customized styles...</div>
            </div>
          </div>
        </div>
      </div>
    `;
    await getCustomTextStyles();
  }
  const stylesList = store.customizedStyles || [];
  const currentStyleId = store.customizedTextStyle && store.customizedTextStyle.id || _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem("customTextStyleId") || (stylesList[0] ? stylesList[0].id : "");
  container.innerHTML = (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.customizedStylePopup)(currentStyleId, stylesList);
  const cancelBtn = document.getElementById("customized-style-popup-cancel");
  const okBtn = document.getElementById("customized-style-popup-confirm");
  const dropdown = document.getElementById("customized-style-dropdown");
  const preview = document.getElementById("customized-style-preview");
  const applyStylePreview = async () => {
    if (!dropdown || !preview) return;
    const selectedStyle = stylesList.find(s => s.id === dropdown.value);
    if (selectedStyle) {
      const props = selectedStyle.properties;
      if (props) {
        preview.style.fontWeight = props.bold ? "bold" : "normal";
        preview.style.fontStyle = props.italic ? "italic" : "normal";
        preview.style.textDecoration = props.underline ? "underline" : "none";
        preview.style.fontFamily = props.fontFamily;
        preview.style.fontSize = props.size;
        preview.style.color = props.fontColor;
        preview.style.backgroundColor = props.backgroundColor;
        preview.textContent = `${selectedStyle.name} Preview`;
      } else {
        preview.textContent = "Loading style details...";
        try {
          const availableStyles = await getDocumentParagraphStyles();
          const styleNameInWord = availableStyles.find(s => s.toLowerCase() === selectedStyle.name.toLowerCase() || s.toLowerCase() === selectedStyle.id.toLowerCase());
          if (styleNameInWord) {
            const details = await getDocumentStyleDetails(styleNameInWord);
            preview.style.fontWeight = details.bold ? "bold" : "normal";
            preview.style.fontStyle = details.italic ? "italic" : "normal";
            preview.style.textDecoration = details.underline ? "underline" : "none";
            preview.style.fontFamily = details.fontFamily;
            preview.style.fontSize = details.size;
            preview.style.color = details.fontColor;
            preview.style.backgroundColor = details.backgroundColor;
            preview.textContent = `${selectedStyle.name} Preview (Word Style)`;
          } else {
            preview.style.fontWeight = "normal";
            preview.style.fontStyle = "normal";
            preview.style.textDecoration = "none";
            preview.style.fontFamily = "Calibri";
            preview.style.fontSize = "11pt";
            preview.style.color = "#000000";
            preview.style.backgroundColor = "transparent";
            preview.textContent = `${selectedStyle.name} Preview (Not found in Word)`;
          }
        } catch (e) {
          console.error("Error updating preview for Word style:", e);
          preview.textContent = `${selectedStyle.name} Preview (Error loading)`;
        }
      }
    }
  };

  // Initial preview
  applyStylePreview();
  dropdown?.addEventListener("change", applyStylePreview);
  if (cancelBtn) {
    cancelBtn.addEventListener("click", () => {
      container.innerHTML = "";
    });
  }
  if (okBtn && dropdown) {
    okBtn.addEventListener("click", async () => {
      const selectedStyle = stylesList.find(s => s.id === dropdown.value);
      if (selectedStyle) {
        if (selectedStyle.properties === null) {
          try {
            const availableStyles = await getDocumentParagraphStyles();
            const styleNameInWord = availableStyles.find(s => s.toLowerCase() === selectedStyle.name.toLowerCase() || s.toLowerCase() === selectedStyle.id.toLowerCase());
            if (styleNameInWord) {
              store.defaultTextStyle = styleNameInWord;
              _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("defaultTextStyle", styleNameInWord);
            }
          } catch (e) {
            console.error("Error setting default style from Word style:", e);
          }
        }
        store.customizedTextStyle = selectedStyle;
        _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("customTextStyleId", selectedStyle.id);
        _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.setItem("customTextStyle", JSON.stringify(selectedStyle));
        store.saveToStorage();
        (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)("Customized style applied successfully", "success");
      }
      container.innerHTML = "";
    });
  }
}

async function createTextGenTag(payload) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  try {
    const iconelement = document.getElementById(`text-gen-save`);
    const cancelBtnGenAi = document.getElementById('cancel-btn-gen-ai');
    cancelBtnGenAi.disabled = true;
    iconelement.innerHTML = `<i class="fa fa-spinner fa-spin text-white me-2"></i>Save`;
    iconelement.disabled = true;
    store.isPendingResponse = true;
    let data;
    if (store.mode === "Summary") {
      const summaryPayload = {
        ReportHeadID: payload.ReportHeadID,
        Name: payload.DisplayName,
        Description: payload.Description,
        Prompt: payload.Prompt,
        Selected: 1,
        SourceTypeID: payload.SourceTypeID,
        AllClient: payload.AllClient,
        SaveGlobally: payload.SaveGlobally ? 1 : 0,
        SummaryTagClient: payload.SummaryTagClient
      };
      data = await (0,_summary_summary_api__WEBPACK_IMPORTED_MODULE_13__.addSummaryTag)(summaryPayload, store.jwt);
    } else {
      data = await (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_9__.addGroupKey)(payload, store.jwt);
    }
    store.isPendingResponse = false;
    if (data['Status']) {
      if (store.mode === "Summary") {
        (0,_summary_summary__WEBPACK_IMPORTED_MODULE_12__.loadSummarypage)(store.availableKeys);
      } else {
        fetchDocument('AIpanel');
      }
      (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)('Saved successfully', 'success');
    } else {
      cancelBtnGenAi.disabled = false;
      iconelement.disabled = false;
      iconelement.innerHTML = `<i class="fa fa-check-circle me-2"></i>Save`;
      (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)('Something went wrong', 'error');
      // showAddTagError(data['Data']);
    }
  } catch (error) {
    (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)('Something went wrong', 'error');
    console.error('Error creating text generation tag:', error);
  }
}
function mentionDropdownFn(textareaId, DropdownId, action) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  const filterMentions = query => {
    // Assuming availableKeys is an array of objects with DisplayName and EditorValue properties
    const filtered = store.availableKeys.filter(item => item.AIFlag === 0).filter(item => item.DisplayName.toLowerCase().includes(query.toLowerCase()));
    return filtered;
  };
  let highlightedIndex = -1;
  const promptField = document.getElementById(`${textareaId}`);
  const mentionDropdown = document.getElementById(`${DropdownId}`);
  if (promptField) {
    // Handle input events on prompt field for mentions
    promptField.addEventListener('input', e => {
      const cursorPosition = promptField.selectionStart;
      const textBeforeCursor = promptField.value.slice(0, cursorPosition);
      const lastHashtag = textBeforeCursor.lastIndexOf('#');
      if (lastHashtag !== -1) {
        const query = textBeforeCursor.slice(lastHashtag + 1).trim();
        if (query.length > 0) {
          const mentions = filterMentions(query);
          if (mentions.length > 0) {
            mentionDropdown.innerHTML = mentions.map(item => {
              let editorValue = '';
              if (action === 'add') {
                editorValue = `#${item.DisplayName}#`;
              } else {
                editorValue = item.EditorValue || `#${item.DisplayName}#`;
              }
              return `<li class="dropdown-item" data-editor-value="${editorValue}">${item.DisplayName}</li>`;
            }).join('');

            // Get the position of the textarea and place the dropdown above it
            const textareaRect = promptField.getBoundingClientRect();
            mentionDropdown.style.left = `${textareaRect.left}px`;
            mentionDropdown.style.bottom = `75px`; // Position above the textarea
            mentionDropdown.style.display = 'block';
          } else {
            mentionDropdown.style.display = 'none';
          }
        } else {
          mentionDropdown.style.display = 'none';
        }
      } else {
        mentionDropdown.style.display = 'none';
      }
    });

    // Handle keyboard navigation in the dropdown
    promptField.addEventListener('keydown', e => {
      const items = document.querySelectorAll(`#${DropdownId} .dropdown-item`);
      const totalItems = items.length;
      if (e.key === 'ArrowDown') {
        // Prevent default behavior to stop cursor from moving
        e.preventDefault();

        // Move the highlight down and wrap around to the top if at the end
        if (highlightedIndex < totalItems - 1) {
          highlightedIndex++;
        } else {
          highlightedIndex = 0; // Wrap to the first item
        }
        updateHighlightedItem(`${DropdownId}`);
      } else if (e.key === 'ArrowUp') {
        // Prevent default behavior to stop cursor from moving
        e.preventDefault();

        // Move the highlight up and wrap around to the bottom if at the top
        if (highlightedIndex > 0) {
          highlightedIndex--;
        } else {
          highlightedIndex = totalItems - 1; // Wrap to the last item
        }
        updateHighlightedItem(`${DropdownId}`);
      } else if (e.key === 'Enter' && highlightedIndex !== -1) {
        // Select the highlighted item
        const selectedItem = items[highlightedIndex];
        if (selectedItem) {
          selectMention(selectedItem.getAttribute('data-editor-value'));
          mentionDropdown.style.display = 'none'; // Hide the dropdown after selection
          e.preventDefault(); // Prevent form submission on Enter key
        }
      }
    });

    // Function to highlight the selected item
    function updateHighlightedItem(id) {
      const items = document.querySelectorAll(`#${id} .dropdown-item`);
      const dropdown = document.getElementById(`${id}`);
      const totalItems = items.length;

      // Remove the 'active' class from all items
      items.forEach(item => item.classList.remove('active'));

      // Add the 'active' class to the currently highlighted item
      if (highlightedIndex >= 0 && highlightedIndex < totalItems) {
        const highlightedItem = items[highlightedIndex];
        highlightedItem.classList.add('active');

        // Ensure the highlighted item is visible within the dropdown
        highlightedItem.scrollIntoView({
          behavior: 'smooth',
          // Smooth scroll
          block: 'nearest' // Scroll only if necessary
        });
      }
    }

    // Handle selecting an item from the dropdown via mouse click
    mentionDropdown.addEventListener('click', e => {
      if (e.target && e.target.matches('li')) {
        const editorValue = e.target.getAttribute('data-editor-value');
        selectMention(editorValue);
        mentionDropdown.style.display = 'none'; // Hide the dropdown after selection
      }
    });

    // Function to insert the selected mention into the prompt field
    const selectMention = editorValue => {
      const textarea = document.getElementById(`${textareaId}`);
      const currentValue = textarea.value;
      const cursorPosition = textarea.selectionStart;
      const textBefore = currentValue.slice(0, cursorPosition);
      const textAfter = currentValue.slice(cursorPosition);
      const lastHashPosition = textBefore.lastIndexOf('#');
      const updatedTextBefore = textBefore.slice(0, lastHashPosition); // Removing '#' symbol

      textarea.value = `${updatedTextBefore}${editorValue}${textAfter}`;
      const newCursorPosition = updatedTextBefore.length + editorValue.length;
      textarea.setSelectionRange(newCursorPosition, newCursorPosition);
    };

    // Hide the dropdown if clicked outside
    document.addEventListener('click', e => {
      if (!mentionDropdown.contains(e.target) && e.target !== promptField) {
        mentionDropdown.style.display = 'none';
      }
    });
  }
}
function removeQuotes(value) {
  return value ? value.replace(/^"|"$/g, '').replace(/\\n/g, '').replace(/\*\*/g, '').replace(/\\r/g, '') : '';
}
function createMultiSelectDropdown(tag, type) {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  const isDark = store.theme === 'Dark';
  const btnClass = isDark ? 'btn-dark text-light border-0' : 'btn-light text-dark border';
  const dropdownMenuClass = isDark ? 'bg-dark text-light border-light' : 'bg-white text-dark border';
  const itemClass = isDark ? 'bg-dark text-light' : 'bg-white text-dark';

  // Group sources by SourceType
  const sourceList = type === 'Summary' ? store.sourceSummaryList : store.sourceList;
  const groupedSources = sourceList.reduce((groups, source) => {
    if (!groups[source.SourceType]) groups[source.SourceType] = [];
    groups[source.SourceType].push(source);
    return groups;
  }, {});
  const multiSelectHTML = `
  <div class='p-3 w-100'>
    <div class="mb-3">
      <label for="source-select" class="form-label"><span class="text-danger">*</span> Select Sources</label>
      <div class="dropdown w-100">
        <button 
          class="btn ${btnClass} w-100 text-start d-flex justify-content-between align-items-start dropdown-toggle dropdown-toggle-sources" 
          type="button" 
          id="sourceDropdown" 
          data-bs-toggle="dropdown" 
          aria-expanded="false">
          <span id="sourceDropdownLabel" class='sourceDropdownLabel'></span>
          <span class="dropdown-toggle-icon dropdown-toggle-icon-s"></span>
        </button>
        <ul class="dropdown-menu ${dropdownMenuClass} w-100 p-2" style="box-shadow: 0 4px 8px rgba(0,0,0,0.1); z-index: 10000; max-height: 300px; overflow-y: auto;">
          
          <!-- Select All -->
          <li class="dropdown-item p-2 ${itemClass}" data-checkbox-id="selectAll">
            <div class="form-check">
              <input class="form-check-input" type="checkbox" value="selectAll" id="selectAll">
              <label class="form-check-label w-100" for="selectAll">Select All</label>
            </div>
          </li>

          <!-- Grouped Sources -->
          ${Object.keys(groupedSources).map((group, groupIndex) => {
    const groupItems = groupedSources[group].map((source, index) => `
                  <li class="dropdown-item ps-4 ${itemClass}" style="cursor: pointer;" data-checkbox-id="source-${groupIndex}-${index}">
                    <div class="form-check">
                      <input class="form-check-input source-checkbox" type="checkbox" value="${type === 'Summary' ? source.FileName : source.SourceName}" id="source-${groupIndex}-${index}">
                      <label class="form-check-label w-100 text-prewrap" for="source-${groupIndex}-${index}">${type === 'Summary' ? source.FileName : source.SourceName}</label>
                    </div>
                  </li>
                `).join('');
    return `
                <!-- Group Header -->
                <li class="dropdown-item p-2 ${itemClass}" data-group-id="group-${groupIndex}">
                  <div class="form-check">
                    <input class="form-check-input group-checkbox" type="checkbox" value="${group}" id="group-${groupIndex}">
                    <label class="form-check-label fw-bold" for="group-${groupIndex}">${group}</label>
                  </div>
                </li>
                ${groupItems}
              `;
  }).join('')}
        </ul>
      </div>
    </div>
    <div class="mt-3 d-flex justify-content-between">
      <span id="cancel-src-btn" class="fw-bold text-primary my-auto c-pointer">Cancel</span>
      <button id="ok-src-btn" class="btn btn-primary">Save</button>
    </div>
  </div>
  `;
  const accordionBody = document.getElementById(`chatFooter`);
  accordionBody.innerHTML = multiSelectHTML;
  let selectedSources = [];
  const selectAllCheckbox = document.getElementById(`selectAll`);
  const groupCheckboxes = document.querySelectorAll(`.group-checkbox`);
  const individualCheckboxes = document.querySelectorAll(`.source-checkbox`);
  const sourceDropdownLabel = document.getElementById(`sourceDropdownLabel`);
  function updateLabel() {
    sourceDropdownLabel.innerText = selectedSources.length > 0 ? selectedSources.join(', ') : ' ';
  }

  // Select All logic
  selectAllCheckbox.addEventListener("change", function () {
    const checked = this.checked;
    groupCheckboxes.forEach(cb => cb.checked = checked);
    individualCheckboxes.forEach(cb => {
      cb.checked = checked;
      if (checked && !selectedSources.includes(cb.value)) {
        selectedSources.push(cb.value);
      }
      if (!checked) {
        selectedSources = [];
      }
    });
    updateLabel();
  });

  // Group checkbox logic
  groupCheckboxes.forEach(groupCb => {
    groupCb.addEventListener("change", function () {
      const groupIndex = this.id.split('-')[1];
      const groupItems = document.querySelectorAll(`[data-checkbox-id^="source-${groupIndex}-"] .source-checkbox`);
      groupItems.forEach(cb => {
        cb.checked = this.checked;
        if (this.checked && !selectedSources.includes(cb.value)) {
          selectedSources.push(cb.value);
        }
        if (!this.checked) {
          selectedSources = selectedSources.filter(s => s !== cb.value);
        }
      });

      // Update Select All state
      selectAllCheckbox.checked = Array.from(individualCheckboxes).every(child => child.checked);
      updateLabel();
    });
  });

  // Individual checkbox logic
  individualCheckboxes.forEach(cb => {
    cb.addEventListener("change", function () {
      if (cb.checked) {
        if (!selectedSources.includes(cb.value)) selectedSources.push(cb.value);
      } else {
        selectedSources = selectedSources.filter(s => s !== cb.value);
      }

      // Update parent group checkbox
      const groupIndex = cb.id.split("-")[1];
      const groupItems = document.querySelectorAll(`[data-checkbox-id^="source-${groupIndex}-"] .source-checkbox`);
      const groupCheckbox = document.getElementById(`group-${groupIndex}`);
      groupCheckbox.checked = Array.from(groupItems).every(child => child.checked);

      // Update Select All checkbox
      selectAllCheckbox.checked = Array.from(individualCheckboxes).every(child => child.checked);
      updateLabel();
    });
  });

  // Initialize with pre-selected sources
  if (tag.Sources && tag.Sources.length > 0) {
    individualCheckboxes.forEach(cb => {
      if (tag.Sources.includes(cb.value)) {
        cb.checked = true;
        selectedSources.push(cb.value);
      }
    });

    // Update group checkboxes
    groupCheckboxes.forEach(groupCb => {
      const groupIndex = groupCb.id.split("-")[1];
      const groupItems = document.querySelectorAll(`[data-checkbox-id^="source-${groupIndex}-"] .source-checkbox`);
      groupCb.checked = Array.from(groupItems).every(child => child.checked);
    });

    // Update Select All
    selectAllCheckbox.checked = Array.from(individualCheckboxes).every(child => child.checked);
    updateLabel();
  }

  // Save
  document.getElementById(`ok-src-btn`).addEventListener("click", function () {
    tag.Sources = [...selectedSources];
    const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
    const receivedEntry = sourceList.filter(source => selectedSources.includes(type === 'Summary' ? source.FileName : source.SourceName));
    tag.TempSourceValue = receivedEntry.map(item => {
      return item.VectorID ? String(item.VectorID) : item.SourceValue;
    });
    if (type === 'Summary') {
      tag.FileName = receivedEntry.map(item => {
        return item.FileName;
      });
    } else {
      tag.SourceName = receivedEntry.map(item => {
        return item.SourceName;
      });
    }
    tag.SourceValueID = receivedEntry.map(item => {
      return String(item.VectorID);
    });
    tag.SourceValue = receivedEntry.map(source => source.SourceValue);
    accordionBody.innerHTML = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.chatfooter)(tag);
    (0,_draft_home__WEBPACK_IMPORTED_MODULE_6__.initializeAIHistoryEvents)(tag, store.jwt, store.availableKeys, type);
  });

  // Cancel
  document.getElementById(`cancel-src-btn`).addEventListener("click", function () {
    accordionBody.innerHTML = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.chatfooter)(tag);
    (0,_draft_home__WEBPACK_IMPORTED_MODULE_6__.initializeAIHistoryEvents)(tag, store.jwt, store.availableKeys, type);
  });
}
async function loadPromptTemplates() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  try {
    const data = await (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_9__.getAllPromptTemplates)(store.jwt);
    if (data.Status && data.Data) {
      store.promptBuilderList = data.Data;
    }
    // Do something with the data
  } catch (error) {
    console.error('Error fetching prompt templates:', error);
  }
}
async function logBookmarksInSelection() {
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();

  // if (store.currentChatTagId !== -1 && store.currentChatTagId !== undefined && store.currentChatTagId !== null) {
  //   return;
  // }
  return Word.run(async context => {
    const selection = context.document.getSelection();
    const rawBookmarks = await getBookmarksFromSelection(context); // internally calls context.sync()
    const bookmarks = pickRelevantBookmarks(rawBookmarks);
    try {
      const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
      if (store.mode === 'Home' || store.mode === 'Summary') {
        selection.load('text');
        const para = selection.paragraphs.getFirst();
        const paraStartRange = para.getRange('Start');
        const cursorStartRange = selection.getRange('Start');
        const beforeCursorRange = paraStartRange.expandTo(cursorStartRange);
        para.load('text');
        beforeCursorRange.load('text');
        await context.sync();
        const selectedText = selection.text || '';
        const paraText = para.text || '';
        const cursorOffset = (beforeCursorRange.text || '').length;
        const tagRegex = /#([^#\r\n]+)#/g;
        let match;
        const matchedNames = [...bookmarks];

        // Helper to find matching tag in current mode
        const findTag = tagName => {
          if (store.mode === 'Home') {
            return store.availableKeys.find(k => k.AIFlag === 1 && (k.DisplayName.toLowerCase() === tagName.toLowerCase() || `id${k.ID}`.toLowerCase() === tagName.toLowerCase()));
          } else if (store.mode === 'Summary') {
            return store.summaryTagList?.find(k => k.Name?.toLowerCase() === tagName.toLowerCase() || `sm${k.ID || k.ReportHeadSummaryTagID}`.toLowerCase() === tagName.toLowerCase());
          }
          return null;
        };

        // Case 1: Selection is not empty, check if it contains matching keys
        if (selectedText.trim().length > 0) {
          while ((match = tagRegex.exec(selectedText)) !== null) {
            const tagName = match[1].trim();
            const tag = findTag(tagName);
            if (tag) {
              const nameToPush = store.mode === 'Home' ? tag.DisplayName : tag.Name;
              if (nameToPush) {
                const isAlreadyAdded = matchedNames.some(existingName => {
                  if (store.mode === 'Home') {
                    return existingName.toLowerCase() === nameToPush.toLowerCase() || existingName.toLowerCase() === `id${tag.ID}`.toLowerCase();
                  } else {
                    return existingName.toLowerCase() === nameToPush.toLowerCase() || existingName.toLowerCase() === `sm${tag.ID || tag.ReportHeadSummaryTagID}`.toLowerCase();
                  }
                });
                if (!isAlreadyAdded) {
                  matchedNames.push(nameToPush);
                }
              }
            }
          }
        }

        // Case 2: Selection was empty or no tags were found inside selection,
        // look at the tag where the cursor is currently placed in the paragraph
        if (matchedNames.length === 0) {
          tagRegex.lastIndex = 0;
          while ((match = tagRegex.exec(paraText)) !== null) {
            const tagStart = match.index;
            const tagEnd = match.index + match[0].length;
            if (cursorOffset >= tagStart && cursorOffset <= tagEnd) {
              const tagName = match[1].trim();
              const tag = findTag(tagName);
              if (tag) {
                const nameToPush = store.mode === 'Home' ? tag.DisplayName : tag.Name;
                if (nameToPush) {
                  matchedNames.push(nameToPush);
                }
                break;
              }
            }
          }
        }
        if (matchedNames.length > 1) {
          document.getElementById('tags-in-selected-text')?.classList.replace('d-none', 'd-block');
          store.selectedNames = matchedNames;
          (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.renderSelectedTags)(store.selectedNames, store.availableKeys);
          return;
        } else if (matchedNames.length === 1) {
          const singleName = matchedNames[0];
          const tag = findTag(singleName);
          if (tag) {
            const tagId = tag.ID || tag.ReportHeadSummaryTagID;
            if (store.currentChatTagId !== -1 && store.currentChatTagId !== undefined && store.currentChatTagId !== null && String(store.currentChatTagId) === String(tagId)) {
              document.getElementById('tags-in-selected-text')?.classList.replace('d-none', 'd-block');
              store.selectedNames = [singleName];
              (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.renderSelectedTags)(store.selectedNames, store.availableKeys);
              return;
            }
            (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.confirmSwitchChatHistory)(async () => {
              const appBody = document.getElementById('app-body');
              appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';
              await (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.selectMatchingBookmarkFromSelection)(singleName);
              if (store.mode === 'Home') {
                appBody.innerHTML = await (0,_draft_home__WEBPACK_IMPORTED_MODULE_6__.generateCheckboxHistory)(tag, "AITag");
              } else if (store.mode === 'Summary') {
                appBody.innerHTML = await (0,_draft_home__WEBPACK_IMPORTED_MODULE_6__.generateCheckboxHistory)(tag, "Summary");
              }
              document.getElementById('tags-in-selected-text')?.classList.replace('d-none', 'd-block');
              store.selectedNames = [singleName];
              (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.renderSelectedTags)(store.selectedNames, store.availableKeys);
            });
            return;
          }
        }
      }
    } catch (e) {
      console.error('Tag placeholder detection error:', e);
    }
    document.getElementById('tags-in-selected-text')?.classList.replace('d-block', 'd-none');
  });
}
async function getBookmarksFromSelection(context) {
  const selection = context.document.getSelection();
  const bookmarks = selection.getBookmarks();
  await context.sync();
  return bookmarks.value || [];
}
function normalizeBookmark(name) {
  return name.split('_Split_')[0].replace(/_/g, ' ');
}
function pickRelevantBookmarks(bookmarks) {
  // Remove duplicates & internal splits
  const normalized = Array.from(new Set(bookmarks.map(normalizeBookmark)));

  // Prefer AI tags only
  const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
  return normalized.filter(name => {
    if (store.mode === 'Home') {
      return store.availableKeys.some(k => k.AIFlag === 1 && (k.DisplayName.toLowerCase() === name.toLowerCase() || `id${k.ID}`.toLowerCase() === name.toLowerCase()));
    } else if (store.mode === 'Summary') {
      return store.summaryTagList.some(k => k.Name?.toLowerCase() === name.toLowerCase() || `sm${k.ID || k.ReportHeadSummaryTagID}`.toLowerCase() === name.toLowerCase());
    }
    return false;
  });
}
async function getImages() {
  try {
    const store = _services_store_service__WEBPACK_IMPORTED_MODULE_4__.StoreService.getInstance();
    const userId = _utils_doc_storage__WEBPACK_IMPORTED_MODULE_5__.DocStorage.getItem('userId') || '0';

    // Fetch Images and Clients in parallel
    const generalImagesPromise = (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_9__.getGeneralImages)(store.jwt);
    const documentImagesPromise = (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_9__.getReportHeadImageById)(store.dataList.ID, store.jwt);
    const clientsPromise = (0,_draft_draft_api__WEBPACK_IMPORTED_MODULE_9__.getAllClients)(userId, store.jwt);
    const [generalImages, documentImages, clientsData] = await Promise.all([generalImagesPromise, documentImagesPromise, clientsPromise]);
    const mappedGeneral = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.mapImagesToComponentObjects)(generalImages['Data']);
    const mappedDocument = (0,_draft_draft_functions__WEBPACK_IMPORTED_MODULE_7__.mapImagesToComponentObjects)(documentImages['Data']);

    // Update Store with Images
    store.dataList.GroupKeyAll.push(...mappedGeneral);
    store.dataList.GroupKeyAll.push(...mappedDocument);
    store.availableKeys.push(...mappedGeneral);
    store.availableKeys.push(...mappedDocument);
    store.imageList = store.dataList.GroupKeyAll.filter(element => element.ComponentKeyDataType === 'IMAGE');

    // Update Store with Clients
    if (clientsData.Status && clientsData.Data) {
      store.clientList = clientsData.Data;
    }
    if (store.imageList && store.imageList.length > 0) {
      (0,_components_bodyelements__WEBPACK_IMPORTED_MODULE_8__.toaster)('Images and data are loaded and ready for use', 'success');
    }
  } catch (error) {
    console.error("Error loading background data:", error);
  }
}

/***/ }),

/***/ "./src/taskpane/utils/config.ts":
/*!**************************************!*\
  !*** ./src/taskpane/utils/config.ts ***!
  \**************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   CONFIG: function() { return /* binding */ CONFIG; }
/* harmony export */ });
const CONFIG = {
  dataUrl: 'https://plsdevapp.azurewebsites.net',
  storeUrl: 'https://linkwordplugin-aphgcwcgbfdqeccs.eastus-01.azurewebsites.net',
  version: '2.5.3',
  environment: ['Dev']
};

/***/ }),

/***/ "./src/taskpane/utils/doc-storage.ts":
/*!*******************************************!*\
  !*** ./src/taskpane/utils/doc-storage.ts ***!
  \*******************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   DocStorage: function() { return /* binding */ DocStorage; },
/* harmony export */   setDocStorageId: function() { return /* binding */ setDocStorageId; }
/* harmony export */ });
/**
 * DocStorage – a localStorage wrapper that scopes every key to the current
 * document ID, so that two add-in instances opened against different documents
 * never share credentials or settings.
 *
 * Usage:  DocStorage.getItem('token')
 *         DocStorage.setItem('token', jwt)
 *         DocStorage.removeItem('token')
 *         DocStorage.clearAll()   ← clears only THIS document's keys
 *
 * The document ID is read lazily from StoreService each time a method is
 * called, so it is always up-to-date even if the store is initialised after
 * the first import.
 */

let _docId = '';

/** Called once in taskpane.ts after documentID is known. */
function setDocStorageId(id) {
  _docId = id || 'default';
}
function prefix(key) {
  return `${_docId || 'default'}__${key}`;
}
const DocStorage = {
  getItem(key) {
    return localStorage.getItem(prefix(key));
  },
  setItem(key, value) {
    localStorage.setItem(prefix(key), value);
  },
  removeItem(key) {
    localStorage.removeItem(prefix(key));
  },
  /** Remove every localStorage entry that belongs to the current document. */
  clearAll() {
    const pfx = prefix('');
    const toDelete = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k && k.startsWith(pfx)) {
        toDelete.push(k);
      }
    }
    toDelete.forEach(k => localStorage.removeItem(k));
  }
};

/***/ }),

/***/ "./src/taskpane/utils/fontawesome-icons.ts":
/*!*************************************************!*\
  !*** ./src/taskpane/utils/fontawesome-icons.ts ***!
  \*************************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   faFolderGear: function() { return /* binding */ faFolderGear; },
/* harmony export */   faMagicWandSparkles: function() { return /* binding */ faMagicWandSparkles; },
/* harmony export */   faMicrochipAi: function() { return /* binding */ faMicrochipAi; },
/* harmony export */   faTableLayout: function() { return /* binding */ faTableLayout; },
/* harmony export */   getIconSvg: function() { return /* binding */ getIconSvg; }
/* harmony export */ });
function getIconSvg(iconDef, extraClass = '', extraStyle = '', extraAttrs = '') {
  const [width, height,,, pathData] = iconDef.icon;
  const paths = Array.isArray(pathData) ? pathData : [pathData];
  const pathElements = paths.map(p => `<path fill="currentColor" d="${p}"></path>`).join('');
  const styleAttr = extraStyle ? `style="${extraStyle}"` : '';
  return `<svg class="svg-inline--fa ${extraClass}" aria-hidden="true" focusable="false" role="img" viewBox="0 0 ${width} ${height}" ${styleAttr} ${extraAttrs}>${pathElements}</svg>`;
}
const faFolderGear = {
  prefix: 'far',
  iconName: 'folder-gear',
  icon: [512, 512, ['folder-cog'], 'e187', 'M448 96h-172.1L226.7 50.75C214.7 38.74 198.5 32 181.5 32H64C28.65 32 0 60.66 0 96v320c0 35.34 28.65 64 64 64h384c35.35 0 64-28.66 64-64V160C512 124.7 483.3 96 448 96zM464 416c0 8.824-7.178 16-16 16H64c-8.822 0-16-7.176-16-16V96c0-8.824 7.178-16 16-16h117.5c4.273 0 8.293 1.664 11.31 4.688L256 144h192c8.822 0 16 7.176 16 16V416zM338.5 303.3C339.4 298.3 340 293.2 340 288s-.625-10.31-1.541-15.28L358.8 261c2.811-1.625 4.207-4.984 3.229-8.082C359.7 245.8 356.9 238.8 352.1 232s-8.562-12.73-13.62-18.25c-2.193-2.398-5.805-2.875-8.615-1.246l-20.52 11.85C302.5 217.8 293.7 212.6 284 209.1V185.5c0-3.246-2.215-6.137-5.387-6.836C271.3 177 263.8 176 256 176S240.7 177 233.4 178.7C230.2 179.4 228 182.3 228 185.5v23.64C218.3 212.6 209.5 217.8 201.8 224.4L181.2 212.5C178.4 210.9 174.8 211.4 172.6 213.8C167.6 219.3 162.9 225.2 159 232c-3.898 6.754-6.744 13.78-8.998 20.91C149 256 150.4 259.4 153.2 261l20.3 11.72C172.6 277.7 172 282.8 172 288s.623 10.31 1.539 15.28L153.2 315c-2.812 1.621-4.207 4.984-3.23 8.082C152.3 330.2 155.1 337.2 159 344c3.9 6.754 8.562 12.73 13.62 18.25c2.195 2.398 5.805 2.875 8.617 1.246l20.52-11.85C209.5 358.2 218.3 363.4 228 366.9v23.64c0 3.242 2.215 6.137 5.387 6.836C240.7 398.1 248.2 400 256 400s15.3-1.047 22.61-2.664c3.172-.6992 5.387-3.594 5.387-6.836v-23.64c9.734-3.465 18.54-8.637 26.24-15.21l20.52 11.85c2.811 1.629 6.422 1.152 8.615-1.246c5.055-5.52 9.716-11.5 13.62-18.25s6.745-13.78 8.999-20.92c.9766-3.098-.4199-6.461-3.23-8.082L338.5 303.3zM256 323c-19.33 0-35-15.67-35-35S236.7 253 256 253c19.33 0 35 15.67 35 35S275.3 323 256 323z']
};
const faMicrochipAi = {
  prefix: 'far',
  iconName: 'microchip-ai',
  icon: [512, 512, [], 'e1ec', 'M226.3 183.1c-6.375-14.56-30.28-14.56-36.66 0l-56 128c-4.422 10.12 .1875 21.91 10.31 26.34c10.14 4.406 21.91-.2031 26.34-10.31l5.289-12.09C175.8 315.9 175.9 316 176 316h64c.1348 0 .248-.0762 .3828-.0781l5.289 12.09C248.1 335.5 256.3 340 264 340c2.672 0 5.391-.5313 8-1.672c10.12-4.438 14.73-16.22 10.31-26.34L226.3 183.1zM193.1 276L208 241.9L222.9 276H193.1zM336 172c-11.05 0-20 8.953-20 20v128c0 11.05 8.953 20 20 20s20-8.953 20-20V192C356 180.1 347 172 336 172zM488 280C501.3 280 512 269.3 512 256s-10.75-24-24-24H448v-48h40C501.3 184 512 173.3 512 160s-10.75-24-24-24H448V128c0-35.35-28.65-64-64-64h-8V24C376 10.75 365.3 0 352 0s-24 10.75-24 24V64h-48V24C280 10.75 269.3 0 256 0S232 10.75 232 24V64h-48V24C184 10.75 173.3 0 160 0S136 10.75 136 24V64H128C92.65 64 64 92.65 64 128v8H24C10.75 136 0 146.8 0 160s10.75 24 24 24H64v48H24C10.75 232 0 242.8 0 256s10.75 24 24 24H64v48H24C10.75 328 0 338.8 0 352s10.75 24 24 24H64V384c0 35.35 28.65 64 64 64h8v40C136 501.3 146.8 512 160 512s24-10.75 24-24V448h48v40C232 501.3 242.8 512 256 512s24-10.75 24-24V448h48v40c0 13.25 10.75 24 24 24s24-10.75 24-24V448H384c35.35 0 64-28.65 64-64v-8h40c13.25 0 24-10.75 24-24s-10.75-24-24-24H448v-48H488zM400 384c0 8.822-7.178 16-16 16H128c-8.822 0-16-7.178-16-16V128c0-8.822 7.178-16 16-16h256c8.822 0 16 7.178 16 16V384z']
};
const faMagicWandSparkles = {
  prefix: 'far',
  iconName: 'magic-wand-sparkles',
  icon: [576, 512, ['magic-wand-sparkles'], 'e2ca', 'M248.8 4.994C249.9 1.99 252.8 .0001 256 .0001C259.2 .0001 262.1 1.99 263.2 4.994L277.3 42.67L315 56.79C318 57.92 320 60.79 320 64C320 67.21 318 70.08 315 71.21L277.3 85.33L263.2 123C262.1 126 259.2 128 256 128C252.8 128 249.9 126 248.8 123L234.7 85.33L196.1 71.21C193.1 70.08 192 67.21 192 64C192 60.79 193.1 57.92 196.1 56.79L234.7 42.67L248.8 4.994zM495.3 14.06L529.9 48.64C548.6 67.38 548.6 97.78 529.9 116.5L148.5 497.9C129.8 516.6 99.38 516.6 80.64 497.9L46.06 463.3C27.31 444.6 27.31 414.2 46.06 395.4L427.4 14.06C446.2-4.686 476.6-4.686 495.3 14.06V14.06zM461.4 48L351.7 157.7L386.2 192.3L495.1 82.58L461.4 48zM114.6 463.1L352.3 226.2L317.7 191.7L80 429.4L114.6 463.1zM7.491 117.2L64 96L85.19 39.49C86.88 34.98 91.19 32 96 32C100.8 32 105.1 34.98 106.8 39.49L128 96L184.5 117.2C189 118.9 192 123.2 192 128C192 132.8 189 137.1 184.5 138.8L128 160L106.8 216.5C105.1 221 100.8 224 96 224C91.19 224 86.88 221 85.19 216.5L64 160L7.491 138.8C2.985 137.1 0 132.8 0 128C0 123.2 2.985 118.9 7.491 117.2zM359.5 373.2L416 352L437.2 295.5C438.9 290.1 443.2 288 448 288C452.8 288 457.1 290.1 458.8 295.5L480 352L536.5 373.2C541 374.9 544 379.2 544 384C544 388.8 541 393.1 536.5 394.8L480 416L458.8 472.5C457.1 477 452.8 480 448 480C443.2 480 438.9 477 437.2 472.5L416 416L359.5 394.8C354.1 393.1 352 388.8 352 384C352 379.2 354.1 374.9 359.5 373.2z']
};
const faTableLayout = {
  prefix: 'far',
  iconName: 'table-layout',
  icon: [512, 512, [], 'e290', 'M448 32C483.3 32 512 60.65 512 96V416C512 451.3 483.3 480 448 480H64C28.65 480 0 451.3 0 416V96C0 60.65 28.65 32 64 32H448zM448 80H64C55.16 80 48 87.16 48 96V160H464V96C464 87.16 456.8 80 448 80zM64 432H144V208H48V416C48 424.8 55.16 432 64 432zM192 432H448C456.8 432 464 424.8 464 416V208H192V432z']
};

/***/ }),

/***/ "./src/taskpane/index.html":
/*!*********************************!*\
  !*** ./src/taskpane/index.html ***!
  \*********************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../node_modules/html-loader/dist/runtime/getUrl.js */ "./node_modules/html-loader/dist/runtime/getUrl.js");
/* harmony import */ var _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0__);
// Imports

var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_1___ = new URL(/* asset import */ __webpack_require__(/*! ./css/bootstrap3.css */ "./src/taskpane/css/bootstrap3.css"), __webpack_require__.b);
// Module
var ___HTML_LOADER_REPLACEMENT_0___ = _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default()(___HTML_LOADER_IMPORT_0___);
var ___HTML_LOADER_REPLACEMENT_1___ = _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default()(___HTML_LOADER_IMPORT_1___);
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\r\n<!-- This file shows how to design a first-run page that provides a welcome screen to the user about the features of the add-in. -->\r\n\r\n<!DOCTYPE html>\r\n<html>\r\n\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\r\n    <title>Link addin</title>\r\n\r\n    <!-- Office JavaScript API -->\r\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\r\n\r\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\r\n    <link rel=\"stylesheet\"\r\n        href=\"https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css\" />\r\n\r\n    <!-- Template styles -->\r\n    <link href=\"" + ___HTML_LOADER_REPLACEMENT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n\r\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\r\n    <link rel=\"stylesheet\"\r\n        href=\"https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css\" />\r\n    <link href=\"https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css\" rel=\"stylesheet\"\r\n        integrity=\"sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65\" crossorigin=\"anonymous\">\r\n    <link rel=\"stylesheet\" href=\"https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css\">\r\n\r\n    <!-- Template styles -->\r\n    <link href=\"" + ___HTML_LOADER_REPLACEMENT_1___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n    <link rel=\"stylesheet\" href=\"https://cdn.jsdelivr.net/gh/mdbassit/Coloris@latest/dist/coloris.min.css\" />\r\n    <" + "script src=\"https://cdn.jsdelivr.net/gh/mdbassit/Coloris@latest/dist/coloris.min.js\"><" + "/script>\r\n</head>\r\n\r\n<body class=\"ms-font-m ms-welcome ms-Fabric\">\r\n    <" + "script src=\"https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/js/bootstrap.bundle.min.js\"\r\n        integrity=\"sha384-kenU1KFdBIe4zVF0s0G1M5b4hcpxyD9F7jL+jjXkk+Q2h455rYXK/7HAuoJl+0I4\"\r\n        crossorigin=\"anonymous\"><" + "/script>\r\n\r\n    <div id=\"header-nav\">\r\n        <div class=\"logo-header d-flex w-100 justify-content-between align-items-center bg-light\" id=\"logo-header\">\r\n        </div>\r\n        <!-- <div class=\"header bg-dark mb-2\" id=\"header\">\r\n        </div> -->\r\n    </div>\r\n    <div class=\"loader-wrapper\" id=\"page-loader\" style=\"display: none;\">\r\n        <div class=\"loader\"></div>\r\n    </div>\r\n    <main id=\"app-body\" class=\"d-block mh-34 \">\r\n        <div id=\"ai-tag-list-container\" class=\"accordion\"></div>\r\n    </main>\r\n\r\n    <section id=\"confirmation-popup\"></section>\r\n    <section id=\"toastr\"></section>\r\n    <div id=\"footer\" class=\"py-2 text-center footer\">\r\n\r\n    </div>\r\n</body>\r\n\r\n</html>";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);

/***/ }),

/***/ "./node_modules/html-loader/dist/runtime/getUrl.js":
/*!*********************************************************!*\
  !*** ./node_modules/html-loader/dist/runtime/getUrl.js ***!
  \*********************************************************/
/***/ (function(module) {



module.exports = function (url, options) {
  if (!options) {
    // eslint-disable-next-line no-param-reassign
    options = {};
  }
  if (!url) {
    return url;
  }

  // eslint-disable-next-line no-underscore-dangle, no-param-reassign
  url = String(url.__esModule ? url.default : url);
  if (options.hash) {
    // eslint-disable-next-line no-param-reassign
    url += options.hash;
  }
  if (options.maybeNeedQuotes && /[\t\n\f\r "'=<>`]/.test(url)) {
    return "\"".concat(url, "\"");
  }
  return url;
};

/***/ }),

/***/ "./src/taskpane/css/bootstrap3.css":
/*!*****************************************!*\
  !*** ./src/taskpane/css/bootstrap3.css ***!
  \*****************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "cca4de5cf945742b4fdc.css";

/***/ }),

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "59ebb5cdad1ba8dd9547.css";

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	!function() {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = function(module) {
/******/ 			var getter = module && module.__esModule ?
/******/ 				function() { return module['default']; } :
/******/ 				function() { return module; };
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	!function() {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = function(exports, definition) {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/publicPath */
/******/ 	!function() {
/******/ 		var scriptUrl;
/******/ 		if (__webpack_require__.g.importScripts) scriptUrl = __webpack_require__.g.location + "";
/******/ 		var document = __webpack_require__.g.document;
/******/ 		if (!scriptUrl && document) {
/******/ 			if (document.currentScript)
/******/ 				scriptUrl = document.currentScript.src;
/******/ 			if (!scriptUrl) {
/******/ 				var scripts = document.getElementsByTagName("script");
/******/ 				if(scripts.length) {
/******/ 					var i = scripts.length - 1;
/******/ 					while (i > -1 && (!scriptUrl || !/^http(s?):/.test(scriptUrl))) scriptUrl = scripts[i--].src;
/******/ 				}
/******/ 			}
/******/ 		}
/******/ 		// When supporting browsers where an automatic publicPath is not supported you must specify an output.publicPath manually via configuration
/******/ 		// or pass an empty string ("") and set the __webpack_public_path__ variable from your code to use your own logic.
/******/ 		if (!scriptUrl) throw new Error("Automatic publicPath is not supported in this browser");
/******/ 		scriptUrl = scriptUrl.replace(/#.*$/, "").replace(/\?.*$/, "").replace(/\/[^\/]+$/, "/");
/******/ 		__webpack_require__.p = scriptUrl;
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	!function() {
/******/ 		__webpack_require__.b = document.baseURI || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"taskpane": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	}();
/******/ 	
/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module is referenced by other modules so it can't be inlined
/******/ 	__webpack_require__("./src/taskpane/taskpane.ts");
/******/ 	var __webpack_exports__ = __webpack_require__("./src/taskpane/index.html");
/******/ 	
/******/ })()
;
//# sourceMappingURL=taskpane.js.map