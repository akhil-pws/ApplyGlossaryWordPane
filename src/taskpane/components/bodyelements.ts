function addtagbody(sponsorOptions) {
  const body = `<div class="modal-dialog">
  <div class="modal-content">
    <div class="modal-body p-3">
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

function logoheader(storedUrl) {
  const body = `
    <img id="main-logo" src="${storedUrl}/assets/logo.png" alt="" class="logo">
    <div class="icon-nav me-3">
    <i class="fa fa-home c-pointer me-3" title="Home" id="home"></i>
<div class="dropdown d-inline">
  <i class="fa fa-tools c-pointer me-3" id="settingsDropdown" data-bs-toggle="dropdown" aria-expanded="false" title="Settings"></i>
  <ul class="dropdown-menu" aria-labelledby="settingsDropdown">
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

    <i class="fa fa-sign-out c-pointer me-3" id="logout" title="Logout"></i>
    </div>    
`
  return body
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


export { navTabs, addtagbody, promptbuilderbody, logoheader };