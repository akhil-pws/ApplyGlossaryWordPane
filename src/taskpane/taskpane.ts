// Imports
import { CONFIG } from "./utils/config";
import { AuthService } from "./services/auth.service";
import { DocumentService } from "./services/document.service";
import { UIService } from "./services/ui.service";
import { StoreService } from "./services/store.service";
import { AIService } from "./services/ai.service";
// Note: GlossaryService import removed if unused or moved

// Restoration of variables needed by the rest of the file (Legacy Support - check if needed)
import { generateCheckboxHistory, getDateTimeStamp, initializeAIHistoryEvents, loadHomepage, replaceMention, setupPromptBuilderUI } from "./draft/home";
import { chatfooter, colorTable, insertLineWithHeadingStyle, mapImagesToComponentObjects, resolveWordTableStyle, selectMatchingBookmarkFromSelection, svgBase64ToPngBase64, switchModeIcon, switchToAddTag, switchToPromptBuilder, updateEditorFinalTable } from "./draft/draft-functions";
import { addtagbody, customizeTablePopup, logoheader, navTabs, toaster } from "./components/bodyelements";
import { addAiHistory, addGroupKey, fetchGlossaryTemplate, getAiHistory, getAllClients, getAllCustomTables, getAllPromptTemplates, getGeneralImages, getReportById, getReportHeadImageById, loginUser, updateGroupKey } from "./draft/draft.api";
import { wordTableStyles } from "./components/tablestyles";
import { renderSelectedTags } from "./draft/draft-functions";
import { loadSummarypage } from "./summary/summary";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("footer").innerText = `© ${new Date().getFullYear()} - TrialAssure LINK AI Assistant ${CONFIG.version}`;

    // Initialize Services
    // AuthService.init(); // if needed

    // Retrieve Properties via Service
    DocumentService.retrieveDocumentProperties().then((props) => {
      if (props) {
        // Update local state for legacy compatibility
        // documentID = props.documentID; // Moved to Store
        // organizationName = props.organizationName; // Moved to Store
        const store = StoreService.getInstance();
        store.documentID = props.documentID;
        store.organizationName = props.organizationName;

        // Check Session
        const session = AuthService.restoreSession();
        if (session) {
          // Restore session state
          store.jwt = session.jwt;
          store.UserRole = session.userRole;
          if (session.tableStyle) store.tableStyle = session.tableStyle;
          if (session.colorPallete) store.colorPallete = session.colorPallete;

          window.location.hash = '#/dashboard';
          toaster('You are successfully logged in', 'success');
          displayMenu(); // Trigger legacy menu display
        } else {
          loadLoginPage();
        }

      } else {
        document.getElementById('app-body').innerHTML = `
        <p class="px-3 text-center">Export a document from the LINK AI application to use this functionality.</p>`
        console.log(`Custom property "documentID" not found.`);
      }
    }).catch(err => {
      console.error("Failed to initialize", err);
    });

    // Restore theme preference from sessionStorage
    const store = StoreService.getInstance();
    const savedTheme = sessionStorage.getItem('theme');
    if (savedTheme) {
      store.theme = savedTheme;
      UIService.applyTheme(savedTheme as 'Light' | 'Dark');
    }

    // Setup UI
    setupEventHandlers();
  }
});

function setupEventHandlers() {
  // Porting event listeners
  document.getElementById("login-btn")?.addEventListener("click", handleLogin);
}

async function login() {
  // document.getElementById('header').innerHTML = ``
  const sessionToken = sessionStorage.getItem('token');
  const store = StoreService.getInstance();
  if (sessionToken) {
    store.UserRole = JSON.parse(sessionStorage.getItem('userRole')) || ''
    store.jwt = sessionToken;
    window.location.hash = '#/dashboard';
    const style = sessionStorage.getItem('tableStyle');
    if (style) {
      store.tableStyle = style;
    }
    const localPallete = sessionStorage.getItem('colorPallete');
    if (localPallete) {
      store.colorPallete = JSON.parse(localPallete);
    }

  } else {
    loadLoginPage();
  }
}

function loadLoginPage() {
  const store = StoreService.getInstance();
  UIService.renderLoginPage(CONFIG.storeUrl, handleLogin, () => {
    store.theme = store.theme === 'Light' ? 'Dark' : 'Light';
    UIService.applyTheme(store.theme as 'Light' | 'Dark');
    sessionStorage.setItem('theme', store.theme);
  });
}

async function handleLogin(event) {
  event.preventDefault();

  const organization = (document.getElementById('organization') as HTMLInputElement).value;
  const username = (document.getElementById('username') as HTMLInputElement).value;
  const password = (document.getElementById('password') as HTMLInputElement).value;

  const store = StoreService.getInstance();

  if (organization.toLowerCase().trim() === store.organizationName.toLocaleLowerCase().trim()) {
    UIService.toggleLoader(true);

    // Use AuthService
    const result = await AuthService.login(organization, username, password);

    // Hide loader is handled by UI replacement or overwrite below? 
    // Legacy code overwrote app-body with loader, so we need to be careful.
    // Actually, legacy code replaced innerHTML with loader. calling displayMenu() replaces it again.

    if (result.success) {
      const data = result.data;
      const store = StoreService.getInstance();
      store.jwt = data.token;
      store.UserRole = data.userRole;

      // Preserve legacy logic for style restoring if it was there
      const style = sessionStorage.getItem('tableStyle');
      if (style) store.tableStyle = style;

      const localPallete = sessionStorage.getItem('colorPallete');
      if (localPallete) store.colorPallete = JSON.parse(localPallete);

      toaster('You are successfully logged in', 'success');
      displayMenu();
      window.location.hash = '#/dashboard';
    } else {
      showLoginError(result.message || "Login failed");
    }
  } else {
    showLoginError("The organization specified is not associated with this document")
  }
}

function showLoginError(message) {
  loadLoginPage();  // Reload the form UI
  const errorDiv = document.getElementById('login-error');
  errorDiv.style.display = 'block';
  errorDiv.textContent = message;
}

function displayMenu() {
  const store = StoreService.getInstance();
  store.userId = Number(sessionStorage.getItem('userId'))
  // document.getElementById('aitag').addEventListener('click', redirectAI);
  fetchDocument('Init');

}

async function getTableStyle() {
  const store = StoreService.getInstance();
  const tableStyleObj = await getAllCustomTables(store.jwt);
  store.customTableStyle = tableStyleObj['Data'];
  const selectedTable = store.customTableStyle.find(style => style.ID === store.dataList.TableCustomizationID);
  if (selectedTable) {
    sessionStorage.setItem("CustomStyle", selectedTable ? selectedTable.Name : '');
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

async function fetchDocument(action) {
  UIService.toggleLoader(true);
  try {
    const store = StoreService.getInstance();
    const userId = sessionStorage.getItem('userId') || '0';
    const reportData = await DocumentService.loadReportData(store.documentID, store.jwt, userId);

    // Assign to store
    store.dataList = reportData.dataList;
    await getTableStyle();
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
      if (store.mode === "Home") loadHomepage(store.availableKeys);
      if (store.mode === "Summary") loadSummarypage(store.availableKeys);
    }

    // Render navigation header
    const logoHeaderEl = document.getElementById('logo-header');
    if (logoHeaderEl) {
      logoHeaderEl.innerHTML = logoheader(CONFIG.storeUrl);
    }

    switchModeIcon();
    UIService.toggleLoader(false);

    // Fetch images in background
    getImages();

    // Event Wiring
    UIService.attachDashboardEvents({
      onHome: async () => {
        if (!store.isPendingResponse) {
          if (store.isGlossaryActive) await removeMatchingContentControls();
          loadHomepage(store.availableKeys);
        }
        store.mode = 'Home';
        switchModeIcon();
      },
      onSummary: async () => {
        if (!store.isPendingResponse) {
          if (store.isGlossaryActive) await removeMatchingContentControls();
          loadSummarypage(store.availableKeys);
        }
        store.mode = 'Summary';
        switchModeIcon();
      },
      onGlossary: () => {
        if (store.emptyFormat) fetchGlossary();
      },
      onFormat: () => {
        if (!store.isPendingResponse) formatOptionsDisplay();
      },
      onRemoveFormat: () => {
        if (Object.keys(store.capturedFormatting).length > 0) removeOptionsConfirmation();
      },
      onThemeToggle: () => {
        store.theme = store.theme === 'Light' ? 'Dark' : 'Light';
        UIService.applyTheme(store.theme as 'Light' | 'Dark');
        sessionStorage.setItem('theme', store.theme);
      },
      onLogout: async () => {
        if (!store.isPendingResponse) {
          if (store.isGlossaryActive) await removeMatchingContentControls();
          logout();
        }
      }
    });

    // Register selection change handler for tag detection
    if (action === 'Init') {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        handleSelectionChange
      );
    }

    UIService.toggleLoader(false);
  } catch (error) {
    console.error("Error loading document data", error);
    UIService.showNotification("Error loading data", "error");
    UIService.toggleLoader(false);
  }
}

export async function formatOptionsDisplay() {
  const store = StoreService.getInstance();
  if (!store.isTagUpdating) { // Check if isTagUpdating is false
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

    const glossaryBtn = document.getElementById('glossary') as HTMLButtonElement;
    if (!glossaryBtn.classList.contains('disabled-link')) {
      glossaryBtn.classList.add('disabled-link');
    }

    if (store.emptyFormat) {
      clearCapturedFormatting();
    }
    else {
      if (store.capturedFormatting.Bold === null || store.capturedFormatting.Bold === undefined ||
        store.capturedFormatting.Underline === 'Mixed' || store.capturedFormatting.Underline === undefined ||
        store.capturedFormatting.Size === null || store.capturedFormatting.Size === undefined ||
        store.capturedFormatting["Font Name"] === null || store.capturedFormatting["Font Name"] === undefined ||
        store.capturedFormatting["Background Color"] === '' || store.capturedFormatting["Background Color"] === undefined ||
        store.capturedFormatting["Text Color"] === '' || store.capturedFormatting["Text Color"] === undefined) {
        const formatList = document.getElementById("format-list");
        formatList.innerHTML = "<p>Multiple style values found. Try again</p>";
        const removeFormatBtn = document.getElementById('removeFormatting') as HTMLButtonElement;


        if (!removeFormatBtn.classList.contains('disabled-link')) {
          removeFormatBtn.classList.add('disabled-link');
        }
      } else {
        const removeFormatBtn = document.getElementById('removeFormatting') as HTMLButtonElement;
        removeFormatBtn.classList.remove('disabled-link');
        displayCapturedFormatting();
      }
    }
    // Event listeners for the buttons

    document.getElementById("capture-format-btn").addEventListener("click", captureFormatting);

    const emptyFormatCheckbox = document.getElementById("empty-format-checkbox") as HTMLInputElement;
    if (store.isNoFormatTextAvailable) {
      emptyFormatCheckbox.checked = true;
      clearCapturedFormatting();
    }

    emptyFormatCheckbox.addEventListener("change", () => {
      if (emptyFormatCheckbox.checked) {
        store.isNoFormatTextAvailable = true;
        clearCapturedFormatting();
      } else {
        const CaptureBtn = document.getElementById('capture-format-btn') as HTMLButtonElement;
        CaptureBtn.disabled = false;
        store.isNoFormatTextAvailable = false;
        store.emptyFormat = false;
        const glossaryBtn = document.getElementById('glossary') as HTMLButtonElement;
        if (!glossaryBtn.classList.contains('disabled-link')) {
          glossaryBtn.classList.add('disabled-link');
        }
      }
    });

  }
}



function displayCapturedFormatting() {
  const store = StoreService.getInstance();
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
  const store = StoreService.getInstance();
  store.capturedFormatting = {}; // Clear the captured formatting object
  const formatDetails = document.getElementById("format-details");
  formatDetails.style.display = 'none';
  // formatList.innerHTML = `<li>No formatting selected.</li>`;
  store.emptyFormat = true;
  const glossaryBtn = document.getElementById('glossary') as HTMLButtonElement;
  glossaryBtn.classList.remove('disabled-link');
  const CaptureBtn = document.getElementById('capture-format-btn') as HTMLButtonElement;
  CaptureBtn.disabled = true;


  const removeFormatBtn = document.getElementById('removeFormatting') as HTMLButtonElement;
  if (!removeFormatBtn.classList.contains('disabled-link')) {
    removeFormatBtn.classList.add('disabled-link');
  }
  console.log("Captured formatting cleared.");
}

async function captureFormatting() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const font = selection.font;
      font.load(["bold", "italic", "underline", "size", "highlightColor", "name", 'color']);

      await context.sync();

      const store = StoreService.getInstance();

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

      if (store.capturedFormatting.Bold === null ||
        store.capturedFormatting.Underline === 'Mixed' ||
        store.capturedFormatting.Size === null ||
        store.capturedFormatting["Font Name"] === null ||
        store.capturedFormatting["Background Color"] === '' ||
        store.capturedFormatting["Text Color"] === ''

      ) {
        const formatList = document.getElementById("format-list");
        formatList.innerHTML = "<p>Multiple style values found. Try again</p>";
        const removeFormatBtn = document.getElementById('removeFormatting') as HTMLButtonElement;
        if (!removeFormatBtn.classList.contains('disabled-link')) {
          removeFormatBtn.classList.add('disabled-link');
        }

      } else {
        const removeFormatBtn = document.getElementById('removeFormatting') as HTMLButtonElement;
        removeFormatBtn.classList.remove('disabled-link');
        displayCapturedFormatting();
      }
    });
  } catch (error) {
    console.error("Error capturing formatting:", error);
  }
}



async function removeOptionsConfirmation() {
  const store = StoreService.getInstance();
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

    if (store.capturedFormatting['Background Color'] === null &&
      store.capturedFormatting['Text Color'] === '#000000') {
      const warningEle = document.getElementById('warning-rem-fmt').innerHTML = 'Warning : The captured formatting is broad. This might result in unintended text removal throughout the document. Proceed?'
    }

    // Event listeners for the buttons
    document.getElementById("clear-ft-btn").addEventListener("click", removeFormattedText);
    document.getElementById("change-ft-btn").addEventListener("click", formatOptionsDisplay);

  }
}

async function removeFormattedText() {
  try {
    await Word.run(async (context) => {

      const iconelement = document.getElementById(`clear-ft-btn`);
      iconelement.innerHTML = `<i class="fa fa-spinner fa-spin text-white me-2"></i>Yes`;
      const clrBtn = document.getElementById('clear-ft-btn') as HTMLButtonElement;
      clrBtn.disabled = true;

      const changeBtn = document.getElementById('change-ft-btn') as HTMLButtonElement;
      changeBtn.disabled = true;
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items"); // Load paragraphs from the body

      await context.sync();

      const store = StoreService.getInstance();

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
            if (
              font.highlightColor === store.capturedFormatting['Background Color'] &&
              font.color === store.capturedFormatting['Text Color'] &&
              font.bold === store.capturedFormatting['Bold'] &&
              font.italic === store.capturedFormatting['Italic'] &&
              font.size === store.capturedFormatting['Size'] &&
              font.underline === store.capturedFormatting['Underline'] &&
              font.name === store.capturedFormatting['Font Name']
            ) {
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
      const glossaryBtn = document.getElementById('glossary') as HTMLButtonElement;
      glossaryBtn.classList.remove('disabled-link');
      formatOptionsDisplay()
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
  const store = StoreService.getInstance();
  if (store.isGlossaryActive) {
    await removeMatchingContentControls();
  }
  sessionStorage.clear();
  window.location.hash = '#/new';
  store.initialised = true;
  document.getElementById('logo-header').innerHTML = ``;
  login();
}

export async function applyTagFn() {

  return Word.run(async (context) => {
    try {
      const body = context.document.body;

      context.load(body, 'text');
      await context.sync();
      await applyAITagFn(body, context);
      await applyImageTagFn(body, context);
    } catch (err) {
      toaster("Something went wrong", "error")
      console.error("Error during tag application:", err);
      const store = StoreService.getInstance();
      loadHomepage(store.availableKeys);
    }
  });
}

async function applyImageTagFn(body: Word.Body, context: Word.RequestContext) {
  const store = StoreService.getInstance();
  for (let i = 0; i < store.imageList.length; i++) {
    const tag = store.imageList[i];
    const searchResults = body.search(`$${tag.DisplayName}$`, {
      matchCase: false,
      matchWholeWord: false,
    });
    context.load(searchResults, 'items');
    await context.sync();

    for (const item of searchResults.items) {
      if (tag.EditorValue !== "") {
        let base64Image: string = tag.EditorValue;

        // Clean base64
        if (!base64Image) continue;

        // Convert SVG → PNG
        if (base64Image.startsWith("data:image/svg+xml")) {
          base64Image = await svgBase64ToPngBase64(base64Image);
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
  toaster("AI tag application completed!", "success");
  loadHomepage(store.availableKeys);
}

export async function applyAITagFn(
  body: Word.Body,
  context: Word.RequestContext
) {
  document.getElementById('app-body').innerHTML = `
  <div id="button-container">
    <div class="loader" id="loader"></div>
    <div id="highlighted-text"></div>
  </div>`
  toaster("Please wait... applying AI tags", "info");

  const store = StoreService.getInstance();
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

      let bookmarkStart: Word.Range | null = null;
      let bookmarkEnd: Word.Range | null = null;

      const include = (r: Word.Range) => {
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
            for (const line of txt.split("\n")) {
              if (!line.trim()) continue;

              const p = cursor.insertParagraph("", Word.InsertLocation.after);
              insertLineWithHeadingStyle(p, line);

              include(p.getRange());
              cursor = p.getRange();
            }
          }

          // ELEMENT NODE
          else if (node.nodeType === Node.ELEMENT_NODE) {
            const el = node as HTMLElement;

            // TABLE
            if (el.tagName.toLowerCase() === "table") {
              const rows = Array.from(el.querySelectorAll("tr"));
              if (!rows.length) continue;

              const maxCols = Math.max(
                ...rows.map(r =>
                  Array.from(r.querySelectorAll("td, th")).reduce(
                    (s, c) =>
                      s + parseInt(c.getAttribute("colspan") || "1", 10),
                    0
                  )
                )
              );

              const p = cursor.insertParagraph("", Word.InsertLocation.after);
              const table = p.insertTable(
                rows.length,
                maxCols,
                Word.InsertLocation.after
              );

              const store = StoreService.getInstance();
              const resolvedTableStyle = resolveWordTableStyle(store.tableStyle);
              if (resolvedTableStyle !== 'none') {
                table.style = resolvedTableStyle;
              }

              if (store.colorPallete.Customize) {
                await colorTable(table, rows, context);
              } else {
                const rowspanTracker: number[] = new Array(maxCols).fill(0);
                let lastParamRowIndex = -1; // row index of last non-empty Parameter cell

                rows.forEach((row, rowIndex) => {
                  const cells = Array.from(row.querySelectorAll('td, th'));
                  let cellIndex = 0;

                  // track first column text in HTML row
                  let firstColText = "";

                  cells.forEach((cell) => {
                    while (rowspanTracker[cellIndex] > 0) {
                      rowspanTracker[cellIndex]--;
                      cellIndex++;
                    }

                    const cellText = Array.from(cell.childNodes)
                      .map(node => {
                        if (node.nodeType === Node.TEXT_NODE) {
                          return node.textContent?.trim() || '';
                        } else if (node.nodeType === Node.ELEMENT_NODE) {
                          return (node as HTMLElement).innerText.trim();
                        }
                        return '';
                      })
                      .filter(text => text.length > 0)
                      .join(' ');

                    const colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
                    const rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10);

                    // save first column text if this is first column
                    if (cellIndex === 0) {
                      firstColText = cellText.trim();
                    }

                    table.getCell(rowIndex, cellIndex).value = cellText;

                    // colspan blank filling
                    for (let i = 1; i < colspan; i++) {
                      if (cellIndex + i < maxCols) {
                        table.getCell(rowIndex, cellIndex + i).value = "";
                      }
                    }

                    // rowspan tracking
                    if (rowspan > 1) {
                      for (let i = 0; i < colspan; i++) {
                        if (cellIndex + i < maxCols) {
                          rowspanTracker[cellIndex + i] = rowspan - 1;
                        }
                      }
                    }

                    cellIndex += colspan;
                  });

                  // ✅ AFTER filling the row → apply merge logic for 1st column
                  if (rowIndex === 0) {
                    // header row typically
                    return;
                  }

                  if (firstColText) {
                    // new parameter starts
                    lastParamRowIndex = rowIndex;
                  } else {
                    // empty parameter row → merge vertically with previous parameter row
                    if (lastParamRowIndex !== -1) {
                      const topCell = table.getCell(lastParamRowIndex, 0);
                      const bottomCell = table.getCell(rowIndex, 0);

                      // merge bottom into top
                      topCell.merge(bottomCell);

                      // ✅ center align merged cell
                      try {
                        topCell.verticalAlignment = Word.VerticalAlignment.center;
                        topCell.body.paragraphs.getFirst().alignment = Word.Alignment.center;
                      } catch (e) {
                        // ignore (safe fallback)
                      }
                    }
                  }
                });

              }

              include(table.getRange());
              cursor = table.getRange();
            }

            // OTHER ELEMENTS
            else {
              let txt = el.innerText?.trim();
              if (!txt) continue;

              txt = txt.replace(/\n- /g, "\n• ");
              for (const line of txt.split("\n")) {
                if (!line.trim()) continue;

                const p = cursor.insertParagraph("", Word.InsertLocation.after);
                insertLineWithHeadingStyle(p, line);

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
          base64 = await svgBase64ToPngBase64(base64);
        } else if (base64.startsWith("data:image")) {
          base64 = base64.split(",")[1];
        }

        const pic = cursor.insertInlinePictureFromBase64(
          base64,
          Word.InsertLocation.after
        );

        include(pic.getRange());
        cursor = pic.getRange();
      }

      // TEXT CONTENT
      else {
        const txt = tag.EditorValue
          .replace(/\n- /g, "\n• ")
          .trim();

        for (const line of txt.split("\n")) {
          if (!line.trim()) continue;

          const p = cursor.insertParagraph("", Word.InsertLocation.after);
          insertLineWithHeadingStyle(p, line);

          include(p.getRange());
          cursor = p.getRange();
        }
      }

      await context.sync();

      /* --------------------------------------------------
         3️⃣ Create SINGLE bookmark
      -------------------------------------------------- */
      if (bookmarkStart && bookmarkEnd) {
        const bookmarkName = `ID${tag.ID}_Split_${getDateTimeStamp()}`;
        bookmarkStart.expandTo(bookmarkEnd).insertBookmark(bookmarkName);
      }
    }
  }
}

export async function normalizeBlankLines(
  context: Word.RequestContext
) {
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

export async function removeTrailingEmptyParagraphs(
  context: Word.RequestContext
) {
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
  const store = StoreService.getInstance();
  if (!store.isTagUpdating) {

    document.getElementById('app-body').innerHTML = `
  <div id="button-container">

          <div class="loader" id="loader"></div>

        <div id="highlighted-text"></div>`

    loadGlossary()

  }

}


function loadGlossary() {
  document.getElementById('app-body').innerHTML = `
        <div id="button-container">
          <button class="btn btn-secondary me-2 mark-glossary btn-sm" id="applyglossary">Apply Glossary</button>
        </div>
  `
  document.getElementById('applyglossary').addEventListener('click', applyglossary);


}



export async function applyglossary() {
  document.getElementById('app-body').innerHTML = `
  <div id="button-container">

          <div class="loader" id="loader"></div>

        <div id="highlighted-text"></div>`

  try {

    await Word.run(async (context) => {


      const body = context.document.body;
      body.load("text");
      await context.sync(); // Sync to get the text content

      const bodyText = {
        "Content": body.text.replace(/[\n\r]/g, ' ')
      };
      try {
        const store = StoreService.getInstance();
        const data = await fetchGlossaryTemplate(store.dataList?.ClientID, bodyText, store.jwt);

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
      const store = StoreService.getInstance();
      store.layTerms.sort((a, b) => b.ClinicalTerm.length - a.ClinicalTerm.length);

      const processedTerms = new Set(); // Track added larger terms

      // Filter out smaller terms if they are included in a larger term
      const filteredTerms = store.layTerms.filter(term => {
        for (const biggerTerm of processedTerms) {
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

      const searchPromises = store.filteredGlossaryTerm.map(term => {
        const searchResults = body.search(term.ClinicalTerm, { matchCase: false, matchWholeWord: false });
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
      
`
      const displayElement = document.getElementById('loader');
      displayElement.style.display = 'none';
      await context.sync();
      document.getElementById('clearGlossary').addEventListener('click', removeMatchingContentControls);
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        handleSelectionChange
      );


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
  const store = StoreService.getInstance();

  // Handle glossary mode
  if (store.isGlossaryActive) {
    await checkGlossary();
  }

  // Handle Home mode - detect bookmarks/tags in selection
  if (store.mode === 'Home') {
    await logBookmarksInSelection();
  }
}

export async function checkGlossary() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      selection.load("text, font.highlightColor");

      await context.sync();



      if (selection.text) {
        const loader = document.getElementById('loader');
        if (loader) {
          loader.style.display = 'block';
        }
        const store = StoreService.getInstance();
        const searchPromises = store.layTerms.map(term => {
          const searchResults = selection.search(term.ClinicalTerm, { matchCase: false, matchWholeWord: false });
          searchResults.load("items");
          return searchResults;
        });

        await context.sync();
        const selectedWords = []
        for (const searchResults of searchPromises) {

          for (const range of searchResults.items) {
            const font = range.font;
            font.load(["bold", "italic", "underline", "size", "highlightColor", "name", "color"]);

            await context.sync();
            if (
              font.highlightColor !== store.capturedFormatting['Background Color'] ||
              font.color !== store.capturedFormatting['Text Color'] ||
              font.bold !== store.capturedFormatting['Bold'] ||
              font.italic !== store.capturedFormatting['Italic'] ||
              font.size !== store.capturedFormatting['Size'] ||
              font.underline !== store.capturedFormatting['Underline'] ||
              font.name !== store.capturedFormatting['Font Name']
            ) {
              selectedWords.push(range.text);
            }

          }
        }
        // searchPromises.forEach(searchResults => {
        //   searchResults.items.forEach(item => {
        //   });
        // });
        displayHighlightedText(selectedWords)

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



function displayHighlightedText(words: string[]) {

  const displayElement = document.getElementById('highlighted-text');

  if (displayElement) {
    displayElement.innerHTML = ''; // Clear previous content
    const loader = document.getElementById('loader');
    loader.style.display = 'block';
    // Group lay terms by their clinical term
    const groupedTerms: { [clinicalTerm: string]: string[] } = {};

    const store = StoreService.getInstance();
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
async function replaceClinicalTerm(clinicalTerm: string, layTerm: string) {
  const displayElement = document.getElementById('loader');
  displayElement.style.display = 'block';

  try {
    await Word.run(async (context) => {
      // Get the current selection
      const selection = context.document.getSelection();
      selection.load('text');
      await context.sync();

      if (selection.text.toLowerCase().includes(clinicalTerm.toLowerCase())) {
        // Search for the clinicalTerm in the document
        const searchResults = selection.search(clinicalTerm, { matchCase: false, matchWholeWord: false });
        searchResults.load('items');

        await context.sync();

        // Replace each occurrence of the clinicalTerm with the layTerm
        for (const item of searchResults.items) {
          // Load the font properties
          item.font.load(['bold', 'italic', 'underline', 'color', 'highlightColor', 'size', 'name']);
          await context.sync();  // Ensure the properties are loaded before accessing them

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


export async function removeMatchingContentControls() {
  try {
    await Word.run(async (context) => {
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
        const store = StoreService.getInstance();
        if (control.title && store.filteredGlossaryTerm.some(term => term.ClinicalTerm.toLowerCase() === control.title.toLowerCase())) {
          const range = control.getRange();
          range.load("text");
          await context.sync();
          if (control.tag && /^#[0-9A-Fa-f]{6}$/.test(control.tag)) {
            range.font.highlightColor = control.tag;
          } else {
            range.font.highlightColor = null
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
      const store = StoreService.getInstance();
      store.isGlossaryActive = false;
      document.getElementById('applyglossary').addEventListener('click', applyglossary);
    });
  } catch (error) {
    console.error("Error removing content controls:", error);
  }
}

export async function addGenAITags() {
  const store = StoreService.getInstance();
  if (!store.isTagUpdating) {

    if (store.isGlossaryActive) {
      await removeMatchingContentControls();
    }

    let selectedClient = store.clientList.filter(item => item.ID === store.clientId);

    // Build Primary Source List
    let sourceTypeList = [
      ...Array.from(new Map(
        store.dataList.SourceTypeList
          .filter(item => item.VectorID > 0)
          .map(item => [item.SourceTypeID, { Name: item.SourceType, ID: item.SourceTypeID }])
      ).values())
    ];


    let sourceOptions = sourceTypeList.map((src: any) => {
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

    document.getElementById('app-body').innerHTML = navTabs;

    // Inject modal
    document.getElementById('add-tag-body').innerHTML = addtagbody(sponsorOptions, sourceOptions, store.mode === "Summary");

    const promptTemplateElement = document.getElementById('add-prompt-template');
    setupPromptBuilderUI(promptTemplateElement, store.promptBuilderList);

    document.getElementById('tag-tab').addEventListener('click', () => switchToAddTag());
    document.getElementById('prompt-tab').addEventListener('click', () => switchToPromptBuilder());

    mentionDropdownFn('prompt', 'mention-dropdown', 'add');

    const form = document.getElementById('genai-form');
    const nameField = document.getElementById('name');
    const descriptionField = document.getElementById('description') as HTMLInputElement;
    const promptField = document.getElementById('prompt') as HTMLTextAreaElement;
    // const primarySourceField = document.getElementById('primarySource');

    const saveGloballyCheckbox = document.getElementById('saveGlobally') as HTMLInputElement;
    const availableForAllCheckbox = document.getElementById('isAvailableForAll') as HTMLInputElement;
    const sponsorDropdownButton = document.getElementById('sponsorDropdown');
    const sponsorDropdownItems = document.querySelectorAll('.sponsor-dropdown-item .form-check-input');


    const sourceDropdownButton = document.getElementById('sourceDropdown');
    const sourceDropdownItems = document.querySelectorAll('.source-dropdown-item .form-check-input');

    const isSummaryMode = store.mode === "Summary";

    document.getElementById('cancel-btn-gen-ai').addEventListener('click', () => {
      const store = StoreService.getInstance();
      if (!store.isPendingResponse) loadHomepage(store.availableKeys);
    });


    if (form && nameField && promptField && sponsorDropdownItems.length > 0 && (isSummaryMode || sourceDropdownItems.length > 0)) {

      const updateSponsorDropdownLabel = () => {
        if ((availableForAllCheckbox as HTMLInputElement).checked) {
          sponsorDropdownButton.textContent = store.clientList.map(x => x.Name).join(", ");
        } else {
          const selectedNames = Array.from(sponsorDropdownItems)
            .filter(cb => (cb as HTMLInputElement).checked && cb.id !== 'sponsorSelectAll')
            .map(cb => cb.parentElement.textContent.trim());

          sponsorDropdownButton.textContent = selectedNames.length
            ? selectedNames.join(", ")
            : "Select Sponsors";
        }
      };

      const updateSourceDropdownLabel = () => {
        const selectedNames = Array.from(sourceDropdownItems)
          .filter(cb => (cb as HTMLInputElement).checked && cb.id !== 'sourceSelectAll')
          .map(cb => cb.parentElement.textContent.trim());

        sourceDropdownButton.textContent = selectedNames.length
          ? selectedNames.join(", ")
          : "Select Source Types";
      }

      // Submit Handler
      form.addEventListener('submit', async (e) => {
        e.preventDefault();

        form.querySelectorAll('.is-invalid').forEach(i => i.classList.remove('is-invalid'));

        let valid = true;

        if (!nameField.value.trim()) { nameField.classList.add('is-invalid'); valid = false; }
        if (!promptField.value.trim()) { promptField.classList.add('is-invalid'); valid = false; }

        // PRIMARY SOURCE VALIDATION
        let selectedPrimarySources = [];
        if (!isSummaryMode) {
          selectedPrimarySources = Array.from(sourceDropdownItems)
            .filter(cb => (cb as HTMLInputElement).checked && cb.id !== 'sourceSelectAll')
            .map(cb => (cb as HTMLInputElement).value);

          if (!selectedPrimarySources.length) {
            document.getElementById("primarySourceError").style.display = "block";
            valid = false;
          } else {
            document.getElementById("primarySourceError").style.display = "none";
          }
        }

        if (!valid) return;

        const selectedSponsors = Array.from(sponsorDropdownItems)
          .filter(cb => (cb as HTMLInputElement).checked && cb.id !== 'sponsorSelectAll')
          .map(cb => store.clientList.find(c => c.ID == (cb as HTMLInputElement).value));

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
          SourceTypeID: isSummaryMode ? "0" : selectedPrimarySources.join(","),

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
        sponsorDropdownItems.forEach(cb => {
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
            (sponsorDropdownButton as HTMLButtonElement).disabled = true;

            sponsorDropdownItems.forEach(cb => {
              if (!(cb as HTMLInputElement).disabled) {
                (cb as HTMLInputElement).checked = false;
                (cb as HTMLInputElement).disabled = false;
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
          const checkbox = this.querySelector('.sponsor-dropdown-item .form-check-input');
          if (!checkbox) return;

          if (checkbox.id === 'sponsorSelectAll') {
            const isChecked = (checkbox as HTMLInputElement).checked;
            sponsorDropdownItems.forEach(cb => {
              if (!(cb as HTMLInputElement).disabled) (cb as HTMLInputElement).checked = isChecked;
            });
          }

          updateSponsorDropdownLabel();
        });
      });

      if (!isSummaryMode) {
        document.querySelectorAll('.source-dropdown-item').forEach(item => {
          item.addEventListener('click', function (e) {
            e.stopPropagation();
            const checkbox = this.querySelector('.source-dropdown-item .form-check-input');
            if (!checkbox) return;

            if (checkbox.id === 'sourceSelectAll') {
              const isChecked = (checkbox as HTMLInputElement).checked;
              sourceDropdownItems.forEach(cb => {
                (cb as HTMLInputElement).checked = isChecked;
              });
            }

            updateSourceDropdownLabel();
            const selectedCount = Array.from(sourceDropdownItems)
              .filter(cb => (cb as HTMLInputElement).checked).length;

            if (selectedCount === 0) {
              document.getElementById("primarySourceError").style.display = "block";
            } else {
              document.getElementById("primarySourceError").style.display = "none";
            }

          });
        });
      }

      updateSponsorDropdownLabel();
      if (!isSummaryMode) updateSourceDropdownLabel();

      [nameField, promptField].forEach(field => {
        field.addEventListener('input', function () {
          const input = this as HTMLInputElement;
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


export async function customizeTable(type: string) {
  const store = StoreService.getInstance();
  const container = document.getElementById("confirmation-popup");
  if (!container) return;

  const customStyleName = sessionStorage.getItem("CustomStyle") || "";
  const defaultStyle = sessionStorage.getItem("DefaultStyle") || store.tableStyle;
  let styleObj: any = type === "Custom" ? customStyleName : defaultStyle;
  container.innerHTML = customizeTablePopup(styleObj, type);

  const cancelBtn = document.getElementById("confirmation-popup-cancel");
  const okBtn = document.getElementById("confirmation-popup-confirm");
  const dropdown = document.getElementById("confirmation-popup-dropdown") as HTMLSelectElement;
  const tablePreview = document.getElementById("confirmation-popup-table-preview") as HTMLTableElement;

  const applyStyle = () => {
    if (!dropdown || !tablePreview) return;
    let styleObj: any;
    if (type === "Custom") {
      styleObj = store.customTableStyle.find(s => s.Name === dropdown.value);
    } else {
      styleObj = wordTableStyles.find(s => s.style === dropdown.value);
    }

    if (styleObj && type === 'Pre') {
      // Clear existing styles
      Array.from(tablePreview.rows).forEach(row => {
        Array.from(row.cells).forEach(cell => (cell as HTMLTableCellElement).removeAttribute("style"));
      });

      if (styleObj.tableClass) tablePreview.style.cssText = styleObj.tableClass;

      if (styleObj.headerClass) {
        const thead = tablePreview.querySelector("thead");
        if (thead) {
          Array.from(thead.rows).forEach(row => {
            Array.from(row.cells).forEach(cell => {
              (cell as HTMLTableCellElement).style.cssText = styleObj.headerClass!;
            });
          });
        }
      }

      if (styleObj.sideHeader && styleObj.rowClass) {
        Array.from(tablePreview.rows).forEach((row, index) => {
          Array.from(row.cells).forEach((cell, cellIndex) => {
            if (cellIndex === 0 && index !== 0) {
              (cell as HTMLTableCellElement).style.cssText = "font-weight:bold;";
            }
          });
        });
      }

      if (styleObj.format === "empty" && styleObj.rowClass) {
        Array.from(tablePreview.rows).forEach((row, index) => {
          if (index % 2 === 1) (row as HTMLTableRowElement).style.cssText = styleObj.rowClass!;
        });
      } else if (styleObj.format === "partial" && styleObj.rowClass) {
        Array.from(tablePreview.rows).forEach((row, index) => {
          Array.from(row.cells).forEach((cell, cellIndex) => {
            if (cellIndex === 0) {
              (cell as HTMLTableCellElement).style.cssText =
                styleObj.tableClass! + "font-weight:bold;";
            } else if (index % 2 === 1) {
              (cell as HTMLTableCellElement).style.cssText = styleObj.rowClass!;
            }
          });
        });
      } else if (styleObj.format === "full") {
        Array.from(tablePreview.rows).forEach((row, index) => {
          Array.from(row.cells).forEach((cell, cellIndex) => {
            const headerClass = index === 0 ? styleObj.headerClass! : "";
            if (cellIndex === 0 && styleObj.sideHeader) {
              (cell as HTMLTableCellElement).style.cssText =
                styleObj.tableClass! + "font-weight:bold;" + headerClass;
            } else {
              (cell as HTMLTableCellElement).style.cssText = styleObj.tableClass! + headerClass;
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

  if (cancelBtn) cancelBtn.addEventListener("click", () => (container.innerHTML = ""));
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
        sessionStorage.setItem("CustomStyle", styleObj.Name);
        store.tableStyle = styleObj.Setting.BaseStyle; // stores full object as 
      } else {
        store.colorPallete.Customize = false;
        store.tableStyle = dropdown.value; // normal style string
        sessionStorage.setItem("DefaultStyle", store.tableStyle);

      }

      sessionStorage.setItem("colorPallete", JSON.stringify(store.colorPallete));
      sessionStorage.setItem("tableStyle", store.tableStyle);

      container.innerHTML = "";
    });
  }
}

import { addSummaryTag } from "./summary/summary.api";

async function createTextGenTag(payload) {
  const store = StoreService.getInstance();
  try {
    const iconelement = document.getElementById(`text-gen-save`);
    const cancelBtnGenAi = document.getElementById('cancel-btn-gen-ai');


    (cancelBtnGenAi as HTMLButtonElement).disabled = true;
    iconelement.innerHTML = `<i class="fa fa-spinner fa-spin text-white me-2"></i>Save`;
    (iconelement as HTMLButtonElement).disabled = true;
    store.isPendingResponse = true;

    let data: any;
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
        SummaryTagClient: payload.GroupKeyClient
      };
      data = await addSummaryTag(summaryPayload, store.jwt);
    } else {
      data = await addGroupKey(payload, store.jwt);
    }

    store.isPendingResponse = false;

    if (data['Status']) {
      if (store.mode === "Summary") {
        loadSummarypage(store.availableKeys);
      } else {
        fetchDocument('AIpanel');
      }
      toaster('Saved successfully', 'success');
    } else {
      (cancelBtnGenAi as HTMLButtonElement).disabled = false;
      (iconelement as HTMLButtonElement).disabled = false;
      iconelement.innerHTML = `<i class="fa fa-check-circle me-2"></i>Save`;
      toaster('Something went wrong', 'error');
      // showAddTagError(data['Data']);
    }

  } catch (error) {
    toaster('Something went wrong', 'error');
    console.error('Error creating text generation tag:', error);
  }
}



export function mentionDropdownFn(textareaId, DropdownId, action) {
  const store = StoreService.getInstance();
  const filterMentions = (query) => {
    // Assuming availableKeys is an array of objects with DisplayName and EditorValue properties
    const filtered = store.availableKeys.filter(item => item.AIFlag === 0).filter(item =>
      item.DisplayName.toLowerCase().includes(query.toLowerCase())
    );
    return filtered;
  };
  let highlightedIndex = -1;

  const promptField = document.getElementById(`${textareaId}`);
  const mentionDropdown = document.getElementById(`${DropdownId}`);
  if (promptField) {

    // Handle input events on prompt field for mentions
    promptField.addEventListener('input', (e) => {
      const cursorPosition = (promptField as HTMLTextAreaElement).selectionStart;
      const textBeforeCursor = (promptField as HTMLTextAreaElement).value.slice(0, cursorPosition);
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
    promptField.addEventListener('keydown', (e) => {
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
          mentionDropdown.style.display = 'none';  // Hide the dropdown after selection
          e.preventDefault();  // Prevent form submission on Enter key
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
          behavior: 'smooth',    // Smooth scroll
          block: 'nearest'      // Scroll only if necessary
        });
      }
    }



    // Handle selecting an item from the dropdown via mouse click
    mentionDropdown.addEventListener('click', (e) => {
      if (e.target && e.target.matches('li')) {
        const editorValue = (e.target as HTMLLIElement).getAttribute('data-editor-value');
        selectMention(editorValue);
        mentionDropdown.style.display = 'none';  // Hide the dropdown after selection
      }
    });

    // Function to insert the selected mention into the prompt field
    const selectMention = (editorValue) => {
      const textarea = document.getElementById(`${textareaId}`) as HTMLTextAreaElement;
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
    document.addEventListener('click', (e) => {
      if (!mentionDropdown.contains(e.target as Node) && e.target !== promptField) {
        mentionDropdown.style.display = 'none';
      }
    });
  }
}

function removeQuotes(value: string): string {
  return value
    ? value
      .replace(/^"|"$/g, '')
      .replace(/\\n/g, '')
      .replace(/\*\*/g, '')
      .replace(/\\r/g, '')
    : '';
}

export function createMultiSelectDropdown(tag, type: "Summary" | "AITag") {
  const store = StoreService.getInstance();
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
  debugger

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
          ${Object.keys(groupedSources)
      .map((group, groupIndex) => {
        const groupItems = groupedSources[group]
          .map(
            (source, index) => `
                  <li class="dropdown-item ps-4 ${itemClass}" style="cursor: pointer;" data-checkbox-id="source-${groupIndex}-${index}">
                    <div class="form-check">
                      <input class="form-check-input source-checkbox" type="checkbox" value="${type === 'Summary' ? source.FileName : source.SourceName}" id="source-${groupIndex}-${index}">
                      <label class="form-check-label w-100 text-prewrap" for="source-${groupIndex}-${index}">${type === 'Summary' ? source.FileName : source.SourceName}</label>
                    </div>
                  </li>
                `
          )
          .join('');

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
      })
      .join('')}
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

  const selectAllCheckbox = document.getElementById(`selectAll`) as HTMLInputElement;
  const groupCheckboxes = document.querySelectorAll(`.group-checkbox`);
  const individualCheckboxes = document.querySelectorAll(`.source-checkbox`);
  const sourceDropdownLabel = document.getElementById(`sourceDropdownLabel`);

  function updateLabel() {
    sourceDropdownLabel.innerText = selectedSources.length > 0 ? selectedSources.join(', ') : ' ';
  }

  // Select All logic
  selectAllCheckbox.addEventListener("change", function () {
    const checked = this.checked;
    groupCheckboxes.forEach(cb => (cb as HTMLInputElement).checked = checked);
    individualCheckboxes.forEach(cb => {
      (cb as HTMLInputElement).checked = checked;
      if (checked && !selectedSources.includes((cb as HTMLInputElement).value)) {
        selectedSources.push((cb as HTMLInputElement).value);
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
        (cb as HTMLInputElement).checked = (this as HTMLInputElement).checked;
        if ((this as HTMLInputElement).checked && !selectedSources.includes((cb as HTMLInputElement).value)) {
          selectedSources.push((cb as HTMLInputElement).value);
        }
        if (!(this as HTMLInputElement).checked) {
          selectedSources = selectedSources.filter(s => s !== (cb as HTMLInputElement).value);
        }
      });

      // Update Select All state
      selectAllCheckbox.checked = Array.from(individualCheckboxes).every(child => (child as HTMLInputElement).checked);
      updateLabel();
    });
  });

  // Individual checkbox logic
  individualCheckboxes.forEach(cb => {
    cb.addEventListener("change", function () {
      if ((cb as HTMLInputElement).checked) {
        if (!selectedSources.includes((cb as HTMLInputElement).value)) selectedSources.push((cb as HTMLInputElement).value);
      } else {
        selectedSources = selectedSources.filter(s => s !== (cb as HTMLInputElement).value);
      }

      // Update parent group checkbox
      const groupIndex = cb.id.split("-")[1];
      const groupItems = document.querySelectorAll(`[data-checkbox-id^="source-${groupIndex}-"] .source-checkbox`);
      const groupCheckbox = document.getElementById(`group-${groupIndex}`) as HTMLInputElement;
      groupCheckbox.checked = Array.from(groupItems).every(child => (child as HTMLInputElement).checked);

      // Update Select All checkbox
      selectAllCheckbox.checked = Array.from(individualCheckboxes).every(child => (child as HTMLInputElement).checked);

      updateLabel();
    });
  });

  // Initialize with pre-selected sources
  if (tag.Sources && tag.Sources.length > 0) {
    individualCheckboxes.forEach(cb => {
      if (tag.Sources.includes((cb as HTMLInputElement).value)) {
        (cb as HTMLInputElement).checked = true;
        selectedSources.push((cb as HTMLInputElement).value);
      }
    });

    // Update group checkboxes
    groupCheckboxes.forEach(groupCb => {
      const groupIndex = groupCb.id.split("-")[1];
      const groupItems = document.querySelectorAll(`[data-checkbox-id^="source-${groupIndex}-"] .source-checkbox`);
      (groupCb as HTMLInputElement).checked = Array.from(groupItems).every(child => (child as HTMLInputElement).checked);
    });

    // Update Select All
    selectAllCheckbox.checked = Array.from(individualCheckboxes).every(child => (child as HTMLInputElement).checked);
    updateLabel();
  }

  // Save
  document.getElementById(`ok-src-btn`).addEventListener("click", function () {
    tag.Sources = [...selectedSources];
    const store = StoreService.getInstance();
    const receivedEntry = sourceList.filter(source => selectedSources.includes(type === 'Summary' ? source.FileName : source.SourceName));
    tag.TempSourceValue = receivedEntry.map((item) => {
      return item.VectorID ? String(item.VectorID) : item.SourceValue;
    });
    if (type === 'Summary') {
      tag.FileName = receivedEntry.map((item) => {
        return item.FileName;
      });
    } else {
      tag.SourceName = receivedEntry.map((item) => {
        return item.SourceName;
      });
    }



    tag.SourceValueID = receivedEntry.map((item) => {
      return String(item.VectorID);
    });

    tag.SourceValue = receivedEntry
      .map(source => source.SourceValue);
    accordionBody.innerHTML = chatfooter(tag);
    initializeAIHistoryEvents(tag, store.jwt, store.availableKeys, type);
  });

  // Cancel
  document.getElementById(`cancel-src-btn`).addEventListener("click", function () {
    accordionBody.innerHTML = chatfooter(tag);
    initializeAIHistoryEvents(tag, store.jwt, store.availableKeys, type);
  });
}



async function loadPromptTemplates() {
  const store = StoreService.getInstance();
  try {
    const data = await getAllPromptTemplates(store.jwt);
    if (data.Status && data.Data) {
      store.promptBuilderList = data.Data;
    }
    // Do something with the data
  } catch (error) {
    console.error('Error fetching prompt templates:', error);
  }
}

async function logBookmarksInSelection() {
  return Word.run(async (context) => {

    const rawBookmarks = await getBookmarksFromSelection(context);
    const bookmarks = pickRelevantBookmarks(rawBookmarks);

    if (bookmarks.length === 0) {
      document.getElementById('tags-in-selected-text')
        ?.classList.replace('d-block', 'd-none');
      return;
    }

    document.getElementById('tags-in-selected-text')
      ?.classList.replace('d-none', 'd-block');

    if (bookmarks.length > 1) {
      const store = StoreService.getInstance();
      store.selectedNames = bookmarks;
      renderSelectedTags(store.selectedNames, store.availableKeys);
      return;
    }

    // Single bookmark
    const seachBox = document.getElementById('search-box') as HTMLInputElement;
    if (seachBox) {
      const processedName = bookmarks[0];
      const store = StoreService.getInstance();
      store.selectedNames = [processedName];
      selectMatchingBookmarkFromSelection(processedName);

      const aiTag = store.availableKeys.find(k =>
        k.AIFlag === 1 &&
        (k.DisplayName.toLowerCase() === processedName.toLowerCase() ||
          `id${k.ID}`.toLowerCase() === processedName.toLowerCase())
      );

      if (!aiTag) return;

      const appBody = document.getElementById('app-body');
      appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';

      appBody.innerHTML = await generateCheckboxHistory(aiTag, "AITag");
    }


  });
}

async function getBookmarksFromSelection(context: Word.RequestContext): Promise<string[]> {
  const selection = context.document.getSelection();
  const bookmarks = selection.getBookmarks();
  await context.sync();
  return bookmarks.value || [];
}


function normalizeBookmark(name: string) {
  return name.split('_Split_')[0].replace(/_/g, ' ');
}

function pickRelevantBookmarks(bookmarks: string[]) {
  // Remove duplicates & internal splits
  const normalized = Array.from(new Set(bookmarks.map(normalizeBookmark)));

  // Prefer AI tags only
  const store = StoreService.getInstance();
  return normalized.filter(name =>
    store.availableKeys.some(
      k => k.AIFlag === 1 &&
        (k.DisplayName.toLowerCase() === name.toLowerCase() ||
          `id${k.ID}`.toLowerCase() === name.toLowerCase())
    )
  );
}

async function getImages() {
  try {
    const store = StoreService.getInstance();
    const userId = sessionStorage.getItem('userId') || '0';

    // Fetch Images and Clients in parallel
    const generalImagesPromise = getGeneralImages(store.jwt);
    const documentImagesPromise = getReportHeadImageById(store.dataList.ID, store.jwt);
    const clientsPromise = getAllClients(userId, store.jwt);

    const [generalImages, documentImages, clientsData] = await Promise.all([
      generalImagesPromise,
      documentImagesPromise,
      clientsPromise
    ]);

    const mappedGeneral = mapImagesToComponentObjects(generalImages['Data']);
    const mappedDocument = mapImagesToComponentObjects(documentImages['Data']);

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
      toaster('Images and data are loaded and ready for use', 'success');
    }
  } catch (error) {
    console.error("Error loading background data:", error);
  }
}
