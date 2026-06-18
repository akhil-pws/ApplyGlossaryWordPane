// Imports
import { CONFIG } from "./utils/config";
import { AuthService } from "./services/auth.service";
import { DocumentService } from "./services/document.service";
import { UIService } from "./services/ui.service";
import { StoreService } from "./services/store.service";
import { AIService } from "./services/ai.service";
import { DocStorage } from "./utils/doc-storage";
// Note: GlossaryService import removed if unused or moved

// Restoration of variables needed by the rest of the file (Legacy Support - check if needed)
import { generateCheckboxHistory, getDateTimeStamp, initializeAIHistoryEvents, loadHomepage, replaceMention, setupPromptBuilderUI } from "./draft/home";
import { chatfooter, colorTable, insertLineWithHeadingStyle, mapImagesToComponentObjects, resolveWordTableStyle, selectMatchingBookmarkFromSelection, svgBase64ToPngBase64, switchModeIcon, switchToAddTag, switchToPromptBuilder, updateEditorFinalTable, parseHtmlTableToGrid, transposeGrid, detectTableCase, confirmSwitchChatHistory, applyCustomTextStyleToCell } from "./draft/draft-functions";
import { addtagbody, customizeTablePopup, customizeTextStylePopup, customizedStylePopup, logoheader, navTabs, toaster } from "./components/bodyelements";
import { addAiHistory, addGroupKey, fetchGlossaryTemplate, getAiHistory, getAllClients, getAllCustomTables, getAllPromptTemplates, getGeneralImages, getReportById, getReportHeadImageById, loginUser, updateGroupKey, getAllCustomTexts } from "./draft/draft.api";
import { wordTableStyles } from "./components/tablestyles";
import { mapApiStyleToCustomStyle } from "./components/customstyles";
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
        if (!CONFIG.environment.includes(props.environment) && props.environment !== 'unknown') {
          document.getElementById('app-body').innerHTML = `
        <p class="px-3 text-center">The document is not exported from this environment</p>`
          console.log(`Custom property "documentID" not found.`);
        } else {
          // Update local state for legacy compatibility
          // documentID = props.documentID; // Moved to Store
          // organizationName = props.organizationName; // Moved to Store
          const store = StoreService.getInstance();
          store.initForDocument(props.documentID, props.environment);
          store.organizationName = props.organizationName;
          store.environment = props.environment;
          if (props.URL) {
            CONFIG.dataUrl = props.URL
          }

          // Check Session
          const session = AuthService.restoreSession();
          if (session) {
            // Restore session state
            store.jwt = session.jwt;
            store.UserRole = session.userRole;
            if (session.tableStyle) store.tableStyle = session.tableStyle;
            if (session.colorPallete) store.colorPallete = session.colorPallete;
            if (session.defaultTextStyle) store.defaultTextStyle = session.defaultTextStyle;

            // Handle custom text style null properties fallback check
            if (store.customizedTextStyle && store.customizedTextStyle.properties === null) {
              getDocumentParagraphStyles().then((availableStyles) => {
                const styleNameInWord = availableStyles.find(
                  s => s.toLowerCase() === store.customizedTextStyle.name.toLowerCase() || s.toLowerCase() === store.customizedTextStyle.id.toLowerCase()
                );
                if (styleNameInWord) {
                  store.defaultTextStyle = styleNameInWord;
                  DocStorage.setItem("defaultTextStyle", styleNameInWord);
                  store.saveToStorage();
                }
              }).catch(e => console.error("Error restoring custom style fallback on load:", e));
            }

            window.location.hash = '#/dashboard';
            toaster('You are successfully logged in', 'success');
            displayMenu(); // Trigger legacy menu display
          } else {
            loadLoginPage();
          }
        }
      } else {
        document.getElementById('app-body').innerHTML = `
        <p class="px-3 text-center">Export a document from the LINK AI application to use this functionality.</p>`
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
  const sessionToken = DocStorage.getItem('token');
  const store = StoreService.getInstance();
  if (sessionToken) {
    store.UserRole = JSON.parse(DocStorage.getItem('userRole')) || ''
    store.jwt = sessionToken;
    window.location.hash = '#/dashboard';
    const style = DocStorage.getItem('tableStyle');
    if (style) {
      store.tableStyle = style;
    }
    const defaultTextStyle = DocStorage.getItem('defaultTextStyle');
    if (defaultTextStyle) {
      store.defaultTextStyle = defaultTextStyle;
    }
    const localPallete = DocStorage.getItem('colorPallete');
    if (localPallete) {
      store.colorPallete = JSON.parse(localPallete);
    }

    // Handle custom text style null properties fallback check
    if (store.customizedTextStyle && store.customizedTextStyle.properties === null) {
      getDocumentParagraphStyles().then((availableStyles) => {
        const styleNameInWord = availableStyles.find(
          s => s.toLowerCase() === store.customizedTextStyle.name.toLowerCase() || s.toLowerCase() === store.customizedTextStyle.id.toLowerCase()
        );
        if (styleNameInWord) {
          store.defaultTextStyle = styleNameInWord;
          DocStorage.setItem("defaultTextStyle", styleNameInWord);
          store.saveToStorage();
        }
      }).catch(e => console.error("Error restoring custom style fallback in login:", e));
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
    DocStorage.setItem('theme', store.theme);
  });
}

async function handleLogin(event) {
  event.preventDefault();
  UIService.toggleLoader(true);

  try {
    const organizationInput = (document.getElementById('organization') as HTMLInputElement).value;
    const username = (document.getElementById('username') as HTMLInputElement).value;
    const password = (document.getElementById('password') as HTMLInputElement).value;

    const store = StoreService.getInstance();
    const targetOrg = (store.organizationName || '').toLowerCase().trim();
    const enteredOrg = (organizationInput || '').toLowerCase().trim();

    if (enteredOrg === targetOrg && targetOrg !== '') {
      // Use AuthService
      const result = await AuthService.login(organizationInput, username, password);

      if (result.success) {
        const data = result.data;
        store.jwt = data.token;
        store.UserRole = data.userRole;
        store.userId = data.userId;
        store.saveToStorage();

        // Preserve legacy logic for style restoring
        const style = DocStorage.getItem('tableStyle');
        if (style) store.tableStyle = style;

        const defaultTextStyle = DocStorage.getItem('defaultTextStyle');
        if (defaultTextStyle) store.defaultTextStyle = defaultTextStyle;

        const localPallete = DocStorage.getItem('colorPallete');
        if (localPallete) store.colorPallete = JSON.parse(localPallete);

        toaster('You are successfully logged in', 'success');
        displayMenu();
        window.location.hash = '#/dashboard';
      } else {
        showLoginError(result.message || "Login failed");
      }
    } else {
      showLoginError("The organization specified is not associated with this document");
    }
  } catch (error) {
    console.error("Login process error:", error);
    showLoginError("An unexpected error occurred. Please try again.");
  } finally {
    UIService.toggleLoader(false);
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
  store.userId = Number(DocStorage.getItem('userId'))
  // document.getElementById('aitag').addEventListener('click', redirectAI);
  fetchDocument('Init');

}

async function getTableStyle() {
  const store = StoreService.getInstance();
  const tableStyleObj = await getAllCustomTables(store.jwt);
  store.customTableStyle = tableStyleObj['Data'];
  const selectedTable = store.customTableStyle.find(style => style.ID === store.dataList.TableCustomizationID);
  if (selectedTable) {
    DocStorage.setItem("CustomStyle", selectedTable ? selectedTable.Name : '');
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
  const store = StoreService.getInstance();
  try {
    const textStyleObj = await getAllCustomTexts(store.jwt);
    if (textStyleObj && textStyleObj.Status && Array.isArray(textStyleObj.Data)) {
      store.customizedStyles = textStyleObj.Data.map(mapApiStyleToCustomStyle);
      const selectedTextStyle = store.customizedStyles.find(
        style => style.ID === store.dataList.TextCustomizationID || style.id === store.dataList.TextCustomizationID
      );
      if (selectedTextStyle) {
        if (selectedTextStyle.properties === null) {
          try {
            const availableStyles = await getDocumentParagraphStyles();
            const styleNameInWord = availableStyles.find(
              s => s.toLowerCase() === selectedTextStyle.name.toLowerCase() || s.toLowerCase() === selectedTextStyle.id.toLowerCase()
            );
            if (styleNameInWord) {
              store.defaultTextStyle = styleNameInWord;
              DocStorage.setItem("defaultTextStyle", styleNameInWord);
            }
          } catch (e) {
            console.error("Error checking matching Word style:", e);
          }
        }
        store.customizedTextStyle = selectedTextStyle;
        DocStorage.setItem("customTextStyleId", selectedTextStyle.id);
        DocStorage.setItem("customTextStyle", JSON.stringify(selectedTextStyle));
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
  UIService.toggleLoader(true);
  try {
    const store = StoreService.getInstance();
    const userId = DocStorage.getItem('userId') || '0';
    const reportData = await DocumentService.loadReportData(store.documentID, store.jwt, userId);

    // Assign to store
    store.dataList = reportData.dataList;
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
      onHome: () => {
        confirmSwitchChatHistory(async () => {
          if (!store.isPendingResponse) {
            if (store.isGlossaryActive) await removeMatchingContentControls();
            loadHomepage(store.availableKeys);
          }
          store.mode = 'Home';
          switchModeIcon();
        });
      },
      onSummary: () => {
        confirmSwitchChatHistory(async () => {
          if (!store.isPendingResponse) {
            if (store.isGlossaryActive) await removeMatchingContentControls();
            loadSummarypage(store.availableKeys);
          }
          store.mode = 'Summary';
          switchModeIcon();
        });
      },
      onGlossary: () => {
        confirmSwitchChatHistory(() => {
          if (store.emptyFormat) fetchGlossary();
        });
      },
      onFormat: () => {
        confirmSwitchChatHistory(() => {
          if (!store.isPendingResponse) formatOptionsDisplay();
        });
      },
      onRemoveFormat: () => {
        confirmSwitchChatHistory(() => {
          if (Object.keys(store.capturedFormatting).length > 0) removeOptionsConfirmation();
        });
      },
      onThemeToggle: () => {
        store.theme = store.theme === 'Light' ? 'Dark' : 'Light';
        UIService.applyTheme(store.theme as 'Light' | 'Dark');
        DocStorage.setItem('theme', store.theme);
      },
      onLogout: () => {
        confirmSwitchChatHistory(async () => {
          if (!store.isPendingResponse) {
            if (store.isGlossaryActive) await removeMatchingContentControls();
            logout();
          }
        });
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
  AuthService.logout();
  DocStorage.clearAll();
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
            for (const line of txt.split(/\r?\n/)) {
              if (!line.trim()) {
                const p = cursor.insertParagraph("", Word.InsertLocation.after);
                insertLineWithHeadingStyle(p, "");
                include(p.getRange());
                cursor = p.getRange();
                continue;
              }

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

              let grid = parseHtmlTableToGrid(rows);
              const tableCase = detectTableCase(grid);

              const store = StoreService.getInstance();
              const base = store.tableStyle.split(" - ")[0].trim();

              if (base === 'Table Grid 2') {
                store.isReversed = true;
              } else {
                store.isReversed = false;
              }
              if (store.isReversed && tableCase !== "CASE_1") {
                grid = transposeGrid(grid);
              }

              const numRows = grid.length;
              const numCols = grid[0]?.length || 0;

              const p = cursor.insertParagraph("", Word.InsertLocation.after);
              const table = p.insertTable(
                numRows,
                numCols,
                Word.InsertLocation.after
              );

              const resolvedTableStyle = resolveWordTableStyle(store.tableStyle);
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
                      applyCustomTextStyleToCell(tableCell, store);
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
                      (topCell as any).merge(bottomCell);
                      try {
                        topCell.verticalAlignment = Word.VerticalAlignment.center;
                        topCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered as any;
                      } catch (e) { }
                    }
                  });
                }
              } else {
                // Manual population for transposed table
                grid.forEach((rowGrid, rowIndex) => {
                  rowGrid.forEach((cellValue, cellIndex) => {
                    const tableCell = table.getCell(rowIndex, cellIndex);
                    tableCell.value = cellValue;
                    applyCustomTextStyleToCell(tableCell, store);
                  });
                });
              }

              // Styling logic
              if (store.colorPallete.Customize) {
                await colorTable(table, rows, context, store.isReversed);
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
                  insertLineWithHeadingStyle(p, "");
                  include(p.getRange());
                  cursor = p.getRange();
                  continue;
                }
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

        for (const line of txt.split(/\r?\n/)) {
          if (!line.trim()) {
            const p = cursor.insertParagraph("", Word.InsertLocation.after);
            insertLineWithHeadingStyle(p, "");
            include(p.getRange());
            cursor = p.getRange();
            continue;
          }

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

      const searchPromises = Array.from(store.filteredGlossaryTerm).map((term: any) => {
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

  // Handle Home or Summary mode - detect bookmarks/tags in selection
  if (store.mode === 'Home' || store.mode === 'Summary') {
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
      const updateSponsorSelectAllState = () => {
        const selectAllCb = document.getElementById('sponsorSelectAll') as HTMLInputElement;
        if (!selectAllCb) return;
        const individualSponsors = Array.from(sponsorDropdownItems).filter(cb => cb.id !== 'sponsorSelectAll') as HTMLInputElement[];
        if (individualSponsors.length === 0) {
          selectAllCb.checked = false;
          return;
        }
        const allChecked = individualSponsors.every(cb => cb.checked);
        selectAllCb.checked = allChecked;
      };

      const updateSourceSelectAllState = () => {
        const selectAllCb = document.getElementById('sourceSelectAll') as HTMLInputElement;
        if (!selectAllCb) return;
        const individualSources = Array.from(sourceDropdownItems).filter(cb => cb.id !== 'sourceSelectAll') as HTMLInputElement[];
        if (individualSources.length === 0) {
          selectAllCb.checked = false;
          return;
        }
        const allChecked = individualSources.every(cb => cb.checked);
        selectAllCb.checked = allChecked;
      };
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
        updateSponsorSelectAllState();
      };

      const updateSourceDropdownLabel = () => {
        const selectedNames = Array.from(sourceDropdownItems)
          .filter(cb => (cb as HTMLInputElement).checked && cb.id !== 'sourceSelectAll')
          .map(cb => cb.parentElement.querySelector('label').textContent.trim());

        const labelSpan = document.getElementById('sourceDropdownLabel');
        if (labelSpan) {
          labelSpan.textContent = selectedNames.length
            ? selectedNames.join(", ")
            : "Select Source Types";
        }
        updateSourceSelectAllState();
      }

      // Submit Handler
      form.addEventListener('submit', async (e) => {
        e.preventDefault();

        form.querySelectorAll('.is-invalid').forEach(i => i.classList.remove('is-invalid'));

        let valid = true;

        if (!nameField.value.trim()) { nameField.classList.add('is-invalid'); valid = false; }
        if (!promptField.value.trim()) { promptField.classList.add('is-invalid'); valid = false; }

        // SOURCE VALIDATION
        let selectedPrimarySources = [];
        selectedPrimarySources = Array.from(sourceDropdownItems)
          .filter(cb_node => (cb_node as HTMLInputElement).checked && (cb_node as HTMLInputElement).id !== 'sourceSelectAll')
          .map(cb_node => (cb_node as HTMLInputElement).value);

        if (!selectedPrimarySources.length && !isSummaryMode) {
          document.getElementById("primarySourceError").style.display = "block";
          valid = false;
        } else {
          document.getElementById("primarySourceError").style.display = "none";
        }

        if (!valid) return;

        const selectedSponsors = Array.from(sponsorDropdownItems)
          .filter(cb_node => (cb_node as HTMLInputElement).checked && (cb_node as HTMLInputElement).id !== 'sponsorSelectAll')
          .map(cb_node => (store.clientList.find(c => c.ID == (cb_node as HTMLInputElement).value) as any));

        const selectedSources = Array.from(sourceDropdownItems)
          .filter(cb => (cb as HTMLInputElement).checked && cb.id !== 'sourceSelectAll')
          .map(cb => sourceTypeList.find(s => s.ID == (cb as HTMLInputElement).value));

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
          SummaryTagClient: isSummaryMode ? selectedSponsors.map(s => ({ ClientID: s.ID, Client: s.Name })) : [],

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
          const cb = cb_node as HTMLInputElement;
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
          const checkbox = this.querySelector('.form-check-input') as HTMLInputElement;
          if (!checkbox) return;

          const target = e.target as HTMLElement;
          if (target !== checkbox && target.tagName !== 'LABEL') {
            if (!checkbox.disabled) {
              checkbox.checked = !checkbox.checked;
            }
          }
          if (checkbox.id === 'sponsorSelectAll') {
            const isChecked = checkbox.checked;
            sponsorDropdownItems.forEach(cb_node => {
              const cb = cb_node as HTMLInputElement;
              if (!cb.disabled) cb.checked = isChecked;
            });
          }

          updateSponsorDropdownLabel();
        });
      });

      document.querySelectorAll('.source-dropdown-item').forEach(item => {
        item.addEventListener('click', function (e) {
          e.stopPropagation();
          const checkbox = this.querySelector('.form-check-input') as HTMLInputElement;
          if (!checkbox) return;
          const target = e.target as HTMLElement;
          if (target !== checkbox && target.tagName !== 'LABEL') {
            if (!checkbox.disabled) {
              checkbox.checked = !checkbox.checked;
            }
          }

          if (checkbox.id === 'sourceSelectAll') {
            const isChecked = checkbox.checked;
            sourceDropdownItems.forEach(cb => {
              const cbInput = cb as HTMLInputElement;
              if (!cbInput.disabled) {
                cbInput.checked = isChecked;
              }
            });
          }

          updateSourceDropdownLabel();
          const selectedCount = Array.from(sourceDropdownItems)
            .filter(cb => (cb as HTMLInputElement).checked).length;

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

  const customStyleName = DocStorage.getItem("CustomStyle") || "";
  const defaultStyle = DocStorage.getItem("DefaultStyle") || store.tableStyle;
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
        DocStorage.setItem("CustomStyle", styleObj.Name);
        store.tableStyle = styleObj.Setting.BaseStyle; // stores full object as 
      } else {
        store.colorPallete.Customize = false;
        store.tableStyle = dropdown.value; // normal style string
        DocStorage.setItem("DefaultStyle", store.tableStyle);

      }

      DocStorage.setItem("colorPallete", JSON.stringify(store.colorPallete));
      DocStorage.setItem("tableStyle", store.tableStyle);

      container.innerHTML = "";
    });
  }
}

export async function getDocumentParagraphStyles(): Promise<string[]> {
  return Word.run(async (context) => {
    try {
      const styles = context.document.getStyles();
      styles.load("items/nameLocal,items/type");
      await context.sync();

      const paragraphStyles = styles.items
        .filter(style => style.type === "Paragraph" || (Word.StyleType && style.type === Word.StyleType.paragraph))
        .map(style => style.nameLocal)
        .sort((a, b) => a.localeCompare(b));

      return paragraphStyles.length > 0 ? paragraphStyles : ["Normal", "Body Text", "No Spacing"];
    } catch (error) {
      console.error("Failed to load document styles:", error);
      return ["Normal", "Body Text", "No Spacing"];
    }
  });
}

export async function getDocumentStyleDetails(styleName: string) {
  return Word.run(async (context) => {
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

export async function customizeTextStyle() {
  const store = StoreService.getInstance();
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

  container.innerHTML = customizeTextStylePopup(currentStyle, availableStyles);

  const cancelBtn = document.getElementById("text-style-popup-cancel");
  const okBtn = document.getElementById("text-style-popup-confirm");
  const dropdown = document.getElementById("text-style-dropdown") as HTMLSelectElement;
  const preview = document.getElementById("text-style-preview") as HTMLDivElement;

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
      DocStorage.setItem("defaultTextStyle", store.defaultTextStyle);

      // Clear customized style when default text style is selected
      store.customizedTextStyle = null;
      DocStorage.removeItem("customTextStyleId");
      DocStorage.removeItem("customTextStyle");

      store.saveToStorage();

      toaster("Default text style saved successfully", "success");
      container.innerHTML = "";
    });
  }
}

export async function customizeCustomStyle() {
  const store = StoreService.getInstance();
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
  const currentStyleId = (store.customizedTextStyle && store.customizedTextStyle.id) || DocStorage.getItem("customTextStyleId") || (stylesList[0] ? stylesList[0].id : "");
  container.innerHTML = customizedStylePopup(currentStyleId, stylesList);

  const cancelBtn = document.getElementById("customized-style-popup-cancel");
  const okBtn = document.getElementById("customized-style-popup-confirm");
  const dropdown = document.getElementById("customized-style-dropdown") as HTMLSelectElement;
  const preview = document.getElementById("customized-style-preview") as HTMLDivElement;

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
          const styleNameInWord = availableStyles.find(
            s => s.toLowerCase() === selectedStyle.name.toLowerCase() || s.toLowerCase() === selectedStyle.id.toLowerCase()
          );
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
            const styleNameInWord = availableStyles.find(
              s => s.toLowerCase() === selectedStyle.name.toLowerCase() || s.toLowerCase() === selectedStyle.id.toLowerCase()
            );
            if (styleNameInWord) {
              store.defaultTextStyle = styleNameInWord;
              DocStorage.setItem("defaultTextStyle", styleNameInWord);
            }
          } catch (e) {
            console.error("Error setting default style from Word style:", e);
          }
        }
        store.customizedTextStyle = selectedStyle;
        DocStorage.setItem("customTextStyleId", selectedStyle.id);
        DocStorage.setItem("customTextStyle", JSON.stringify(selectedStyle));
        store.saveToStorage();
        toaster("Customized style applied successfully", "success");
      }
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
        SummaryTagClient: payload.SummaryTagClient
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
  const store = StoreService.getInstance();
  if (store.currentChatTagId !== -1 && store.currentChatTagId !== undefined && store.currentChatTagId !== null) {
    return;
  }
  return Word.run(async (context) => {
    const selection = context.document.getSelection();

    const rawBookmarks = await getBookmarksFromSelection(context); // internally calls context.sync()
    const bookmarks = pickRelevantBookmarks(rawBookmarks);

    try {
      const store = StoreService.getInstance();
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
        let match: RegExpExecArray | null;

        const matchedNames = [...bookmarks];

        // Helper to find matching tag in current mode
        const findTag = (tagName: string) => {
          if (store.mode === 'Home') {
            return store.availableKeys.find(k =>
              k.AIFlag === 1 &&
              (k.DisplayName.toLowerCase() === tagName.toLowerCase() ||
                `id${k.ID}`.toLowerCase() === tagName.toLowerCase())
            );
          } else if (store.mode === 'Summary') {
            return store.summaryTagList?.find(k =>
              k.Name?.toLowerCase() === tagName.toLowerCase() ||
              `sm${k.ID || k.ReportHeadSummaryTagID}`.toLowerCase() === tagName.toLowerCase()
            );
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
                    return existingName.toLowerCase() === nameToPush.toLowerCase() ||
                      existingName.toLowerCase() === `id${tag.ID}`.toLowerCase();
                  } else {
                    return existingName.toLowerCase() === nameToPush.toLowerCase() ||
                      existingName.toLowerCase() === `sm${tag.ID || tag.ReportHeadSummaryTagID}`.toLowerCase();
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
          document.getElementById('tags-in-selected-text')
            ?.classList.replace('d-none', 'd-block');
          store.selectedNames = matchedNames;
          renderSelectedTags(store.selectedNames, store.availableKeys);
          return;
        } else if (matchedNames.length === 1) {
          const singleName = matchedNames[0];
          const tag = findTag(singleName);
          if (tag) {
            confirmSwitchChatHistory(async () => {
              const appBody = document.getElementById('app-body');
              appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';

              await selectMatchingBookmarkFromSelection(singleName);

              if (store.mode === 'Home') {
                appBody.innerHTML = await generateCheckboxHistory(tag, "AITag");
              } else if (store.mode === 'Summary') {
                appBody.innerHTML = await generateCheckboxHistory(tag, "Summary");
              }

              document.getElementById('tags-in-selected-text')
                ?.classList.replace('d-none', 'd-block');
              store.selectedNames = [singleName];
              renderSelectedTags(store.selectedNames, store.availableKeys);
            });
            return;
          }
        }
      }
    } catch (e) {
      console.error('Tag placeholder detection error:', e);
    }

    document.getElementById('tags-in-selected-text')
      ?.classList.replace('d-block', 'd-none');
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
  return normalized.filter(name => {
    if (store.mode === 'Home') {
      return store.availableKeys.some(
        k => k.AIFlag === 1 &&
          (k.DisplayName.toLowerCase() === name.toLowerCase() ||
            `id${k.ID}`.toLowerCase() === name.toLowerCase())
      );
    } else if (store.mode === 'Summary') {
      return store.summaryTagList.some(
        k => (k.Name?.toLowerCase() === name.toLowerCase() ||
          `sm${k.ID || k.ReportHeadSummaryTagID}`.toLowerCase() === name.toLowerCase())
      );
    }
    return false;
  });
}

async function getImages() {
  try {
    const store = StoreService.getInstance();
    const userId = DocStorage.getItem('userId') || '0';

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
