/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import { dataUrl, storeUrl, versionLink } from "./data";
import { generateCheckboxHistory, initializeAIHistoryEvents, loadHomepage, setupPromptBuilderUI } from "./components/home";
import { applyThemeClasses, chatfooter, colorTable, renderSelectedTags, selectMatchingBookmarkFromSelection, swicthThemeIcon, switchToAddTag, switchToPromptBuilder, updateEditorFinalTable } from "./functions";
import { addtagbody, customizeTablePopup, logoheader, navTabs, toaster } from "./components/bodyelements";
import { addAiHistory, addGroupKey, fetchGlossaryTemplate, getAiHistory, getAllClients, getAllCustomTables, getAllPromptTemplates, getReportById, loginUser, updateGroupKey } from "./api";
import { wordTableStyles } from "./components/tablestyles";
export let jwt = '';
export let UserRole: any = {};
let storedUrl = storeUrl
let documentID = ''
let organizationName = ''
export let aiTagList = [];
let initialised = true;
export let availableKeys = [];
let promptBuilderList = [];
let glossaryName = ''
let isGlossaryActive: boolean = false;
let GroupName: string = '';
let layTerms = [];
let dataList: any = []
let isTagUpdating: boolean = false;
let capturedFormatting: any = {};
let emptyFormat: boolean = false;
let isNoFormatTextAvailable: boolean = false;
let clientId = '0';
let userId = 0;
let clientList = [];
let version = versionLink;
let currentYear = new Date().getFullYear();
export let sourceList;
let filteredGlossaryTerm;
export let selectedNames = [];
export let isPendingResponse = false;
export let theme = 'Light';
export let tableStyle = 'Grid Table 4 - Accent 1';
export let colorPallete: any = {
  "Header": '#FFFFFF',
  "Primary": '#FFFFFF',
  "Secondary": '#FFFFFF',
  "Customize": false
};

export let customTableStyle = [];

/* global document, Office, Word */

window.addEventListener('hashchange', () => {
  const hash = window.location.hash;
  if (hash === '#/dashboard' && initialised) {
    initialised = false;
    displayMenu();

  }
});


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("footer").innerText = `© ${currentYear} - TrialAssure LINK AI Assistant ${version}`
    const editor = document.getElementById('editor');

    window.location.hash = '#/login';
    retrieveDocumentProperties()

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => {
        logBookmarksInSelection();
      }
    );
  }
});


// Example usage:



async function retrieveDocumentProperties() {
  try {
    await Word.run(async (context) => {
      const properties = context.document.properties.customProperties;
      properties.load("items");

      await context.sync();
      const property = properties.items.find(prop => prop.key === 'DocumentID');
      const orgName = properties.items.find(prop => prop.key === 'Organization');
      if (property && orgName) {
        documentID = property.value;
        organizationName = orgName.value;
        login()
      } else {
        document.getElementById('app-body').innerHTML = `
        <p class="px-3 text-center">Export a document from the LINK AI application to use this functionality.</p>`
        console.log(`Custom property "documentID" not found.`);
        return null;
      }
    });
  } catch (error) {
    console.error("Error retrieving custom property:", error);
  }

}

async function login() {
  // document.getElementById('header').innerHTML = ``
  const sessionToken = sessionStorage.getItem('token');
  if (sessionToken) {
    UserRole = JSON.parse(sessionStorage.getItem('userRole')) || ''
    jwt = sessionToken;
    window.location.hash = '#/dashboard';
    const style = localStorage.getItem('tableStyle');
    if (style) {
      tableStyle = style;
    }
    const localPallete = localStorage.getItem('colorPallete');
    if (localPallete) {
      colorPallete = JSON.parse(localPallete);
    }

  } else {
    loadLoginPage();
  }
}

function loadLoginPage() {

  document.getElementById('logo-header').innerHTML = `
  <img id="main-logo" src="${storedUrl}/assets/logo.png" alt="" class="logo">
  <div class="icon-nav me-3">
    <span id="theme-toggle"><i class="fa fa-moon c-pointer me-3"  title="Toggle Theme"></i><span>
  </div>
`;

  document.getElementById('app-body').innerHTML = `
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
    </div>
  `;
  document.getElementById('theme-toggle').addEventListener('click', () => {
    theme = theme === 'Light' ? 'Dark' : 'Light';
    applyThemeClasses(theme)

    document.body.classList.toggle('dark-theme', theme === 'Dark');
    document.body.classList.toggle('light-theme', theme === 'Light');
    swicthThemeIcon()
  }
  );
  document.getElementById('login-form').addEventListener('submit', handleLogin);
}

async function handleLogin(event) {
  event.preventDefault();

  // Get the values from the form fields
  const organization = (document.getElementById('organization') as HTMLInputElement).value;
  const username = (document.getElementById('username') as HTMLInputElement).value;
  const password = (document.getElementById('password') as HTMLInputElement).value;
  if (organization.toLowerCase().trim() === organizationName.toLocaleLowerCase().trim()) {
    document.getElementById('app-body').innerHTML = `
  <div id="button-container">

          <div class="loader" id="loader"></div>
          </div
`
    try {
      const data = await loginUser(organization, username, password);
      if (data.Status === true && data['Data']) {
        if (data['Data'].ResponseStatus) {
          jwt = data.Data.Token;
          const style = localStorage.getItem('tableStyle');
          if (style) {
            tableStyle = style;
          }

          const localPallete = localStorage.getItem('colorPallete');
          if (localPallete) {
            colorPallete = JSON.parse(localPallete);
          }
          UserRole = data.Data.UserRole;
          sessionStorage.setItem('userRole', JSON.stringify(data.Data.UserRole));
          sessionStorage.setItem('token', jwt)
          sessionStorage.setItem('userId', data.Data.ID);
          toaster('You are successfully logged in', 'success');
          window.location.hash = '#/dashboard';
        } else {
          showLoginError("An error occurred during login. Please try again.")
        }
      } else {
        showLoginError("An error occurred during login. Please try again.")
      }
    } catch (error) {
      showLoginError("An error occurred during login. Please try again.")
      console.error('Error during login:', error);
      // Handle login error (e.g., show an error message)
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
  userId = Number(sessionStorage.getItem('userId'))
  // document.getElementById('aitag').addEventListener('click', redirectAI);
  fetchDocument('Init');

}

async function getTableStyle() {
  const tableStyle = await getAllCustomTables(jwt);
  customTableStyle = tableStyle['Data'];
}

async function fetchDocument(action) {
  try {

    const data = await getReportById(documentID, jwt);
    document.getElementById('app-body').innerHTML = ``
    document.getElementById('logo-header').innerHTML = logoheader(storedUrl);

    dataList = data['Data'];
    console.log(dataList.Group[0]);
    getTableStyle();
    sourceList = dataList?.SourceTypeList?.filter(
      (item) => item.SourceValue !== ''
        && item.AIFlag === 1
    ) // Filter items with an extension
      .map((item) => ({
        ...item, // Spread the existing properties
        SourceName: decodeURIComponent(transformDocumentName(item.SourceValue))
      }));
    clientId = dataList.ClientID;
    const aiGroup = data['Data'].Group.find(element => element.DisplayName === 'AIGroup');
    GroupName = aiGroup ? aiGroup.Name : '';
    aiTagList = aiGroup ? aiGroup.GroupKey : [];

    availableKeys = data['Data'].GroupKeyAll.filter(element => element.ComponentKeyDataType === 'TABLE' || element.ComponentKeyDataType === 'TEXT');
    availableKeys.forEach((key) => {
      if (key.AIFlag === 1) {
        const regex = /<TableStart>([\s\S]*?)<TableEnd>/gi;

        let match;
        if ((match = regex.exec(key.EditorValue) !== null)) {
          {
            key.EditorValue = updateEditorFinalTable(key.EditorValue);
            key.UserValue = key.EditorValue;
            key.InitialTable = true;
            key.ComponentKeyDataType = 'TABLE';
          }

        }
      }
    });

    aiTagList.forEach((key, i) => {
      const regex = /<TableStart>([\s\S]*?)<TableEnd>/gi;

      let match;
      if ((match = regex.exec(key.EditorValue) !== null)) {
        {
          key.EditorValue = updateEditorFinalTable(key.EditorValue);
          key.UserValue = key.EditorValue;
          key.InitialTable = true;
          key.ComponentKeyDataType = 'TABLE';
        }

      }
    }

    );
    fetchClients();
    loadPromptTemplates();
    loadHomepage(availableKeys);
    document.getElementById('home').addEventListener('click', async () => {
      if (!isPendingResponse) {
        if (isGlossaryActive) {
          await removeMatchingContentControls();
        }

        loadHomepage(availableKeys);
      }
    });

    document.getElementById('glossary').addEventListener('click', () => {
      if (emptyFormat) {
        fetchGlossary();
      }
    });

    document.getElementById('define-formatting').addEventListener('click', () => {
      if (!isPendingResponse) {
        formatOptionsDisplay()
      }
    }
    );


    document.getElementById('removeFormatting').addEventListener('click', () => {
      if (Object.keys(capturedFormatting).length > 0) {
        removeOptionsConfirmation();
      }
    });


    document.getElementById('theme-toggle').addEventListener('click', () => {
      theme = theme === 'Light' ? 'Dark' : 'Light';
      applyThemeClasses(theme)

      document.body.classList.toggle('dark-theme', theme === 'Dark');
      document.body.classList.toggle('light-theme', theme === 'Light');
      swicthThemeIcon()
    }
    );

    document.getElementById('logout').addEventListener('click', async () => {
      if (!isPendingResponse) {
        if (isGlossaryActive) {
          await removeMatchingContentControls();
        }

        logout()
      }
    }
    );

  } catch (error) {
    console.error('Error fetching glossary data:', error);
  }
}

async function fetchClients() {
  try {
    const userId = sessionStorage.getItem('userId') || '';


    const data = await getAllClients(userId, jwt);

    if (data.Status && data.Data) {
      clientList = data['Data'];
    } else {
      console.warn("Failed to load clients or no clients found.");
    }
  } catch (error) {
  }
}



export async function formatOptionsDisplay() {
  if (!isTagUpdating) { // Check if isTagUpdating is false
    if (isGlossaryActive) {
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
    if (Object.keys(capturedFormatting).length === 0) {
      const formatDetails = document.getElementById("format-details");
      formatDetails.style.display = 'none';
      // The object is not empty
    }

    const glossaryBtn = document.getElementById('glossary') as HTMLButtonElement;
    if (!glossaryBtn.classList.contains('disabled-link')) {
      glossaryBtn.classList.add('disabled-link');
    }

    if (emptyFormat) {
      clearCapturedFormatting();
    }
    else {
      if (capturedFormatting.Bold === null || capturedFormatting.Bold === undefined ||
        capturedFormatting.Underline === 'Mixed' || capturedFormatting.Underline === undefined ||
        capturedFormatting.Size === null || capturedFormatting.Size === undefined ||
        capturedFormatting["Font Name"] === null || capturedFormatting["Font Name"] === undefined ||
        capturedFormatting["Background Color"] === '' || capturedFormatting["Background Color"] === undefined ||
        capturedFormatting["Text Color"] === '' || capturedFormatting["Text Color"] === undefined) {
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
    if (isNoFormatTextAvailable) {
      emptyFormatCheckbox.checked = true;
      clearCapturedFormatting();
    }

    emptyFormatCheckbox.addEventListener("change", () => {
      if (emptyFormatCheckbox.checked) {
        isNoFormatTextAvailable = true;
        clearCapturedFormatting();
      } else {
        const CaptureBtn = document.getElementById('capture-format-btn') as HTMLButtonElement;
        CaptureBtn.disabled = false;
        isNoFormatTextAvailable = false;
        emptyFormat = false;
        const glossaryBtn = document.getElementById('glossary') as HTMLButtonElement;
        if (!glossaryBtn.classList.contains('disabled-link')) {
          glossaryBtn.classList.add('disabled-link');
        }
      }
    });

  }
}



function displayCapturedFormatting() {
  emptyFormat = false;
  const formatList = document.getElementById("format-list");
  formatList.innerHTML = ""; // Clear the list before adding new items

  for (const [key, value] of Object.entries(capturedFormatting)) {
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
  capturedFormatting = {}; // Clear the captured formatting object
  const formatDetails = document.getElementById("format-details");
  formatDetails.style.display = 'none';
  // formatList.innerHTML = `<li>No formatting selected.</li>`;
  emptyFormat = true;
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

      capturedFormatting = {
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

      if (capturedFormatting.Bold === null ||
        capturedFormatting.Underline === 'Mixed' ||
        capturedFormatting.Size === null ||
        capturedFormatting["Font Name"] === null ||
        capturedFormatting["Background Color"] === '' ||
        capturedFormatting["Text Color"] === ''

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
  if (!isTagUpdating) {
    if (isGlossaryActive) {
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

    if (capturedFormatting['Background Color'] === null &&
      capturedFormatting['Text Color'] === '#000000') {
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

      // Iterate through each paragraph in the document body
      for (const paragraph of paragraphs.items) {

        // Check if the paragraph contains text
        if (paragraph.text.trim() !== "") {
          const textRanges = paragraph.split([" "]); // Split paragraph into individual words/segments
          textRanges.load("items, font");

          await context.sync();

          for (const range of textRanges.items) {
            const font = range.font;
            font.load(["bold", "italic", "underline", "size", "highlightColor", "name", "color"]);

            await context.sync();

            // Check if the text range matches the captured formatting
            if (
              font.highlightColor === capturedFormatting['Background Color'] &&
              font.color === capturedFormatting['Text Color'] &&
              font.bold === capturedFormatting['Bold'] &&
              font.italic === capturedFormatting['Italic'] &&
              font.size === capturedFormatting['Size'] &&
              font.underline === capturedFormatting['Underline'] &&
              font.name === capturedFormatting['Font Name']
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
      capturedFormatting = {}; // Clear the captured formatting object
      const formatDetails = document.getElementById("format-details");
      formatDetails.style.display = 'none';
      // formatList.innerHTML = `<li>No formatting selected.</li>`;
      emptyFormat = true;
      isNoFormatTextAvailable = true;
      const glossaryBtn = document.getElementById('glossary') as HTMLButtonElement;
      glossaryBtn.classList.remove('disabled-link');
      formatOptionsDisplay()
    });
  } catch (error) {
    console.error("Error removing formatted text:", error);
  }
}


export async function fetchAIHistory(tag) {
  try {

    const data = await getAiHistory(tag.ID, jwt);


    if (data.Status && data.Data) {
      tag.ReportHeadAIHistoryList = data['Data'] || [];
      tag.FilteredReportHeadAIHistoryList = [];
      tag.SourceValueID = tag.ReportHeadAIHistoryList[0].SourceValue;
      const selectedSources = sourceList.filter((list) =>
        tag.SourceValueID.includes(String(list.VectorID))
      );

      tag.SourceName = selectedSources.map((item) => {
        return item.SourceName;
      });
      tag.Sources = tag.SourceName.join(',');
      tag.TempSourceValue = selectedSources.map((item) => {
        return item.VectorID ? String(item.VectorID) : item.SourceValue;
      });


      tag.ReportHeadAIHistoryList.forEach((historyList, i) => {
        historyList.Response = removeQuotes(historyList.Response);
        tag.FilteredReportHeadAIHistoryList.unshift(historyList);

      });
      return tag.FilteredReportHeadAIHistoryList;
      // Use the data here
    } else {
      console.warn("No AI history available.");
    }


  } catch (error) {
    console.error('Error fetching AI history:', error);
    return [];
  }
}

export async function sendPrompt(tag, prompt) {
  if (prompt !== '' && !isTagUpdating) {

    isTagUpdating = true;

    const iconelement = document.getElementById(`sendPromptButton`);
    iconelement.innerHTML = `<i class="fa fa-spinner fa-spin text-white"></i>`;

    const payload = {
      ReportHeadID: tag.FilteredReportHeadAIHistoryList[0].ReportHeadID,
      DocumentID: dataList.NCTID,
      DocumentType: dataList.DocumentType,
      TextSetting: dataList.TextSetting,
      DocumentTemplate: dataList.ReportTemplate,
      ReportHeadGroupKeyID: tag.FilteredReportHeadAIHistoryList[0].ReportHeadGroupKeyID,
      ThreadID: tag.ThreadID,
      AssistantID: dataList.AssistantID,
      Container: dataList.Container,
      GroupName: GroupName,
      Prompt: prompt,
      PromptType: 1,
      Response: '',
      VectorID: dataList.VectorID,
      Selected: 0,
      ID: 0,
      SourceValue: tag.TempSourceValue ? tag.TempSourceValue : []
    };

    try {
      isPendingResponse = true;
      const data = await addAiHistory(payload, jwt);

      if (data['Data'] && data['Data'] !== 'false') {
        tag.ReportHeadAIHistoryList = JSON.parse(JSON.stringify(data['Data']));
        tag.FilteredReportHeadAIHistoryList = [];

        tag.ReportHeadAIHistoryList.forEach((historyList) => {
          historyList.Response = removeQuotes(historyList.Response);
          tag.FilteredReportHeadAIHistoryList.unshift(historyList);
        });
        const chat = tag.ReportHeadAIHistoryList[0];
        aiTagList.forEach(currentTag => {
          if (currentTag.ID === tag.ID) {
            const isTable = chat.FormattedResponse !== '';
            const finalResponse = chat.FormattedResponse
              ? '\n' + updateEditorFinalTable(chat.FormattedResponse)
              : chat.Response;


            currentTag.ComponentKeyDataType = isTable ? 'TABLE' : 'TEXT';
            currentTag.UserValue = finalResponse;
            currentTag.EditorValue = finalResponse;
            currentTag.text = finalResponse;
            currentTag.IsApplied = tag.IsApplied;

          }
        });

        availableKeys.forEach(currentTag => {
          if (currentTag.ID === tag.ID) {
            const isTable = chat.FormattedResponse !== '';
            const finalResponse = chat.FormattedResponse
              ? '\n' + updateEditorFinalTable(chat.FormattedResponse)
              : chat.Response;
            currentTag.ComponentKeyDataType = isTable ? 'TABLE' : 'TEXT';
            currentTag.UserValue = finalResponse;
            currentTag.EditorValue = finalResponse;
            currentTag.text = finalResponse;
            currentTag.IsApplied = tag.IsApplied;

          }
        })



        const appbody = document.getElementById('app-body')
        appbody.innerHTML = await generateCheckboxHistory(tag);
        isPendingResponse = false;

      }

      iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`;
      document.getElementById(`chatInput`).value = '';
      isTagUpdating = false;
      isPendingResponse = false;
      // sourceListBtn.disabled = false;

    } catch (error) {
      iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`;
      isTagUpdating = false;
      isPendingResponse = false;
      console.error('Error sending AI prompt:', error);
    }
  } else {
    console.error('No empty prompt allowed');
  }
}




// Your existing copyText function



async function logout() {
  if (isGlossaryActive) {
    await removeMatchingContentControls();
  }
  sessionStorage.clear();
  window.location.hash = '#/new';
  initialised = true;
  document.getElementById('logo-header').innerHTML = ``;
  login();
}

function addAccordionListeners() {
  const accordionButtons = document.querySelectorAll('.accordion-button');

  accordionButtons.forEach(button => {
    button.addEventListener('click', function () {
      const collapseElement = this.nextElementSibling;

      // Check if the element exists before accessing its classList
      if (collapseElement && collapseElement.classList) {
        collapseElement.classList.toggle('show');
      }
    });
  });
}

function addCopyListeners() {
  const copyIcons = document.querySelectorAll('.fa-copy');
  copyIcons.forEach(icon => {
    icon.addEventListener('click', function () {
      const textToCopy = this.closest('.p-2').querySelector('.form-control').textContent;
    });
  });
}

export async function applyAITagFn() {
  toaster("Please wait... applying AI tags", "info");

  return Word.run(async (context) => {
    try {
      const body = context.document.body;

      context.load(body, 'text');
      await context.sync();

      for (let i = 0; i < aiTagList.length; i++) {
        const tag = aiTagList[i];
        tag.EditorValue = removeQuotes(tag.EditorValue);

        const searchResults = body.search(`#${tag.DisplayName}#`, {
          matchCase: false,
          matchWholeWord: false,
        });
        context.load(searchResults, 'items');
        await context.sync();


        for (const item of searchResults.items) {
          if (tag.EditorValue !== "" && !tag.IsApplied) {
            const cleanDisplayName = tag.ID;
            const uniqueStr = new Date().getTime();
            const bookmarkName = `ID${cleanDisplayName}_Split_${uniqueStr}`;

            const paragraphs = item.paragraphs;
            context.load(paragraphs, 'text, font/hidden');
            await context.sync();

            let visibleParagraph = paragraphs.items.find(p => !p.font.hidden);
            if (visibleParagraph) {
              const startMarker = visibleParagraph.insertParagraph('[[BOOKMARK_START]]', Word.InsertLocation.before);
              await context.sync();
            }

            if (tag.ComponentKeyDataType === 'TABLE') {
              const range = item.getRange();
              const parser = new DOMParser();
              const doc = parser.parseFromString(tag.EditorValue, 'text/html');
              const bodyNodes = Array.from(doc.body.childNodes);

              range.delete();

              for (const node of bodyNodes) {
                if (node.nodeType === Node.TEXT_NODE) {
                  let textContent = node.textContent?.trim();

                  if (textContent) {
                    textContent = textContent.replace(/\n- /g, "\n• ");

                    textContent.split('\n').forEach(line => {
                      if (line.trim()) {
                        insertLineWithHeadingStyle(range, line);
                      }
                    });
                  }
                } else if (node.nodeType === Node.ELEMENT_NODE) {
                  const element = node as HTMLElement;

                  if (element.tagName.toLowerCase() === 'table') {
                    const rows = Array.from(element.querySelectorAll('tr'));

                    if (rows.length === 0) {
                      range.insertParagraph("[Empty Table]", Word.InsertLocation.before);
                      continue;
                    }

                    const maxCols = Math.max(...rows.map(row => {
                      return Array.from(row.querySelectorAll('td, th')).reduce((sum, cell) => {
                        return sum + (parseInt(cell.getAttribute('colspan') || '1', 10));
                      }, 0);
                    }));

                    const paragraph = range.insertParagraph("", Word.InsertLocation.before);
                    await context.sync();

                    const table = paragraph.insertTable(rows.length, maxCols, Word.InsertLocation.after);
                    table.style = tableStyle;
                    await context.sync();
                    if (colorPallete.Customize) {
                      await colorTable(table, rows, context);
                    }
                    const rowspanTracker: number[] = new Array(maxCols).fill(0);

                    rows.forEach((row, rowIndex) => {
                      const cells = Array.from(row.querySelectorAll('td, th'));
                      let cellIndex = 0;

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

                        table.getCell(rowIndex, cellIndex).value = cellText;

                        for (let i = 1; i < colspan; i++) {
                          if (cellIndex + i < maxCols) {
                            table.getCell(rowIndex, cellIndex + i).value = "";
                          }
                        }

                        if (rowspan > 1) {
                          for (let i = 0; i < colspan; i++) {
                            if (cellIndex + i < maxCols) {
                              rowspanTracker[cellIndex + i] = rowspan - 1;
                            }
                          }
                        }

                        cellIndex += colspan;
                      });
                    });
                  } else {
                    let elementText = element.innerText.trim();
                    if (elementText) {
                      elementText = elementText.replace(/\n- /g, "\n• ");

                      elementText.split('\n').forEach(line => {
                        if (line.trim()) {
                          insertLineWithHeadingStyle(range, line);
                        }
                      });
                    }
                  }
                }
              }

              await context.sync();
            } else {

              let text = tag.EditorValue.trim();
              text = text.replace(/\n- /g, "\n• ");
              // text = text.replace(/\n- /g, "\n    • ");

              // Now insert the updated text
              item.insertText(text, Word.InsertLocation.replace);

              await context.sync();


            }

            // 1. Find last visible paragraph of the replaced region
            const itemParagraphs = item.paragraphs;
            context.load(itemParagraphs, 'items, font/hidden');
            await context.sync();

            let lastVisiblePara = null;
            for (let p of itemParagraphs.items) {
              if (!p.font.hidden) lastVisiblePara = p;
            }

            // 2. Force end marker into visible para
            let endMarker = null;
            if (lastVisiblePara) {
              endMarker = lastVisiblePara.insertParagraph('[[BOOKMARK_END]]', Word.InsertLocation.after);
              await context.sync();
            }


            const markers = context.document.body.paragraphs;
            context.load(markers, 'text');
            await context.sync();

            const start = markers.items.find(p => p.text === '[[BOOKMARK_START]]');
            const end = markers.items.find(p => p.text === '[[BOOKMARK_END]]');

            if (start && end) {
              const bookmarkRange = start.getRange('Start').expandTo(end.getRange('End'));
              bookmarkRange.insertBookmark(bookmarkName);
              console.log(`Bookmark added: ${bookmarkName}`);
              const afterBookmark = end.insertParagraph("", Word.InsertLocation.after);

              afterBookmark.select();
              start.delete();
              end.delete();
              afterBookmark.delete();
              await context.sync();
            }
          }
        }
      }

      await context.sync();
      toaster("AI tag application completed!", "success");

    } catch (err) {
      console.error("Error during tag application:", err);
    }
  });
}


function selectResponse(tagIndex, chatIndex) {
  // Handle the response selection logic here
  console.log(`Response selected for tagIndex ${tagIndex}, chatIndex ${chatIndex}`);
}


async function fetchGlossary() {
  if (!isTagUpdating) {

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

        const data = await fetchGlossaryTemplate(dataList?.ClientID, bodyText, jwt);

        layTerms = data.Data;

        if (data.Data.length > 0) {
          glossaryName = data.Data[0].GlossaryTemplate;
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
      layTerms.sort((a, b) => b.ClinicalTerm.length - a.ClinicalTerm.length);

      const processedTerms = new Set(); // Track added larger terms

      // Filter out smaller terms if they are included in a larger term
      const filteredTerms = layTerms.filter(term => {
        for (const biggerTerm of processedTerms) {
          if (typeof biggerTerm === 'string' && biggerTerm.includes(term.ClinicalTerm.toLowerCase())) {
            console.log(`Skipping "${term.ClinicalTerm}" because it's part of "${biggerTerm}"`);
            return false; // Exclude this smaller term
          }
        }
        processedTerms.add(term.ClinicalTerm.toLowerCase());
        return true;
      });

      filteredGlossaryTerm = filteredTerms;
      await removeMatchingContentControls();

      const foundRanges = new Map(); // Track words already processed

      const searchPromises = filteredGlossaryTerm.map(term => {
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
      isGlossaryActive = true;
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
  await checkGlossary();
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
        const searchPromises = layTerms.map(term => {
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
              font.highlightColor !== capturedFormatting['Background Color'] ||
              font.color !== capturedFormatting['Text Color'] ||
              font.bold !== capturedFormatting['Bold'] ||
              font.italic !== capturedFormatting['Italic'] ||
              font.size !== capturedFormatting['Size'] ||
              font.underline !== capturedFormatting['Underline'] ||
              font.name !== capturedFormatting['Font Name']
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

    words.forEach(word => {
      layTerms.forEach(term => {
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
      heading.textContent = `${clinicalTerm} (${glossaryName})`;
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
          item.insertText(layTerm, 'replace');

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
        if (control.title && filteredGlossaryTerm.some(term => term.ClinicalTerm.toLowerCase() === control.title.toLowerCase())) {
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
      isGlossaryActive = false;
      document.getElementById('applyglossary').addEventListener('click', applyglossary);
    });
  } catch (error) {
    console.error("Error removing content controls:", error);
  }
}


async function displayMentions() {
  if (!isTagUpdating) {
    if (isGlossaryActive) {
      await removeMatchingContentControls();
    }


  }
}

export async function addGenAITags() {

  if (!isTagUpdating) {

    if (isGlossaryActive) {
      await removeMatchingContentControls();
    }

    let selectedClient = clientList.filter(item => item.ID === clientId);

    // Build Primary Source List
    let sourceTypeList = [
      ...new Map(
        dataList.SourceTypeList
          .filter(item => item.VectorID > 0)
          .map(item => [item.SourceTypeID, { Name: item.SourceType, ID: item.SourceTypeID }])
      ).values()
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

    let sponsorOptions = clientList.map(client => {
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
    document.getElementById('add-tag-body').innerHTML = addtagbody(sponsorOptions, sourceOptions);

    const promptTemplateElement = document.getElementById('add-prompt-template');
    setupPromptBuilderUI(promptTemplateElement, promptBuilderList);

    document.getElementById('tag-tab').addEventListener('click', () => switchToAddTag());
    document.getElementById('prompt-tab').addEventListener('click', () => switchToPromptBuilder());

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

    document.getElementById('cancel-btn-gen-ai').addEventListener('click', () => {
      if (!isPendingResponse) loadHomepage(availableKeys);
    });


    if (form && nameField && promptField && sponsorDropdownItems.length > 0 && sourceDropdownItems.length > 0) {

      const updateSponsorDropdownLabel = () => {
        if (availableForAllCheckbox.checked) {
          sponsorDropdownButton.textContent = clientList.map(x => x.Name).join(", ");
        } else {
          const selectedNames = Array.from(sponsorDropdownItems)
            .filter(cb => cb.checked && cb.id !== 'sponsorSelectAll')
            .map(cb => cb.parentElement.textContent.trim());

          sponsorDropdownButton.textContent = selectedNames.length
            ? selectedNames.join(", ")
            : "Select Sponsors";
        }
      };

      const updateSourceDropdownLabel = () => {
        const selectedNames = Array.from(sourceDropdownItems)
          .filter(cb => cb.checked && cb.id !== 'sourceSelectAll')
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
        const selectedPrimarySources = Array.from(sourceDropdownItems)
          .filter(cb => cb.checked && cb.id !== 'sourceSelectAll')
          .map(cb => cb.value);

        if (!selectedPrimarySources.length) {
          document.getElementById("primarySourceError").style.display = "block";
          valid = false;
        } else {
          document.getElementById("primarySourceError").style.display = "none";
        }

        if (!valid) return;

        const selectedSponsors = Array.from(sponsorDropdownItems)
          .filter(cb => cb.checked && cb.id !== 'sponsorSelectAll')
          .map(cb => clientList.find(c => c.ID == cb.value));

        const isAvailableForAll = availableForAllCheckbox.checked;
        const isSaveGlobally = saveGloballyCheckbox.checked;
        const aigroup = dataList.Group.find(el => el.DisplayName === 'AIGroup');

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
          DocumentTypeID: dataList.DocumentTypeID,
          ReportHeadID: dataList.ID,

          // MULTI SELECT SOURCE TYPE
          SourceTypeID: selectedPrimarySources.join(","),

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
        if (!isPendingResponse) {
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
        if (!isPendingResponse) {
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
            const isChecked = checkbox.checked;
            sponsorDropdownItems.forEach(cb => {
              if (!cb.disabled) cb.checked = isChecked;
            });
          }

          updateSponsorDropdownLabel();
        });
      });

      document.querySelectorAll('.source-dropdown-item').forEach(item => {
        item.addEventListener('click', function (e) {
          e.stopPropagation();
          const checkbox = this.querySelector('.source-dropdown-item .form-check-input');
          if (!checkbox) return;

          if (checkbox.id === 'sourceSelectAll') {
            const isChecked = checkbox.checked;
            sourceDropdownItems.forEach(cb => {
              cb.checked = isChecked;
            });
          }

          updateSourceDropdownLabel();
          const selectedCount = Array.from(sourceDropdownItems)
            .filter(cb => cb.checked).length;

          if (selectedCount === 0) {
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
          if (this.classList.contains('is-invalid') && this.value.trim()) {
            this.classList.remove('is-invalid');
          }
        });
      });

    } else {
      console.error("Required elements missing.");
    }
  }
}


export async function customizeTable(type: string) {
  const container = document.getElementById("confirmation-popup");
  if (!container) return;

  const customStyleName = localStorage.getItem("CustomStyle") || "";
  const defaultStyle = localStorage.getItem("DefaultStyle") || tableStyle;
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
      styleObj = customTableStyle.find(s => s.Name === dropdown.value);
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
              (cell as HTMLTableCellElement).style.cssText = styleObj.sideHeader
                ? styleObj.tableClass! + "font-weight:bold;"
                : styleObj.tableClass!;
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
        const styleObj = customTableStyle.find(s => s.Name === dropdown.value);
        colorPallete.Header = styleObj.HeaderColor;
        colorPallete.Primary = styleObj.PrimaryColor;
        colorPallete.Secondary = styleObj.SecondaryColor;
        colorPallete.Customize = true;
        localStorage.setItem("CustomStyle", styleObj.Name);

        tableStyle = styleObj.BaseStyle; // stores full object as 
      } else {
        colorPallete.Customize = false;
        tableStyle = dropdown.value; // normal style string
        localStorage.setItem("DefaultStyle", tableStyle);

      }

      localStorage.setItem("colorPallete", JSON.stringify(colorPallete));
      localStorage.setItem("tableStyle", tableStyle);

      container.innerHTML = "";
    });
  }
}

async function createTextGenTag(payload) {
  try {
    const iconelement = document.getElementById(`text-gen-save`);
    const cancelBtnGenAi = document.getElementById('cancel-btn-gen-ai');


    (cancelBtnGenAi as HTMLButtonElement).disabled = true;
    iconelement.innerHTML = `<i class="fa fa-spinner fa-spin text-white me-2"></i>Save`;
    (iconelement as HTMLButtonElement).disabled = true;
    isPendingResponse = true;

    const data = await addGroupKey(payload, jwt);
    isPendingResponse = false;

    if (data['Data'] && data['Status']) {
      fetchDocument('AIpanel');
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
  const filterMentions = (query) => {
    // Assuming availableKeys is an array of objects with DisplayName and EditorValue properties
    const filtered = availableKeys.filter(item => item.AIFlag === 0).filter(item =>
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
        const editorValue = e.target.getAttribute('data-editor-value');
        selectMention(editorValue);
        mentionDropdown.style.display = 'none';  // Hide the dropdown after selection
      }
    });

    // Function to insert the selected mention into the prompt field
    const selectMention = (editorValue) => {
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
    document.addEventListener('click', (e) => {
      if (!mentionDropdown.contains(e.target) && e.target !== promptField) {
        mentionDropdown.style.display = 'none';
      }
    });
  }
}

export async function replaceMention(word: any, type: any) {
  return Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      await context.sync();

      if (!selection) {
        throw new Error('Selection is invalid or not found.');
      }

      if (type === 'TABLE') {
        const parser = new DOMParser();
        const doc = parser.parseFromString(word.EditorValue, 'text/html');

        const bodyNodes = Array.from(doc.body.childNodes);

        for (const node of bodyNodes) {
          if (node.nodeType === Node.TEXT_NODE) {
            let textContent = node.textContent?.trim();
            if (textContent) {
              textContent = textContent.replace(/\n- /g, "\n• ");

              textContent.split('\n').forEach(line => {
                if (line.trim()) {
                  insertLineWithHeadingStyle(selection, line);
                }
              });
            }
          } else if (node.nodeType === Node.ELEMENT_NODE) {
            const element = node as HTMLElement;

            if (element.tagName.toLowerCase() === 'table') {
              const rows = Array.from(element.querySelectorAll('tr'));

              if (rows.length === 0) {
                selection.insertParagraph("[Empty Table]", Word.InsertLocation.before);
                continue;
              }

              const maxCols = Math.max(...rows.map(row => {
                return Array.from(row.querySelectorAll('td, th')).reduce((sum, cell) => {
                  return sum + (parseInt(cell.getAttribute('colspan') || '1', 10));
                }, 0);
              }));

              const paragraph = selection.insertParagraph("", Word.InsertLocation.before);
              await context.sync();

              const table = paragraph.insertTable(rows.length, maxCols, Word.InsertLocation.after);
              table.style = tableStyle;  // Apply built-in Word table style

              await context.sync();
              if (colorPallete.Customize) {
                await colorTable(table, rows, context);
              }

              const rowspanTracker: number[] = new Array(maxCols).fill(0);

              rows.forEach((row, rowIndex) => {
                const cells = Array.from(row.querySelectorAll('td, th'));
                let cellIndex = 0;

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
                  // if (rowIndex === 0) {
                  //   const cell = table.getCell(rowIndex, cellIndex);
                  //   const paragraph = cell.body.paragraphs.getFirst();
                  //   paragraph.font.bold = true;
                  //   paragraph.font.highlightColor = "lightGray";  // This works!
                  // }
                  table.getCell(rowIndex, cellIndex).value = cellText;

                  for (let i = 1; i < colspan; i++) {
                    if (cellIndex + i < maxCols) {
                      table.getCell(rowIndex, cellIndex + i).value = "";
                    }
                  }

                  if (rowspan > 1) {
                    for (let i = 0; i < colspan; i++) {
                      if (cellIndex + i < maxCols) {
                        rowspanTracker[cellIndex + i] = rowspan - 1;
                      }
                    }
                  }

                  cellIndex += colspan;
                });
              });
            } else {
              let elementText = element.innerText.trim();
              if (elementText) {
                elementText = elementText.replace(/\n- /g, "\n• ");
                elementText.split('\n').forEach(line => {
                  if (line.trim()) {
                    insertLineWithHeadingStyle(selection, line);
                  }
                });
              }
            }
          }
        }
      }

      else {
        if (word.EditorValue === '' || word.IsApplied) {
          selection.insertParagraph(`#${word.DisplayName}#`, Word.InsertLocation.before);
        } else {
          let content = removeQuotes(word.EditorValue);
          let lines = content.split(/\r?\n/); // Handle both \r\n and \n

          lines.forEach(line => {
            selection.insertParagraph(line, Word.InsertLocation.before);
          });
        }
      }

      await context.sync();
    } catch (error) {
      console.error('Detailed error:', error);
    }
  });
}


function insertLineWithHeadingStyle(range: Word.Range, line: string) {
  let style = "Normal";
  let text = line;

  if (line.startsWith('###### ')) {
    style = "Heading 6";
    text = line.substring(7).trim();
  } else if (line.startsWith('##### ')) {
    style = "Heading 5";
    text = line.substring(6).trim();
  } else if (line.startsWith('#### ')) {
    style = "Heading 4";
    text = line.substring(5).trim();
  } else if (line.startsWith('### ')) {
    style = "Heading 3";
    text = line.substring(4).trim();
  } else if (line.startsWith('## ')) {
    style = "Heading 2";
    text = line.substring(3).trim();
  } else if (line.startsWith('# ')) {
    style = "Heading 1";
    text = line.substring(2).trim();
  }

  const paragraph = range.insertParagraph(text, Word.InsertLocation.before);
  paragraph.style = style;
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

function showAddTagError(message) {
  const errorDiv = document.getElementById('submition-error');
  errorDiv.style.display = 'block';
  errorDiv.textContent = message;
}

function transformDocumentName(value: string): string {
  if (!value || value.trim() === '') {
    return value; // Return the input value unchanged
  }

  const parts = value.split('_');
  if (parts.length <= 1) {
    return value; // Return the input value unchanged if no underscores are present
  }

  return parts.slice(1).join('_').replace(/%20/g, ' ').replace(/%25/g, '%');
}



export function createMultiSelectDropdown(tag) {
  const isDark = theme === 'Dark';
  const btnClass = isDark ? 'btn-dark text-light border-0' : 'btn-light text-dark border';
  const dropdownMenuClass = isDark ? 'bg-dark text-light border-light' : 'bg-white text-dark border';
  const itemClass = isDark ? 'bg-dark text-light' : 'bg-white text-dark';

  // Group sources by SourceType
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
                      <input class="form-check-input source-checkbox" type="checkbox" value="${source.SourceName}" id="source-${groupIndex}-${index}">
                      <label class="form-check-label w-100 text-prewrap" for="source-${groupIndex}-${index}">${source.SourceName}</label>
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
    const receivedEntry = sourceList.filter(source => selectedSources.includes(source.SourceName));
    tag.TempSourceValue = receivedEntry.map((item) => {
      return item.VectorID ? String(item.VectorID) : item.SourceValue;
    });

    tag.SourceName = receivedEntry.map((item) => {
      return item.SourceName;
    });

    tag.SourceValueID = receivedEntry.map((item) => {
      return String(item.VectorID);
    });

    tag.SourceValue = receivedEntry
      .map(source => source.SourceValue);
    accordionBody.innerHTML = chatfooter(tag);
    initializeAIHistoryEvents(tag, jwt, availableKeys);
  });

  // Cancel
  document.getElementById(`cancel-src-btn`).addEventListener("click", function () {
    accordionBody.innerHTML = chatfooter(tag);
    initializeAIHistoryEvents(tag, jwt, availableKeys);
  });
}



async function loadPromptTemplates() {
  try {
    const data = await getAllPromptTemplates(jwt);
    if (data.Status && data.Data) {
      promptBuilderList = data.Data;
    }
    // Do something with the data
  } catch (error) {
    console.error('Error fetching prompt templates:', error);
  }
}


async function logBookmarksInSelection() {
  return Word.run(async (context) => {
    let range = context.document.getSelection();
    await context.sync(); // Ensure selection is ready


    // Get bookmarks in the selection
    let bookmarks = range.getBookmarks(); // Returns ClientResult<string[]>

    await context.sync(); // Ensure bookmarks are retrieved
    if (bookmarks.value.length > 1) {
      selectedNames = []
      const badgeWrapper = document.getElementById('tags-in-selected-text');
      if (badgeWrapper) {
        badgeWrapper.classList.remove('d-none');
        badgeWrapper.classList.add('d-block');
      }
      bookmarks.value.forEach((bookmarkName) => {
        let processedName = bookmarkName.split("_Split_")[0];
        processedName = processedName.replace(/_/g, " ");
        selectedNames.push(processedName)
        const container = document.getElementById('tags-in-selected-text');
        if (container) {
          renderSelectedTags(selectedNames, availableKeys)// Trigger function when selection changes
        }
      });
    } else if (bookmarks.value.length === 1) {
      const badgeWrapper = document.getElementById('tags-in-selected-text');
      if (badgeWrapper) {
        badgeWrapper.classList.remove('d-none');
        badgeWrapper.classList.add('d-block');
      }
      bookmarks.value.forEach((bookmarkName) => {
        let processedName = bookmarkName.split("_Split_")[0];
        processedName = processedName.replace(/_/g, " ");
        let aiTag;
        if (/^ID\d+$/i.test(processedName)) {
          aiTag = availableKeys.find(
            mention => mention.AIFlag === 1 && `id${mention.ID}`.toLowerCase() === processedName.toLowerCase()
          );
        } else {
          aiTag = availableKeys.find(
            mention => mention.AIFlag === 1 && mention.DisplayName.toLowerCase() === processedName.toLowerCase()
          );
        }

        const container = document.getElementById('tags-in-selected-text');
        if (container && aiTag) {
          selectMatchingBookmarkFromSelection(processedName);
          selectedNames = [processedName];

          const appBody = document.getElementById('app-body');
          appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';

          generateCheckboxHistory(aiTag).then(html => {
            appBody.innerHTML = html;
          });
        }
      });
    } else {
      const badgeWrapper = document.getElementById('tags-in-selected-text');
      if (badgeWrapper) {
        badgeWrapper.classList.remove('d-block');
        badgeWrapper.classList.add('d-none');
      }
    }
  });
}
