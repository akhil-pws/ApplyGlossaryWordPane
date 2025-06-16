/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import { dataUrl, storeUrl, versionLink } from "./data";
import { generateCheckboxHistory, initializeAIHistoryEvents, loadHomepage, setupPromptBuilderUI } from "./components/home";
import { applyThemeClasses, chatfooter, copyText, renderSelectedTags, swicthThemeIcon, switchToAddTag, switchToPromptBuilder, updateEditorFinalTable } from "./functions";
import { addtagbody, logoheader, navTabs, promptbuilderbody } from "./components/bodyelements";
import { addAiHistory, addGroupKey, fetchGlossaryTemplate, getAiHistory, getAllClients, getAllPromptTemplates, getReportById, loginUser, updateAiHistory, updateGroupKey } from "./api";
export let jwt = '';
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
let sourceList;
let filteredGlossaryTerm;
export let selectedNames = [];
export let isPendingResponse = false;
export let theme = 'Light';


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
  console.log(sessionToken)
  if (sessionToken) {
    jwt = sessionToken;
    window.location.hash = '#/dashboard';
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
          sessionStorage.setItem('token', jwt)
          sessionStorage.setItem('userId', data.Data.ID);
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

async function fetchDocument(action) {
  try {

    const data = await getReportById(documentID, jwt);
    document.getElementById('app-body').innerHTML = ``
    document.getElementById('logo-header').innerHTML = logoheader(storedUrl);

    dataList = data['Data'];
    sourceList = dataList.SourceTypeList.filter(
      (item) => item.SourceValue !== ''
        && item.AIFlag === 1
    ) // Filter items with an extension
      .map((item) => ({
        ...item, // Spread the existing properties
        SourceName: transformDocumentName(item.SourceValue)
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
      SourceValue: tag.SourceValue ? tag.SourceValue : []
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

        console.log(`Found ${searchResults.items.length} instances of #${tag.DisplayName}#`);
        const tableInsertPositions: { range: Word.Range, tag: any }[] = [];

        for (const item of searchResults.items) {
          if (tag.EditorValue !== "" && !tag.IsApplied) {
            const cleanDisplayName = tag.ID;
            const uniqueStr = new Date().getTime();
            const bookmarkName = `ID${cleanDisplayName}_Split_${uniqueStr}`;

            const startMarker = item.insertParagraph("[[BOOKMARK_START]]", Word.InsertLocation.before);
            await context.sync();

            if (tag.ComponentKeyDataType === 'TABLE') {
              const range = item.getRange();
              tableInsertPositions.push({ range, tag });
              range.delete();
            } else {
              item.insertText(tag.EditorValue, Word.InsertLocation.replace);
            }


            for (const { range, tag } of tableInsertPositions) {
              const parser = new DOMParser();
              const doc = parser.parseFromString(tag.EditorValue, 'text/html');
              const bodyNodes = Array.from(doc.body.childNodes);

              for (const node of bodyNodes) {
                if (node.nodeType === Node.TEXT_NODE) {
                  const textContent = node.textContent?.trim();
                  if (textContent) {
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
                    table.style = "Grid Table 4 - Accent 1";
                    await context.sync();

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
                    const elementText = element.innerText.trim();
                    if (elementText) {
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
            }
            const endMarker = item.insertParagraph("[[BOOKMARK_END]]", Word.InsertLocation.after);
            await context.sync();

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
              start.delete();
              end.delete();
              // Move the cursor to this paragraph (now it's outside the bookmark)
              afterBookmark.select();

              afterBookmark.delete();
              await context.sync();


            }



          }
        }

      }

      await context.sync();
    } catch (err) {
      console.error("Error during tag application:", err);
    }
  });
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

    let selectedClient = clientList.filter((item) => item.ID === clientId);

    let sponsorOptions = clientList.map(client => {
      const isSelectedClient = selectedClient.some(selected => selected.ID === client.ID);
      return ` 
        <li class="dropdown-item p-2" style="cursor: pointer;">
          <div class="form-check">
            <input class="form-check-input" type="checkbox" value="${client.ID}" id="sponsor${client.ID}" ${isSelectedClient ? 'checked disabled' : ''}>
            <label class="form-check-label text-prewrap" for="sponsor${client.ID}">${client.Name}</label>
          </div>
        </li>
      `;
    }).join('');

    document.getElementById('app-body').innerHTML = navTabs;
    // Add modal HTML to the DOM
    document.getElementById('add-tag-body').innerHTML = addtagbody(sponsorOptions);
    const promptTemplateElement = document.getElementById('add-prompt-template')
    setupPromptBuilderUI(promptTemplateElement, promptBuilderList)


    document.getElementById('tag-tab').addEventListener('click', () => {
      switchToAddTag()
    });


    document.getElementById('prompt-tab').addEventListener('click', () => {
      switchToPromptBuilder()
    });

    //prompt starting
    mentionDropdownFn('prompt', 'mention-dropdown', 'add');
    //prompt end
    const form = document.getElementById('genai-form');
    const promptField = document.getElementById('prompt');

    const nameField = document.getElementById('name');
    const descriptionField = document.getElementById('description');
    const saveGloballyCheckbox = document.getElementById('saveGlobally');
    const availableForAllCheckbox = document.getElementById('isAvailableForAll');
    const sponsorDropdownButton = document.getElementById('sponsorDropdown');
    const sponsorDropdownItems = document.querySelectorAll('.dropdown-item .form-check-input');

    document.getElementById('cancel-btn-gen-ai').addEventListener('click', () => {
      if (!isPendingResponse) {
        loadHomepage(availableKeys)
      }
    });

    // Check if elements exist
    if (form && nameField && promptField && sponsorDropdownItems.length > 0) {
      const updateDropdownLabel = () => {
        if (availableForAllCheckbox.checked) {
          sponsorDropdownButton.textContent = clientList.map(client => client.Name).join(", ");
        } else {
          const selectedOptions = Array.from(sponsorDropdownItems)
            .filter(cb => cb.checked && cb.id !== 'selectAll')
            .map(cb => cb.parentElement.textContent.trim());
          sponsorDropdownButton.textContent = selectedOptions.length ? selectedOptions.join(", ") : "Select Sponsors";
        }
      };
      // Form validation logic on submit
      form.addEventListener('submit', async function (e) {
        e.preventDefault();

        // Reset previous validation errors
        form.querySelectorAll('.is-invalid').forEach(input => input.classList.remove('is-invalid'));

        let valid = true;

        if (!(nameField as HTMLInputElement).value.trim()) {
          nameField.classList.add('is-invalid');
          valid = false;
        }

        if (!(promptField as HTMLInputElement).value.trim()) {
          promptField.classList.add('is-invalid');
          valid = false;
        }

        if (valid) {
          // Prepare object to pass to createTextGenTag
          const selectedSponsors = Array.from(sponsorDropdownItems)
            .filter(cb => cb.checked && cb.id !== 'selectAll')
            .map(cb => {
              const client = clientList.find(client => client.ID == cb.value);
              return client; // Collect the entire client object
            });

          const isAvailableForAll = availableForAllCheckbox.checked;
          const isSaveGlobally = saveGloballyCheckbox.checked;
          const aigroup = dataList.Group.find(element => element.DisplayName === 'AIGroup');
          const formData = {
            DisplayName: nameField.value.trim(),
            Prompt: promptField.value.trim(),
            Description: descriptionField.value.trim(),
            GroupKeyClient: selectedSponsors, // Array of selected sponsor objects
            AllClient: isAvailableForAll ? 1 : 0,
            SaveGlobally: isSaveGlobally,
            UserDefined: '1',
            ComponentKeyDataTypeID: '1',
            ComponentKeyDataAccessID: '3',
            AIFlag: 1,
            DocumentTypeID: dataList.DocumentTypeID,
            ReportHeadID: dataList.ID,
            SourceTypeID: '',
            ReportHeadGroupID: aigroup.ID,
            ReportHeadSourceID: ''
          };

          await createTextGenTag(formData);
        }
      });


      const checkAndDisableSponsors = () => {
        sponsorDropdownItems.forEach(checkbox => {
          if (!checkbox.disabled) {
            checkbox.checked = true;
            checkbox.disabled = true;
          }
        });
        updateDropdownLabel();
      };

      // Function to enable sponsors without unchecking them
      const enableSponsors = () => {
        sponsorDropdownItems.forEach(checkbox => {
          const isSelectedClient = selectedClient.some(selected => selected.ID === parseInt(checkbox.value));
          if (!isSelectedClient) {
            checkbox.disabled = false;
          }
        });
        updateDropdownLabel();
      };

      // Event listener for "Save Globally" checkbox


      // Event listener for "Available to All Sponsors" checkbox

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
            sponsorDropdownItems.forEach(checkbox => {
              if (!checkbox.disabled) {
                checkbox.checked = false;
                checkbox.disabled = false;
              }
            });
            updateDropdownLabel();
          }
        }
      });

      // Event listener for "Available to All Sponsors" checkbox
      availableForAllCheckbox.addEventListener('change', function () {
        if (!isPendingResponse) {

          if (this.checked) {
            checkAndDisableSponsors();
          } else {
            enableSponsors();
          }
        }
      });

      // Add event listener to prevent dropdown close on item selection
      document.querySelectorAll('.dropdown-item').forEach(item => {
        item.addEventListener('click', function (event) {
          {
            event.stopPropagation(); // Prevent dropdown from closing
            const checkbox = this.querySelector('.form-check-input');
            if (checkbox) {


              if (checkbox.id === 'selectAll') {
                const isChecked = checkbox.checked;
                sponsorDropdownItems.forEach(cb => {
                  if (!cb.disabled) cb.checked = isChecked;
                });
              }
              updateDropdownLabel();
            }
          }
        });
      });

      // Initial label update
      updateDropdownLabel();


      // Clear validation errors when user types
      [nameField, promptField].forEach(field => {
        field.addEventListener('input', function () {
          if (this.classList.contains('is-invalid') && this.value.trim()) {
            this.classList.remove('is-invalid');
          }
          if (nameField) {
            const errorDiv = document.getElementById('submition-error');
            errorDiv.style.display = 'none';
          }
        });
      });
    } else {
      console.error('Required elements are missing or not rendered yet.');
    }
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
    } else {
      (cancelBtnGenAi as HTMLButtonElement).disabled = false;
      (iconelement as HTMLButtonElement).disabled = false;
      iconelement.innerHTML = `<i class="fa fa-check-circle me-2"></i>Save`;
      showAddTagError(data['Data']);
    }

  } catch (error) {
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
            const textContent = node.textContent?.trim();
            if (textContent) {
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
              table.style = "Grid Table 4 - Accent 1";  // Apply built-in Word table style

              await context.sync();

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
              const elementText = element.innerText.trim();
              if (elementText) {
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
  const cancelBtnClass = isDark ? 'btn-danger bg-danger-clr text-light' : 'btn-danger bg-danger-clr text-white';
  const saveBtnClass = isDark ? 'btn-success bg-success-clr text-light' : 'btn-success bg-success-clr text-white';

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
        <ul class="dropdown-menu ${dropdownMenuClass} w-100 p-2" aria-labelledby="sourceDropdown" style="box-shadow: 0 4px 8px rgba(0,0,0,0.1); z-index: 10000;">
          <li class="dropdown-item p-2 ${itemClass}" style="cursor: pointer;" data-checkbox-id="selectAll">
            <div class="form-check">
              <input class="form-check-input" type="checkbox" value="selectAll" id="selectAll">
              <label class="form-check-label w-100" for="selectAll">Select All</label>
            </div>
          </li>
          ${sourceList
      .map(
        (source, index) => `
              <li class="dropdown-item p-2 ${itemClass}" style="cursor: pointer;" data-checkbox-id="source-${index}">
                <div class="form-check">
                  <input class="form-check-input source-checkbox" type="checkbox" value="${source.SourceName}" id="source-${index}">
                  <label class="form-check-label w-100 text-prewrap" for="source-${index}">${source.SourceName}</label>
                </div>
              </li>
            `
      )
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
  const individualCheckboxes = document.querySelectorAll(`.source-checkbox`);
  const sourceDropdownLabel = document.getElementById(`sourceDropdownLabel`);

  function updateLabel() {
    const selectedSourceNames = selectedSources;
    if (selectedSourceNames.length > 0) {
      sourceDropdownLabel.innerText = selectedSourceNames.join(', ');
    } else {
      sourceDropdownLabel.innerText = ' ';
    }
  }

  selectAllCheckbox.addEventListener("change", function () {
    const checkboxes = document.querySelectorAll(`.source-checkbox`);
    checkboxes.forEach((checkbox) => {
      checkbox.checked = this.checked;
      if (checkbox.checked) {
        if (!selectedSources.includes(checkbox.value)) {
          selectedSources.push(checkbox.value);
        }
      } else {
        selectedSources = selectedSources.filter((source) => source !== checkbox.value);
      }
    });

    updateLabel();
  });

  const selectAllItem = document.querySelector(`.dropdown-item[data-checkbox-id="selectAll"]`);
  selectAllItem.addEventListener("click", function (event) {
    event.stopPropagation();
  });

  individualCheckboxes.forEach((checkbox) => {
    checkbox.addEventListener("change", function () {
      if (checkbox.checked) {
        if (!selectedSources.includes(checkbox.value)) {
          selectedSources.push(checkbox.value);
        }
      } else {
        selectedSources = selectedSources.filter((source) => source !== checkbox.value);
      }

      const allChecked = Array.from(individualCheckboxes).every((checkbox) => checkbox.checked);
      selectAllCheckbox.checked = allChecked;

      updateLabel();
    });

    const listItem = checkbox.closest("li");
    listItem.addEventListener("click", function (event) {
      event.stopPropagation();
    });
  });

  if (tag.Sources && tag.Sources.length > 0) {
    individualCheckboxes.forEach((checkbox) => {
      if (tag.Sources.includes(checkbox.value)) {
        checkbox.checked = true;
        selectedSources.push(checkbox.value);
      }
    });

    const allChecked = Array.from(individualCheckboxes).every((checkbox) => checkbox.checked);
    selectAllCheckbox.checked = allChecked;
    updateLabel();
  }

  document.getElementById(`ok-src-btn`).addEventListener("click", function () {
    tag.Sources = [...selectedSources];
    tag.SourceValue = sourceList
      .filter(source => selectedSources.includes(source.SourceName))
      .map(source => source.SourceValue);

    accordionBody.innerHTML = chatfooter(tag);
    initializeAIHistoryEvents(tag, jwt, availableKeys);
  });

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
    if (bookmarks.value.length > 0) {
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
    } else {
      const badgeWrapper = document.getElementById('tags-in-selected-text');
      if (badgeWrapper) {
        badgeWrapper.classList.remove('d-block');
        badgeWrapper.classList.add('d-none');
      }
    }
  });
}
