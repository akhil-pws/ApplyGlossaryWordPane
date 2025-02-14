/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { dataUrl, storeUrl, versionLink } from "./data";
let jwt = '';
let baseUrl = dataUrl
let storedUrl = storeUrl
let documentID = ''
let organizationName = ''
let aiTagList = [];
let initialised = true;
let availableKeys = [];
let glossaryName = ''
let isGlossaryActive: boolean = false;
let imageList = [];
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

interface FontProps {
  highlight?: string;
  color?: string;
  bold?: string;
  italic?: string;
  underline?: string;
  size?: string;
  family?:string;
}




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


  document.getElementById('app-body').innerHTML = `
    <div class="container mt-5">
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

  document.getElementById('login-form').addEventListener('submit', handleLogin);
}

async function handleLogin(event) {
  event.preventDefault();

  // Get the values from the form fields
  const organization = document.getElementById('organization').value;
  const username = document.getElementById('username').value;
  const password = document.getElementById('password').value;
  if (organization.toLowerCase().trim() === organizationName.toLocaleLowerCase().trim()) {
    document.getElementById('app-body').innerHTML = `
  <div id="button-container">

          <div class="loader" id="loader"></div>
          </div
`
    try {
      const response = await fetch(`${baseUrl}/api/user/login`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          ClientName: organization,
          Username: username,
          Password: password
        })
      });


      if (!response.ok) {
        showLoginError("An error occurred during login. Please try again.")
        throw new Error('Network response was not ok.');
      }

      const data = await response.json();
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


      // Handle successful login (e.g., navigate to the next page or show a success message)

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
    const response = await fetch(`${baseUrl}/api/report/id/${documentID}`, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${jwt}`
      }
    });

    if (!response.ok) {
      throw new Error('Network response was not ok.');
    }

    const data = await response.json();
    document.getElementById('app-body').innerHTML = ``
    document.getElementById('logo-header').innerHTML = `
        <img  id="main-logo" src="${storedUrl}/assets/logo.png" alt="" class="logo"> <i class="fa fa-sign-out me-5 c-pointer ngb-tooltip" aria-hidden="true" id="logout"><span class="tooltiptext">Logout</span></i>
`
    document.getElementById('header').innerHTML = `
    <div class="d-flex justify-content-around">
        <button class="btn btn-dark" id="mention">Insert</button>
        <button class="btn btn-dark" id="aitag">Refine</button>

        <!-- Dropdown for Formatting -->
        <div class="dropdown">
            <button class="btn btn-dark dropdown-toggle" type="button" id="formatDropdown" data-bs-toggle="dropdown" aria-expanded="false">
                Actions
            </button>
            <ul class="dropdown-menu" aria-labelledby="formatDropdown" style="z-index: 100000;">
                <li><button class="dropdown-item" id="selectFormat">Define Formatting</button></li>
                <li><button class="dropdown-item" id="glossary" disabled>Glossary</button></li>
                <li><button class="dropdown-item" id="removeFormatting" disabled>Remove Formatted Text</button></li>
            </ul>
        </div>
    </div>
`

    if (action === 'Init') {
      setActiveButton('mention');
      displayMentions();
    }else{
      setActiveButton('aitag');
    }
    document.getElementById('mention').addEventListener('click', () => {
      setActiveButton('mention');
      displayMentions();
    });

    document.getElementById('glossary').addEventListener('click', () => {
      setActiveButton('formatDropdown');
      fetchGlossary()
    });

    document.getElementById('aitag').addEventListener('click', () => {
      setActiveButton('aitag');
      displayAiTagList();
    });
    document.getElementById('selectFormat').addEventListener('click', () => {
      setActiveButton('formatDropdown');
      formatOptionsDisplay();
    });
    document.getElementById('removeFormatting').addEventListener('click', () => {
      setActiveButton('formatDropdown');
      removeOptionsConfirmation();
    });

    document.getElementById('logout').addEventListener('click', logout);

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
    fetchClients();
    if (action === 'AIpanel') {
      displayAiTagList();
    }

  } catch (error) {
    console.error('Error fetching glossary data:', error);
  }
}

function setActiveButton(buttonId) {
  const buttons = ['mention', 'aitag', 'selectFormat', 'removeFormatting', 'formatDropdown'];
  buttons.forEach(id => {
    const button = document.getElementById(id);
    if (button) {
      if (id === buttonId) {
        button.classList.add('active');
      } else {
        button.classList.remove('active');
      }
    }
  });
}

async function fetchClients() {
  try {
    const response = await fetch(`${baseUrl}/api/client/all/${userId}`, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${jwt}`
      }
    });

    if (!response.ok) {
      throw new Error('Network response was not ok.');
    }



    const data = await response.json();
    clientList = data['Data'];
  } catch (error) {
  }
}



async function formatOptionsDisplay() {
  if (!isTagUpdating) { // Check if isTagUpdating is false
    if (isGlossaryActive) {
      await removeMatchingContentControls();
    }
    const htmlBody = `
      <div class="container mt-3">
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
    glossaryBtn.disabled = true;
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
        removeFormatBtn.disabled = true;
      } else {
        const removeFormatBtn = document.getElementById('removeFormatting') as HTMLButtonElement;
        removeFormatBtn.disabled = false;
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
        glossaryBtn.disabled = true;
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
  glossaryBtn.disabled = false;
  const CaptureBtn = document.getElementById('capture-format-btn') as HTMLButtonElement;
  CaptureBtn.disabled = true;

  const removeFormatBtn = document.getElementById('removeFormatting') as HTMLButtonElement;
  removeFormatBtn.disabled = true;

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
        removeFormatBtn.disabled = true;
      } else {
        const removeFormatBtn = document.getElementById('removeFormatting') as HTMLButtonElement;
        removeFormatBtn.disabled = false;
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
      <div class="container mt-3">
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
            <div class="d-flex justify-content-end mt-2">
              <button id="change-ft-btn" class="btn btn-danger bg-danger-clr px-3 me-2"><i class="fa fa-reply me-2"></i>Cancel</button>
              <button id="clear-ft-btn" class="btn btn-success bg-success-clr px-3"><i class="fa fa-check-circle me-2"></i>Yes</button>

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
      glossaryBtn.disabled = false;
      formatOptionsDisplay()
    });
  } catch (error) {
    console.error("Error removing formatted text:", error);
  }
}


async function fetchAIHistory(tag) {
  try {
    const response = await fetch(`${baseUrl}/api/report/ai-history/${tag.ID}`, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${jwt}`
      }
    });

    if (!response.ok) {
      throw new Error('Network response was not ok.');
    }

    const data = await response.json();
    tag.ReportHeadAIHistoryList = data['Data'] || [];
    tag.FilteredReportHeadAIHistoryList = [];
    tag.ReportHeadAIHistoryList.forEach((historyList, i) => {
      historyList.Response = removeQuotes(historyList.Response);
      tag.FilteredReportHeadAIHistoryList.unshift(historyList);

    });
    return tag.FilteredReportHeadAIHistoryList;

  } catch (error) {
    console.error('Error fetching AI history:', error);
    return [];
  }
}


async function generateRadioButtons(tag: any, index: number): Promise<string> {
  if (!tag.FilteredReportHeadAIHistoryList || tag.FilteredReportHeadAIHistoryList.length === 0) {
    await fetchAIHistory(tag);
  }

  if (tag.FilteredReportHeadAIHistoryList.length > 0) {
    // Generate the HTML
    const html = tag.FilteredReportHeadAIHistoryList.map((chat: any, j: number) =>
      `<div class="row chatbox m-0 p-0">
        <div class="col-md-12 mt-2 p-2">
          <span class="ms-3">
            <i class="fa fa-copy text-secondary c-pointer" title="Copy Response" id="copyPrompt-${index}-${j}"></i>
          </span>
          <span class="float-end w-75 me-3">
            <div class="form-control h-34 d-flex align-items-center dynamic-height user">
              ${chat.Prompt}
            </div>
          </span>
        </div>
        <div class="col-md-12 mt-2 p-2 d-flex align-items-center">
          <span class="radio-select">
            <input class="form-check-input c-pointer" type="radio" name="flexRadioDefault-${index}"
              id="flexRadioDefault1-${index}-${j}" ${chat.Selected === 1 ? 'checked' : ''}>
          </span>
          <span class="ms-2 w-75">
            <div class="form-control h-34 d-flex align-items-center dynamic-height ai-reply ${chat.Selected === 1 ? 'ai-selected-reply' : 'bg-light'}" id='selected-response-${index}${j}'>
              ${chat.Response}
            </div>
          </span>
          <span class="ms-2">
            <i class="fa fa-copy text-secondary c-pointer" title="Copy Response" id="copyResponse-${index}-${j}"></i>
          </span>
        </div>


      </div>`
    ).join('');

    // Attach event listeners after the HTML is inserted
    setTimeout(() => {
      tag.FilteredReportHeadAIHistoryList.forEach((chat: any, j: number) => {
        document.getElementById(`copyPrompt-${index}-${j}`)?.addEventListener('click', () => copyText(chat.Prompt));
        document.getElementById(`copyResponse-${index}-${j}`)?.addEventListener('click', () => copyText(chat.Response));
        document.getElementById(`flexRadioDefault1-${index}-${j}`)?.addEventListener('change', () => onRadioChange(tag, index, j));

      });
    }, 0);

    return html;
  } else {
    return '<div>No AI history available.</div>';
  }
}




function accordionContent(headerId, collapseId, tag, radioButtonsHTML, i) {
  const textColorClass = tag.IsApplied ? 'text-secondary' : '';
  const tooltipButton = tag.Sources
    ? `  <span class="tooltiptext">${tag.Sources}</span>`
    : '';
  const body = `
    <div class="accordion-item">
      <h2 class="accordion-header" id="${headerId}">
     <button class="accordion-button collapsed ${textColorClass}"
        type="button"
        data-bs-toggle="collapse"
        data-bs-target="#${collapseId}"
        aria-expanded="false"
        aria-controls="${collapseId}">
    <span id="tagname-${i}">${tag.DisplayName}</span>
</button>

      </h2>
      <div id="${collapseId}"
           class="accordion-collapse collapse"
           aria-labelledby="${headerId}"
           data-bs-parent="#accordionExample">
        <div class="accordion-body p-0" id="accordion-body-${i}">
          <div class="chatbox" id="selected-response-parent-${i}">
            ${radioButtonsHTML}
          </div>
          <div class="form-check form-switch chatbox m-0">
            <label class="form-check-label pb-3" for="doNotApply-${i}">
              <span class="fs-12">Do not apply</span>
            </label>
            <input class="form-check-input"
                   type="checkbox"
                   id="doNotApply-${i}"
                   ${tag.IsApplied ? 'checked' : ''}>
          </div>
          <div class="d-flex align-items-end justify-content-end chatbox p-2">
            <textarea class="form-control"
                      rows="5"
                      id="chatbox-${i}"
                      placeholder="Type here"></textarea>
                <div id="mention-dropdown-${i}" class="dropdown-menu"></div>
            <div class="d-flex flex-column align-self-end me-3">
            <button class="btn btn-secondary text-light ms-2 mb-2 ngb-tooltip" id="insert-tag-${i}">
            <span class="tooltiptext">Insert</span>
                <i class="fa fa-plus text-light c-pointer" ></i>
                </button>
            <button
                    class="btn btn-secondary ms-2 mb-2 text-white ngb-tooltip"
                    id="changeSource-${i}">
                    ${tooltipButton}
              <i class="fa fa-file-lines text-white"></i>
            </button>
            <button type="submit"
                    class="btn btn-primary bg-primary-clr ms-2 text-white"
                    id="sendPrompt-${i}">
              <i class="fa fa-paper-plane text-white"></i>
            </button>
            </div>
          </div>
        </div>
      </div>
    </div>`;
  return body;
}

async function onDoNotApplyChange(event, index, tag: any) {
  tag.IsApplied = event.target.checked;
  let sourceListBtn = document.getElementById(`changeSource-${index}`) as HTMLButtonElement;
  sourceListBtn.disabled = true;
  const isChecked = event.target.checked;
  const tagname = document.getElementById(`tagname-${index}`);
  const dnaBtn = document.getElementById(`doNotApply-${index}`) as HTMLInputElement;

  try {
    dnaBtn.disabled = true
    const response = await fetch(`${baseUrl}/api/report/head/groupkey`, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${jwt}`
      },
      body: JSON.stringify(tag)
    });

    if (!response.ok) {
      dnaBtn.disabled = false
      throw new Error('Network response was not ok.');
    }

    const data = await response.json();
    if (data['Data'] && data['Status'] === true) {
      sourceListBtn.disabled = false;
      dnaBtn.disabled = false
    } else {
      sourceListBtn.disabled = false;
      dnaBtn.disabled = false
    }

  } catch (error) {
    dnaBtn.disabled = false
    sourceListBtn.disabled = false;
    console.error('Error updating do not apply:', error);
  }

  if (tagname) {
    const match = availableKeys.find(item => tag.DisplayName === item.DisplayName);
    if (isChecked) {
      if (match) match.IsApplied = true;
      tagname.classList.add('text-secondary');
    } else {
      if (match) match.IsApplied = false;
      tagname.classList.remove('text-secondary');
    }
  }

}

async function sendPrompt(tag, prompt, index) {
  if (prompt !== '' && !isTagUpdating) {
    let sourceListBtn = document.getElementById(`changeSource-${index}`) as HTMLButtonElement;
    sourceListBtn.disabled = true;
    isTagUpdating = true;

    const iconelement = document.getElementById(`sendPrompt-${index}`);
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
      if (data['Data'] && data['Data'] !== 'false') {
        tag.ReportHeadAIHistoryList = JSON.parse(JSON.stringify(data['Data']));
        tag.FilteredReportHeadAIHistoryList = [];

        tag.ReportHeadAIHistoryList.forEach((historyList) => {
          historyList.Response = removeQuotes(historyList.Response);
          tag.FilteredReportHeadAIHistoryList.unshift(historyList);
        });

        // Update only the inner content of the accordion body
        const innerContainer = document.getElementById(`selected-response-parent-${index}`);
        if (innerContainer) {
          const radioButtonsHTML = await generateRadioButtons(tag, index);
          innerContainer.innerHTML = radioButtonsHTML;
        }

        // Reapply event listeners for the new buttons and radio options
        tag.FilteredReportHeadAIHistoryList.forEach((chat, j) => {
          document.getElementById(`copyPrompt-${index}-${j}`)?.addEventListener('click', () => copyText(chat.Prompt));
          document.getElementById(`copyResponse-${index}-${j}`)?.addEventListener('click', () => copyText(chat.Response));
          document.getElementById(`flexRadioDefault1-${index}-${j}`)?.addEventListener('change', () => onRadioChange(tag, index, j));
        });

        iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`;
        document.getElementById(`chatbox-${index}`).value = '';

        isTagUpdating = false;
        sourceListBtn.disabled = false;
      } else {
        iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`;
        isTagUpdating = false;
        sourceListBtn.disabled = false;
      }
    } catch (error) {
      iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`;
      isTagUpdating = false;
      sourceListBtn.disabled = false;
      console.error('Error sending AI prompt:', error);
    }
  } else {
    console.error('No empty prompt allowed');
  }
}




// Your existing copyText function
function copyText(text: string) {
  // Copy text to clipboard logic
  const tempTextArea = document.createElement('textarea');
  tempTextArea.value = text;
  document.body.appendChild(tempTextArea);
  tempTextArea.select();
  document.execCommand('copy');
  document.body.removeChild(tempTextArea);

}


async function logout() {
  if (isGlossaryActive) {
    await removeMatchingContentControls();
  }
  sessionStorage.clear();
  window.location.hash = '#/new';
  initialised = true;
  document.getElementById('logo-header').innerHTML = ``;
  document.getElementById('header').innerHTML = ``
  login();
}


async function displayAiTagList() {
  if (isGlossaryActive) {
    await removeMatchingContentControls()
  }
  const container = document.getElementById('app-body');
  container.innerHTML = `
  <div class="d-flex justify-content-between px-2">
      <button class="btn btn-primary btn-sm bg-primary-clr c-pointer text-white  mb-3 mt-2 px-3 py-2" id="addgenaitag">
        <i class="fa fa-plus text-light px-1"></i>
        Add
    </button>

     <button class="btn btn-primary btn-sm bg-primary-clr c-pointer text-white mb-3 mt-2 px-3 py-2" id="applyAITag">
        <i class="fa fa-robot text-light px-1"></i>
        Apply
    </button>
    </div>

    <div class="card-container"  id="card-container">
    </div>
  `; // Clear any previous content
  const Cardcontainer = document.getElementById('card-container');
  document.getElementById('applyAITag').addEventListener('click', applyAITagFn);

  document.getElementById('addgenaitag').addEventListener('click', addGenAITags);

  for (let i = 0; i < aiTagList.length; i++) {
    const tag = aiTagList[i];
    const accordionItem = document.createElement('div');
    accordionItem.classList.add('accordion');
    accordionItem.id = `accordion-item-${i}`; // Replace 'yourUniqueId' with your desired ID

    const headerId = `flush-headingOne-${i}`;
    const collapseId = `flush-collapseOne-${i}`;

    const radioButtonsHTML = await generateRadioButtons(tag, i);

    accordionItem.innerHTML = accordionContent(headerId, collapseId, tag, radioButtonsHTML, i);

    Cardcontainer.appendChild(accordionItem);
    mentionDropdownFn(`chatbox-${i}`, `mention-dropdown-${i}`, 'edit');
    document.getElementById(`doNotApply-${i}`)?.addEventListener('change', () => onDoNotApplyChange(event, i, tag));

    document.getElementById(`sendPrompt-${i}`)?.addEventListener('click', () => {
      const textareaValue = (document.getElementById(`chatbox-${i}`) as HTMLTextAreaElement).value;

      sendPrompt(tag, textareaValue, i)
    });

    document.getElementById(`insert-tag-${i}`)?.addEventListener('click', () => {
      insertTagPrompt(i)
    })

    document.getElementById(`changeSource-${i}`)?.addEventListener('click', () => {
      const textareaValue = (document.getElementById(`chatbox-${i}`) as HTMLTextAreaElement).value;

      // const accordionbody=document.getElementById(`accordion-body-${i}`).innerHTML=''
      createMultiSelectDropdown(i, tag, radioButtonsHTML, textareaValue)
    })
  }

  // Add event listeners after rendering
  addAccordionListeners();
  addCopyListeners();

}

async function insertTagPrompt(index: any) {
  return Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      await context.sync();

      if (!selection) {
        throw new Error('Selection is invalid or not found.');
      }


      if (aiTagList[index].EditorValue === '') {
        selection.insertParagraph(`#${aiTagList[index].DisplayName}#`, Word.InsertLocation.before);
      } else {
        let content = removeQuotes(aiTagList[index].EditorValue);
        let lines = content.split(/\r?\n/); // Handle both \r\n and \n

        lines.forEach(line => {
          selection.insertParagraph(line, Word.InsertLocation.before);
        });
      }


      await context.sync();
    } catch (error) {
      console.error('Detailed error:', error);
    }
  });
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

async function applyAITagFn() {
  return Word.run(async (context) => {
    try {
      const body = context.document.body;
      context.load(body, 'text');
      await context.sync();

      // Iterate over the aiTagList to search and replace
      for (let i = 0; i < aiTagList.length; i++) {
        const tag = aiTagList[i];
        // Clean up the EditorValue by removing quotes
        tag.EditorValue = removeQuotes(tag.EditorValue);

        // Search for all instances of the tag.DisplayName enclosed with `#`
        const searchResults = body.search(`#${tag.DisplayName}#`, {
          matchCase: false,
          matchWholeWord: false,
        });

        // Load the search results to ensure they are available for further operations
        context.load(searchResults, 'items');

        await context.sync(); // Synchronize to fetch the search results

        // Log the number of search results for debugging
        console.log(`Found ${searchResults.items.length} instances of #${tag.DisplayName}#`);

        // Replace each found instance with tag.EditorValue
        searchResults.items.forEach((item: any) => {
          // Ensure the EditorValue is not empty before replacing
          if (tag.EditorValue !== "" && !tag.IsApplied) {
            item.insertText(tag.EditorValue, Word.InsertLocation.replace);
          }
        });

        // Additional sync after each replacement
        await context.sync();
      }

      // Final sync to apply all changes
      await context.sync();
    } catch (err) {
      console.error("Error during tag application:", err);
    }
  });
}




async function onRadioChange(tag, tagIndex, chatIndex) {
  if (!isTagUpdating) {
    isTagUpdating = true;
    const iconelement = document.getElementById(`sendPrompt-${tagIndex}`)
    iconelement.innerHTML = `<i class="fa fa-spinner fa-spin text-white"></i>`
    const chat = tag.FilteredReportHeadAIHistoryList[chatIndex];
    let payload = JSON.parse(JSON.stringify(chat));
    payload.Container = dataList.Container;
    payload.Selected = 1;
    const matchingKey = availableKeys.find(prop => prop.DisplayName === tag.DisplayName);
    if (matchingKey) {
      matchingKey.EditorValue = payload.Response;
    }
    try {
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

      if (data['Data']) {
        tag.ReportHeadAIHistoryList = JSON.parse(JSON.stringify(data['Data']));
        tag.FilteredReportHeadAIHistoryList = [];

        tag.ReportHeadAIHistoryList.forEach((historyList) => {
          historyList.Response = removeQuotes(historyList.Response);
          tag.FilteredReportHeadAIHistoryList.unshift(historyList);
        });

        // Use querySelectorAll to remove 'ai-selected-reply' from all elements
        const selectedParent = document.getElementById(`selected-response-parent-${tagIndex}`)
        const allSelectedDivs = selectedParent.querySelectorAll('.ai-selected-reply');
        allSelectedDivs.forEach(div => {
          div.classList.remove('ai-selected-reply');
          div.classList.add('bg-light');
        });

        const selectElement = document.getElementById(`selected-response-${tagIndex}${chatIndex}`);
        if (selectElement) {
          selectElement.classList.remove('bg-light');
          selectElement.classList.add('ai-selected-reply');
        }


        tag.UserValue = chat.Response;
        tag.EditorValue = chat.Response;
        tag.text = chat.Response;
      }

    } catch (error) {
      console.error('Error updating AI data:', error);
    } finally {
      const iconelement = document.getElementById(`sendPrompt-${tagIndex}`)
      iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`
      isTagUpdating = false;
    }
  }
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
        const response = await fetch(`${baseUrl}/api/glossary-template/client-id/${dataList?.ClientID}`, {
          method: 'POST', // or 'POST', depending on your API
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
        layTerms = data.Data
        if (data.Data.length > 0) {
          glossaryName = data.Data[0].GlossaryTemplate
          loadGlossary()
        } else {
          document.getElementById('app-body').innerHTML = `
       <p class="text-center">Data not available<p/>
  `
        }

        // alert('Glossary data loaded successfully.');
      } catch (error) {
        console.error('Error fetching glossary data:', error);
        // Optionally show an error message to the user
        // alert('Error fetching glossary data.');
      }
      // Sort terms by length (longest first)
      layTerms.sort((a, b) => b.ClinicalTerm.length - a.ClinicalTerm.length);

      const processedTerms = new Set(); // Track added larger terms

      // Filter out smaller terms if they are included in a larger term
      const filteredTerms = layTerms.filter(term => {
        for (const biggerTerm of processedTerms) {
          if (biggerTerm.includes(term.ClinicalTerm.toLowerCase())) {
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
            const fontProps = [];
            // Add properties, or "empty" if they are null/undefined
            if (font.highlightColor !== null && font.highlightColor !== undefined) {
              fontProps.push(`highlight:${font.highlightColor}`);
            }
          
            if (font.color !== null && font.color !== undefined) {
              fontProps.push(`color:${font.color}`);
            }
          
            if (font.bold !== undefined) {
              fontProps.push(`bold:${font.bold}`);
            }
          
            if (font.italic !== undefined) {
              fontProps.push(`italic:${font.italic}`);
            }
          
            if (font.underline !== undefined) {
              fontProps.push(`underline:${font.underline}`);
            }
          
            if (font.size !== null && font.size !== undefined) {
              fontProps.push(`size:${font.size}`);
            }
          
            // Add font family if available
            if (font.name !== null && font.name !== undefined) {
              fontProps.push(`family:${font.name}`);
            }
          
            // Set the tag to include all the collected font properties
            contentControl.tag = fontProps.join(', ');

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

      // Load the selection's text
      selection.load('text');
      await context.sync();
      // Check if the selected text contains the clinicalTerm
      if (selection.text.toLowerCase().includes(clinicalTerm.toLowerCase())) {
        // Search for the clinicalTerm in the document
        const searchResults = selection.search(clinicalTerm, { matchCase: false, matchWholeWord: false });
        searchResults.load('items');

        await context.sync();

        // Replace each occurrence of the clinicalTerm with the layTerm
        searchResults.items.forEach(item => {
          item.insertText(layTerm, 'replace');
          // Remove the highlight color (set to white or no highlight)
        });
        await context.sync();
        displayElement.style.display = 'none';

        console.log(`Replaced '${clinicalTerm}' with '${layTerm}' and removed highlight in the document.`);
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
      
          if (control.tag) {
            // Parse the tag into an object of type FontProps
            const fontProps: FontProps = control.tag.split(',').reduce((props, item) => {
              const [key, value] = item.split(':');
              if (key && value) {
                props[key.trim() as keyof FontProps] = value.trim();
              }
              return props;
            }, {} as FontProps); // Explicitly cast to FontProps
      
            // Apply properties to the range
            if (fontProps.highlight) {
              if (/^#[0-9A-Fa-f]{6}$/.test(fontProps.highlight)) {
                range.font.highlightColor = fontProps.highlight;
              } else {
                range.font.highlightColor = null; // Clear highlight if invalid
              }
            }
      
            if (fontProps.color) {
              range.font.color = fontProps.color;
            }
      
            if (fontProps.bold === 'true') {
              range.font.bold = true;
            } else if (fontProps.bold === 'false') {
              range.font.bold = false;
            }
      
            if (fontProps.italic === 'true') {
              range.font.italic = true;
            } else if (fontProps.italic === 'false') {
              range.font.italic = false;
            }
      
            if (fontProps.underline === 'true') {
              range.font.underline = "Single";
            } else if (fontProps.underline === 'false') {
              range.font.underline = "None";
            }
      
            if (fontProps.size) {
              range.font.size = parseFloat(fontProps.size);
            }
      
            if (fontProps.family) {
              range.font.name = fontProps.family; // Apply the font family
            }
            await context.sync();

          }
          await context.sync();
          // Delete the content control after applying styles
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

    const htmlBody = `
      <div class="container mt-3">
        <div class="card">
          <div class="card-header">
            <h5 class="card-title">Search Tags</h5>
          </div>
          <div class="card-body">
            <div class="form-group">
              <input type="text" id="search-box" class="form-control" placeholder="Search Tags..." autocomplete="off" />
            </div>
            <ul id="suggestion-list" class="list-group mt-2"></ul>
          </div>
        </div>
      </div>
    `;

    document.getElementById('app-body').innerHTML = htmlBody;

    const searchBox = document.getElementById('search-box');
    const suggestionList = document.getElementById('suggestion-list');

    // Function to filter and display suggestions
    function updateSuggestions() {
      const searchTerm = searchBox.value.toLowerCase();
      if (searchTerm === '') {
        suggestionList.innerHTML = ``;
        return;
      }
      suggestionList.innerHTML = '';

      // Filter mention list based on search term
      const filteredMentions = availableKeys.filter(mention =>
        mention.DisplayName.toLowerCase().includes(searchTerm)
      );

      // Render filtered suggestions
      filteredMentions.forEach(mention => {
        const listItem = document.createElement('li');
        listItem.className = 'list-group-item list-group-item-action';
        listItem.textContent = mention.DisplayName;
        listItem.onclick = () => {
          // Replace # with the selected value (adjust as needed)
          replaceMention(mention, mention.ComponentKeyDataType);
          suggestionList.innerHTML = ''; // Clear suggestions after selection
        };
        suggestionList.appendChild(listItem);
      });
    }

    // Add input event listener to the search box
    searchBox.addEventListener('input', updateSuggestions);
  }
}

async function addGenAITags() {
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

    const htmlBody = `
      <div class="modal-dialog">
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
              <div class="mb-3">
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
              <div class="text-end mt-3">
                <button id="cancel-btn-gen-ai" class="btn btn-danger bg-danger-clr px-3 me-2"><i class="fa fa-reply me-2"></i>Cancel</button>
                <button type="submit" class="btn btn-success bg-success-clr" id="text-gen-save"><i class="fa fa-check-circle me-2"></i>Save</button>
              </div>
            </form>
          </div>
        </div>
      </div>`;

    // Add modal HTML to the DOM
    document.getElementById('app-body').innerHTML = htmlBody;
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

    document.getElementById('cancel-btn-gen-ai').addEventListener('click', displayAiTagList);

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
      form.addEventListener('submit', function (e) {
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

          createTextGenTag(formData);
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
      });

      // Event listener for "Available to All Sponsors" checkbox
      availableForAllCheckbox.addEventListener('change', function () {
        if (this.checked) {
          checkAndDisableSponsors();
        } else {
          enableSponsors();
        }
      });

      // Add event listener to prevent dropdown close on item selection
      document.querySelectorAll('.dropdown-item').forEach(item => {
        item.addEventListener('click', function (event) {
          event.stopPropagation(); // Prevent dropdown from closing
          const checkbox = this.querySelector('.form-check-input');
          if (checkbox) {
            if (!checkbox.disabled) {
              checkbox.checked = !checkbox.checked;
            }

            if (checkbox.id === 'selectAll') {
              const isChecked = checkbox.checked;
              sponsorDropdownItems.forEach(cb => {
                if (!cb.disabled) cb.checked = isChecked;
              });
            }
            updateDropdownLabel();
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
    const aiTagBtn = document.getElementById('aitag');
    const mentionBtn = document.getElementById('mention');
    const formatDropdownBtn = document.getElementById('formatDropdown');
    const cancelBtnGenAi = document.getElementById('cancel-btn-gen-ai');

    aiTagBtn.disabled = true;
    cancelBtnGenAi.disabled = true;
    mentionBtn.disabled = true;
    formatDropdownBtn.disabled = true;
    iconelement.innerHTML = `<i class="fa fa-spinner fa-spin text-white me-2"></i>Save`;
    iconelement.disabled = true;

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
    if (data['Data'] && data['Status']) {
      fetchDocument('AIpanel');
    } else {
      aiTagBtn.disabled = false;
      cancelBtnGenAi.disabled = false;
      mentionBtn.disabled = false;
      formatDropdownBtn.disabled = false;
      iconelement.disabled = false;
      iconelement.innerHTML = `<i class="fa fa-check-circle me-2"></i>Save`;
      showAddTagError(data['Data'])
    }

  } catch (error) {
    console.error('Error creating text generation tag:', error);
  }
}


function mentionDropdownFn(textareaId, DropdownId, action) {
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
        const tableElement = doc.querySelector('table');

        if (!tableElement) {
          selection.insertParagraph(`#${word.DisplayName}#`, Word.InsertLocation.before);
          throw new Error('No table found in the provided HTML.');
        }

        const rows = Array.from(tableElement.querySelectorAll('tr'));

        if (rows.length === 0) {
          throw new Error('The table does not contain any rows.');
        }

        const maxCols = Math.max(...rows.map(row => {
          return Array.from(row.querySelectorAll('td, th')).reduce((sum, cell) => {
            return sum + (parseInt(cell.getAttribute('colspan') || '1', 10));
          }, 0);
        }));

        const paragraph = selection.insertParagraph("", Word.InsertLocation.before);
        await context.sync();

        if (!paragraph) {
          throw new Error('Failed to insert the paragraph.');
        }

        const table = paragraph.insertTable(rows.length, maxCols, Word.InsertLocation.after);
        await context.sync();

        if (!table) {
          throw new Error('Failed to insert the table.');
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

            // Ensure cellIndex is within bounds
            if (cellIndex >= maxCols) {
              // Adjust cellIndex to fit within table dimensions
              cellIndex = maxCols - 1;
            }

            // Set cell value
            try {
              table.getCell(rowIndex, cellIndex).value = cellText;

              // Clear cells that span columns
              for (let i = 1; i < colspan; i++) {
                if (cellIndex + i < maxCols) {
                  table.getCell(rowIndex, cellIndex + i).value = "";
                }
              }

              // Update rowspanTracker
              if (rowspan > 1) {
                for (let i = 0; i < colspan; i++) {
                  if (cellIndex + i < maxCols) {
                    rowspanTracker[cellIndex + i] = rowspan - 1;
                  }
                }
              }

              // Advance cellIndex by colspan
              cellIndex += colspan;
              if (cellIndex >= maxCols) {
                // Adjust cellIndex if it exceeds the table width
                cellIndex = maxCols - 1;
              }
            } catch (cellError) {
              console.error('Error setting cell value:', cellError);
            }
          });
        });
      } else {
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



function createMultiSelectDropdown(i, tag, radioButtonsHTML, textareaValue) {
  const cardContainer = document.getElementById('card-container');

  // Get the current scroll position
  const scrollPosition = cardContainer.scrollTop;

  const multiSelectHTML = `
  <div class='p-3 bg-light'>
    <div class="mb-3">
      <label for="source-select-${i}" class="form-label"><span class="text-danger">*</span> Select Sources</label>
      <div class="dropdown w-100">
        <button 
          class="btn btn-light border w-100 text-start d-flex justify-content-between align-items-start dropdown-toggle dropdown-toggle-sources" 
          type="button" 
          id="sourceDropdown-${i}" 
          data-bs-toggle="dropdown" 
          aria-expanded="false">
          <span id="sourceDropdownLabel-${i}" class='sourceDropdownLabel'></span>
          <span class="dropdown-toggle-icon dropdown-toggle-icon-s"></span>
        </button>
        <ul class="dropdown-menu w-100 p-2" aria-labelledby="sourceDropdown-${i}" style="box-shadow: 0 4px 8px rgba(0,0,0,0.1); z-index: 10000;">
          <li class="dropdown-item p-2" style="cursor: pointer;" data-checkbox-id="selectAll-${i}">
            <div class="form-check">
              <input class="form-check-input" type="checkbox" value="selectAll" id="selectAll-${i}">
              <label class="form-check-label" for="selectAll-${i}">Select All</label>
            </div>
          </li>
          ${sourceList
      .map(
        (source, index) => `
              <li class="dropdown-item p-2" style="cursor: pointer;" data-checkbox-id="source-${i}-${index}">
                <div class="form-check">
                  <input class="form-check-input source-checkbox" type="checkbox" value="${source.SourceName}" id="source-${i}-${index}">
                  <label class="form-check-label text-prewrap" for="source-${i}-${index}">${source.SourceName}</label>
                </div>
              </li>
            `
      )
      .join('')}
        </ul>
      </div>
    </div>
    <div class="d-flex justify-content-end mt-2">
      <button id="cancel-src-btn-${i}" class="btn btn-danger bg-danger-clr px-3 me-2"><i class="fa fa-reply me-2"></i>Cancel</button>
      <button id="ok-src-btn-${i}" class="btn btn-success bg-success-clr px-3"><i class="fa fa-check-circle me-2"></i>Save</button>
    </div>
  </div>
  `;

  // Get the accordion body element
  const accordionBody = document.getElementById(`accordion-body-${i}`);

  // Set height to ensure the dropdown button appears with some space
  accordionBody.style.minHeight = "20vh";  // You can adjust this based on your needs

  // Clear existing content and insert the dropdown
  accordionBody.innerHTML = multiSelectHTML;

  // Temporary array to hold the selected sources
  let selectedSources = [];

  // Get "Select All" checkbox and individual source checkboxes
  const selectAllCheckbox = document.getElementById(`selectAll-${i}`);
  const individualCheckboxes = document.querySelectorAll(`#accordion-body-${i} .source-checkbox`);
  const sourceDropdownLabel = document.getElementById(`sourceDropdownLabel-${i}`);

  // Function to update the dropdown label based on selected checkboxes
  function updateLabel() {
    const selectedSourceNames = selectedSources;
    if (selectedSourceNames.length > 0) {
      sourceDropdownLabel.innerText = selectedSourceNames.join(', ');  // Display selected sources as comma-separated
    } else {
      sourceDropdownLabel.innerText = ' ';  // Default text when no sources are selected
    }
  }

  // Handle "Select All" functionality
  selectAllCheckbox.addEventListener("change", function () {
    const checkboxes = document.querySelectorAll(`#accordion-body-${i} .source-checkbox`);
    checkboxes.forEach((checkbox) => {
      checkbox.checked = this.checked;  // If "Select All" is checked, all checkboxes are checked, and vice versa
      // Update the temporary array based on the state of the checkboxes
      if (checkbox.checked) {
        if (!selectedSources.includes(checkbox.value)) {
          selectedSources.push(checkbox.value);  // Add source to the array if checked
        }
      } else {
        selectedSources = selectedSources.filter((source) => source !== checkbox.value);  // Remove unchecked sources
      }
    });

    updateLabel();  // Update the label when "Select All" changes
  });

  // Handle individual checkbox state change
  individualCheckboxes.forEach((checkbox) => {
    checkbox.addEventListener("change", function () {
      if (checkbox.checked) {
        if (!selectedSources.includes(checkbox.value)) {
          selectedSources.push(checkbox.value);  // Add source to the array if checked
        }
      } else {
        selectedSources = selectedSources.filter((source) => source !== checkbox.value);  // Remove unchecked sources
      }

      // Check if all individual checkboxes are checked
      const allChecked = Array.from(individualCheckboxes).every((checkbox) => checkbox.checked);
      selectAllCheckbox.checked = allChecked;  // Update "Select All" checkbox based on individual selections

      updateLabel();  // Update the label when an individual checkbox is changed
    });

    // Add click event listener to the whole list item (not just the checkbox)
    const listItem = checkbox.closest("li");
    listItem.addEventListener("click", function (event) {
      checkbox.checked = !checkbox.checked;  // Toggle checkbox state

      if (checkbox.checked) {
        if (!selectedSources.includes(checkbox.value)) {
          selectedSources.push(checkbox.value);  // Add source to the array if checked
        }
      } else {
        selectedSources = selectedSources.filter((source) => source !== checkbox.value);  // Remove unchecked sources
      }

      const allChecked = Array.from(individualCheckboxes).every((checkbox) => checkbox.checked);
      selectAllCheckbox.checked = allChecked;  // Update "Select All" checkbox

      updateLabel();  // Update the label when a list item is clicked

      event.stopPropagation();  // Prevent dropdown from closing when clicking on list item
    });
  });

  // Prevent the dropdown from closing when clicking on the "Select All" checkbox label
  const selectAllItem = document.querySelector(`#accordion-body-${i} .dropdown-item[data-checkbox-id="selectAll-${i}"]`);
  selectAllItem.addEventListener("click", function (event) {
    selectAllCheckbox.checked = !selectAllCheckbox.checked;  // Toggle "Select All" checkbox

    // Update the state of all source checkboxes based on "Select All"
    const checkboxes = document.querySelectorAll(`#accordion-body-${i} .source-checkbox`);
    checkboxes.forEach((checkbox) => {
      checkbox.checked = selectAllCheckbox.checked;  // If "Select All" is checked, all checkboxes are checked
    });

    // Update the temporary array based on "Select All"
    selectedSources = selectAllCheckbox.checked
      ? Array.from(checkboxes).map((checkbox) => checkbox.value)  // Add all sources to the array if "Select All" is checked
      : [];

    updateLabel();  // Update the label when "Select All" item is clicked

    event.stopPropagation();  // Prevent dropdown from closing when clicking on the "Select All" item
  });

  // Initialize checkboxes based on tag.Sources
  if (tag.Sources && tag.Sources.length > 0) {
    individualCheckboxes.forEach((checkbox) => {
      if (tag.Sources.includes(checkbox.value)) {
        checkbox.checked = true;  // Check the checkbox if its value is in tag.Sources
        selectedSources.push(checkbox.value);  // Add the source to the temporary array
      }
    });

    // Also check "Select All" if all checkboxes are selected
    const allChecked = Array.from(individualCheckboxes).every((checkbox) => checkbox.checked);
    selectAllCheckbox.checked = allChecked;
    updateLabel();  // Update the label when "Select All" changes

  }

  // Handle "OK" button click
  document.getElementById(`ok-src-btn-${i}`).addEventListener("click", function () {
    tag.Sources = [...selectedSources];  // Add selected sources to tag.Sources when "OK" is clicked
    tag.SourceValue = sourceList
      .filter(source => selectedSources.includes(source.SourceName))  // Find the sources with matching SourceNames
      .map(source => source.SourceValue);
    appendAccordionBody(i, tag, radioButtonsHTML, textareaValue, scrollPosition);
  });

  document.getElementById(`cancel-src-btn-${i}`).addEventListener("click", function () {
    appendAccordionBody(i, tag, radioButtonsHTML, textareaValue, scrollPosition);
  });
}



function appendAccordionBody(i, tag, radioButtonsHTML, textareaValue, scrollPosition) {

  const tooltipButton = tag.Sources && tag.Sources.length > 0
    ? `  <span class="tooltiptext">${tag.Sources}</span>`
    : '';


  const accordionBody = document.getElementById(`accordion-body-${i}`);

  // Set height to ensure the dropdown button appears with some space
  accordionBody.style.minHeight = "20vh";  // You can adjust this based on your needs

  // Clear existing content and insert the dropdown
  accordionBody.innerHTML =
    `
    <div class="chatbox" id="selected-response-parent-${i}">
   
            ${radioButtonsHTML}
    </div>
    <div class="form-check form-switch chatbox m-0">
        <label class="form-check-label pb-3" for="doNotApply-${i}">
           <span class="fs-12">Do not apply</span>
        </label>
            <input class="form-check-input"
                   type="checkbox"
                   id="doNotApply-${i}"
                   ${tag.IsApplied ? 'checked' : ''}>
      </div>
      <div class="d-flex align-items-end justify-content-end chatbox p-2">
           <textarea class="form-control"
                      rows="5"
                      id="chatbox-${i}"
                      placeholder="Type here">${textareaValue}</textarea>
              <div id="mention-dropdown-${i}" class="dropdown-menu"></div>
              <div class="d-flex flex-column align-self-end me-3">
                <button class="btn btn-secondary text-light ms-2 mb-2 ngb-tooltip" id="insert-tag-${i}">
                <span class="tooltiptext">Insert</span>
                <i class="fa fa-plus text-light c-pointer" ></i>
                </button>

                <button
                    class="btn btn-secondary ms-2 mb-2 text-white ngb-tooltip"
                    id="changeSource-${i}">
                    ${tooltipButton}
                 <i class="fa fa-file-lines text-white"></i>
                </button>
                <button type="submit"
                    class="btn btn-primary bg-primary-clr ms-2 text-white"
                    id="sendPrompt-${i}">
                  <i class="fa fa-paper-plane text-white"></i>
               </button>
             </div>
          </div>`;

  const cardContainer = document.getElementById('card-container');

  cardContainer.scrollTop = scrollPosition;
  mentionDropdownFn(`chatbox-${i}`, `mention-dropdown-${i}`, 'edit');
  document.getElementById(`doNotApply-${i}`)?.addEventListener('change', () => onDoNotApplyChange(event, i, tag));

  document.getElementById(`sendPrompt-${i}`)?.addEventListener('click', () => {
    const textareaValue = (document.getElementById(`chatbox-${i}`) as HTMLTextAreaElement).value;

    sendPrompt(tag, textareaValue, i)
  });


  document.getElementById(`insert-tag-${i}`)?.addEventListener('click', () => {
    insertTagPrompt(i)
  })

  document.getElementById(`changeSource-${i}`)?.addEventListener('click', () => {
    const textareaValue = (document.getElementById(`chatbox-${i}`) as HTMLTextAreaElement).value;

    // const accordionbody=document.getElementById(`accordion-body-${i}`).innerHTML=''
    createMultiSelectDropdown(i, tag, radioButtonsHTML, textareaValue)
  })

}
