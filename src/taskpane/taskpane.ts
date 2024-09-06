/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { data } from "./data";
let jwt = '';
let documentID = ''
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
      if (property) {
        documentID = property.value;
        login()
      } else {
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
          <button type="submit" class="btn btn-primary">Login</button>
        </div>
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
  document.getElementById('app-body').innerHTML = `
  <div id="button-container">

          <div class="loader" id="loader"></div>
          </div
`
  try {
    const response = await fetch('https://plsdevapp.azurewebsites.net/api/user/login', {
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
      loadLoginPage()
      throw new Error('Network response was not ok.');
    }

    const data = await response.json();
    if (data.Status === true && data['Data']) {
      if (data['Data'].ResponseStatus) {
        jwt = data.Data.Token;
        sessionStorage.setItem('token', jwt)

        window.location.hash = '#/dashboard';

      } else {
        loadLoginPage()
      }
    } else {
      loadLoginPage()
    }


    // Handle successful login (e.g., navigate to the next page or show a success message)

  } catch (error) {
    loadLoginPage()
    console.error('Error during login:', error);
    // Handle login error (e.g., show an error message)
  }
}

function displayMenu() {
  // document.getElementById('aitag').addEventListener('click', redirectAI);
  fetchDocument();

}

async function fetchDocument() {
  try {
    const response = await fetch(`https://plsdevapp.azurewebsites.net/api/report/id/${documentID}`, {
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
    document.getElementById('logo-header').innerHTML=`
        <img  id="main-logo" src="./assets/logo.png" alt="" height="60" class="logo">`
    document.getElementById('header').innerHTML = `

    <div class="d-flex justify-content-around">
    <button class="btn  btn-dark " id="mention">Suggestions</button>
            <button class="btn  btn-dark " id="aitag">AI Text Panel</button>

        <button class="btn  btn-dark " id="glossary">Glossary</button>
</div>

`
    // <button class="btn  btn-dark me-2" id="imagebtn">Image</button>


    document.getElementById('mention').addEventListener('click', displayMentions);
    document.getElementById('glossary').addEventListener('click', fetchGlossary);

    document.getElementById('aitag').addEventListener('click', displayAiTagList);
    // document.getElementById('imagebtn').addEventListener('click', fetchGeneralImages);




    // Extracting the relevant AI group from the response
    dataList = data['Data'];
    const aiGroup = data['Data'].Group.find(element => element.DisplayName === 'AIGroup');
    GroupName = aiGroup ? aiGroup.Name : '';
    aiTagList = aiGroup ? aiGroup.GroupKey : [];

    availableKeys = data['Data'].GroupKeyAll.filter(element => element.ComponentKeyDataType === 'TABLE' || element.ComponentKeyDataType === 'TEXT');

    // Call function to display the AI Tag List on the UI

  } catch (error) {
    console.error('Error fetching glossary data:', error);
  }
}

async function fetchAIHistory(tag) {
  try {
    const response = await fetch(`https://plsdevapp.azurewebsites.net/api/report/ai-history/${tag.ID}`, {
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
      `<div class="row chatbox">
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


function accordianContent(headerId, collapseId, tag, radioButtonsHTML, i) {
  const body = `
   <h2 class="accordion-header" id="${headerId}">
  <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse"
    data-bs-target="#${collapseId}" aria-expanded="false" aria-controls="${collapseId}">
    ${tag.DisplayName}
  </button>
</h2>
<div id="${collapseId}" class="accordion-collapse collapse" aria-labelledby="${headerId}">
  <div class="accordion-body chatbox" id="selected-response-parent-${i}">
    ${radioButtonsHTML}
  </div>

  <div class="col-md-12 d-flex align-items-center justify-content-end chatbox p-3">
    <textarea class="form-control" rows="3" id="chatbox-${i}" placeholder="Type here"></textarea>
    <div class="d-flex align-self-end">
      <button type="submit" class="btn btn-primary ms-2 text-white" id="sendPrompt-${i}">
        <i class="fa fa-paper-plane text-white"></i>
      </button>
    </div>
  </div>
</div>
  `

  return body
}


async function sendPrompt(tag, prompt, index) {
  if (prompt !== '' && !isTagUpdating) {

    isTagUpdating = true;
    const iconelement = document.getElementById(`sendPrompt-${index}`)
    iconelement.innerHTML = `<i class="fa fa-spinner fa-spin text-white"></i>`
    const payload = {
      ReportHeadID: tag.FilteredReportHeadAIHistoryList[0].ReportHeadID,
      DocumentID: dataList.NCTID,
      DocumentType: dataList.DocumentType,
      TextSetting: dataList.TextSetting,
      DocumentTemplate: dataList.ReportTemplate,
      ReportHeadGroupKeyID:
        tag.FilteredReportHeadAIHistoryList[0].ReportHeadGroupKeyID,
      Container: dataList.Container,
      GroupName: GroupName,
      Prompt: prompt,
      PromptType: 1,
      Response: '',
      Selected: 0,
      ID: 0
    };

    try {
      const response = await fetch('https://plsdevapp.azurewebsites.net/api/report/ai-history/add', {
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

        tag.ReportHeadAIHistoryList = JSON.parse(
          JSON.stringify(data['Data'])
        );
        tag.FilteredReportHeadAIHistoryList = [];
        // tag.ReportHeadAIHistoryList.reverse();
        tag.ReportHeadAIHistoryList.forEach((historyList, i) => {
          historyList.Response = removeQuotes(historyList.Response);
          tag.FilteredReportHeadAIHistoryList.unshift(historyList);
        });

        const collapseId = `flush-collapseOne-${index}`;
        const headerId = `flush-headingOne-${index}`;

        const accordianBox = document.getElementById(`accordion-item-${index}`); // Replace 'yourUniqueId' with your desired ID


        const radioButtonsHTML = await generateRadioButtons(tag, index);

        accordianBox.innerHTML = accordianContent(headerId, collapseId, tag, radioButtonsHTML, index);

        const iconelement = document.getElementById(`sendPrompt-${index}`)
        iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`
        isTagUpdating = false;
      } else {
        const iconelement = document.getElementById(`sendPrompt-${index}`)
        iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`
        isTagUpdating = false;
      }

      // alert('Glossary data loaded successfully.');
    } catch (error) {
      const iconelement = document.getElementById(`sendPrompt-${index}`)
      iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`
      isTagUpdating = false
      console.error('Error sending ai prompt:', error);
      // Optionally show an error message to the user
      // alert('Error fetching glossary data.');
    }


  } else {
    console.error('No empty prompt allowed')
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


async function displayAiTagList() {
  if (isGlossaryActive) {
    await clearGlossary()
  }
  const container = document.getElementById('app-body');
  container.innerHTML = `
  <div class="d-flex justify-content-end p-1">
     <button class="btn btn-primary btn-sm c-pointer text-white me-2 mb-2" id="applyAITag">
        <i class="fa fa-robot text-light"></i>
        Apply
    </button>
    </div>

    <div class="card-container"  id="card-container">
    </div>
  `; // Clear any previous content
  const Cardcontainer = document.getElementById('card-container');
    document.getElementById('applyAITag').addEventListener('click', applyAITagFn);


  for (let i = 0; i < aiTagList.length; i++) {
    const tag = aiTagList[i];
    const accordionItem = document.createElement('div');
    accordionItem.classList.add('accordion-item');
    accordionItem.id = `accordion-item-${i}`; // Replace 'yourUniqueId' with your desired ID

    const headerId = `flush-headingOne-${i}`;
    const collapseId = `flush-collapseOne-${i}`;

    const radioButtonsHTML = await generateRadioButtons(tag, i);

    accordionItem.innerHTML = accordianContent(headerId, collapseId, tag, radioButtonsHTML, i);

    Cardcontainer.appendChild(accordionItem);
    document.getElementById(`sendPrompt-${i}`)?.addEventListener('click', () => {
      const textareaValue = (document.getElementById(`chatbox-${i}`) as HTMLTextAreaElement).value;

      sendPrompt(tag, textareaValue, i)
    });
  }





  // Add event listeners after rendering
  addAccordionListeners();
  addCopyListeners();
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
          matchCase: true,
          matchWholeWord: true,
        });

        // Load the search results to ensure they are available for further operations
        context.load(searchResults, 'items');

        await context.sync(); // Synchronize to fetch the search results

        // Log the number of search results for debugging
        console.log(`Found ${searchResults.items.length} instances of #${tag.DisplayName}#`);

        // Replace each found instance with tag.EditorValue
        searchResults.items.forEach((item: any) => {
          // Ensure the EditorValue is not empty before replacing
          if (tag.EditorValue !== "") {
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

    try {
      const response = await fetch('https://plsdevapp.azurewebsites.net/api/report/ai-history/update', {
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


    if (layTerms.length === 0) {

      document.getElementById('app-body').innerHTML = `
  <div id="button-container">

          <div class="loader" id="loader"></div>

        <div id="highlighted-text"></div>`

      try {
        const response = await fetch('https://plsdevapp.azurewebsites.net/api/glossary-template/id/3', {
          method: 'GET', // or 'POST', depending on your API
          headers: {
            'Authorization': `Bearer ${jwt}`
          }
        });
        if (!response.ok) {
          throw new Error('Network response was not ok.');
        }

        const data = await response.json();
        layTerms = data.Data.GlossaryTemplateData;
        glossaryName = data.Data.Name
        loadGlossary()
        // alert('Glossary data loaded successfully.');
      } catch (error) {
        console.error('Error fetching glossary data:', error);
        // Optionally show an error message to the user
        // alert('Error fetching glossary data.');
      }

    } else {
      loadGlossary()
    }
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

      const searchPromises = layTerms.map(term => {
        const searchResults = body.search(term.ClinicalTerm, { matchCase: true, matchWholeWord: true });
        searchResults.load("items");
        return searchResults;
      });
      await context.sync();



      searchPromises.forEach(searchResults => {
        searchResults.items.forEach(item => {
          item.font.highlightColor = "yellow";
        });
      });
      // document.getElementById('glossarycheck').style.display='block';
      isGlossaryActive = true;
      document.getElementById('app-body').innerHTML = `
      <div id="button-container">
        <button class="btn btn-secondary me-2 clear-glossary btn-sm" id="clearGlossary">Clear Glossary</button>
      </div>

      <div id="highlighted-text"></div>
      
`




      // document.getElementById('loader').style.display='none';
      // document.getElementById('Clear').style.display='block';

      // Set the flag when glossary is marked

      await context.sync();
      document.getElementById('clearGlossary').addEventListener('click', clearGlossary);
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
        const searchPromises = layTerms.map(term => {
          const searchResults = selection.search(term.ClinicalTerm, { matchCase: false, matchWholeWord: true });
          searchResults.load("items");
          return searchResults;
        });

        await context.sync();
        const selectedWords = []
        searchPromises.forEach(searchResults => {
          searchResults.items.forEach(item => {
            selectedWords.push(item.text);
          });
        });
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

    // Group lay terms by their clinical term
    const groupedTerms: { [clinicalTerm: string]: string[] } = {};

    words.forEach(word => {
      layTerms.forEach(term => {
        if (term.ClinicalTerm === word) {
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
  }
}


async function replaceClinicalTerm(clinicalTerm: string, layTerm: string) {
  try {
    await Word.run(async (context) => {
      // Get the current selection
      const selection = context.document.getSelection();

      // Load the selection's text
      selection.load('text');
      await context.sync();

      // Check if the selected text contains the clinicalTerm
      if (selection.text.includes(clinicalTerm)) {
        // Search for the clinicalTerm in the document
        const searchResults = selection.search(clinicalTerm, { matchCase: false, matchWholeWord: true });
        searchResults.load('items');

        await context.sync();

        // Replace each occurrence of the clinicalTerm with the layTerm
        searchResults.items.forEach(item => {
          item.insertText(layTerm, 'replace');

          // Remove the highlight color (set to white or no highlight)
          item.font.highlightColor = 'white';
        });
        await context.sync();

        console.log(`Replaced '${clinicalTerm}' with '${layTerm}' and removed highlight in the document.`);
      } else {
        console.log(`Selected text does not contain '${clinicalTerm}'.`);
      }
    });
  } catch (error) {
    console.error('Error replacing term:', error);
  }
}


async function clearGlossary() {
  try {
    await Word.run(async (context) => {
      document.getElementById('app-body').innerHTML = `
      <div id="button-container">
    
              <div class="loader" id="loader"></div>
    
            <div id="highlighted-text"></div>`
      const body = context.document.body;

      const searchPromises = layTerms.map(term => {
        const searchResults = body.search(term.ClinicalTerm, { matchCase: false, matchWholeWord: true });
        searchResults.load("items");
        return searchResults;
      });

      await context.sync();

      searchPromises.forEach(searchResults => {
        searchResults.items.forEach(item => {
          item.font.highlightColor = 'white'; // Reset highlight color
        });
      });
      document.getElementById('app-body').innerHTML = `
      <div id="button-container">
        <button class="btn btn-secondary me-2 mark-glossary btn-sm" id="applyglossary">Apply Glossary</button>
      </div>
`
      await context.sync();
      isGlossaryActive = false
      document.getElementById('applyglossary').addEventListener('click', applyglossary);


    });


    console.log('Glossary cleared successfully');
  } catch (error) {
    console.error('Error clearing glossary:', error);
  }
}



async function displayMentions() {
  if (!isTagUpdating) { // Check if isTagUpdating is false
    if (isGlossaryActive) {
      await clearGlossary();
    }
    const htmlBody = `
      <div class="container mt-3">
        <div class="card">
          <div class="card-header">
            <h5 class="card-title">Search Suggestions</h5>
          </div>
          <div class="card-body">
            <div class="form-group">
              <input type="text" id="search-box" class="form-control" placeholder="Search Suggestions..." autocomplete="off" />
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
          replaceMention(mention.EditorValue, mention.ComponentKeyDataType);
          // Clear suggestions after selection
        };
        suggestionList.appendChild(listItem);
      });
    }

    // Add input event listener to the search box
    searchBox.addEventListener('input', updateSuggestions);
  }
}



export async function replaceMention(word: string, type: any) {
  return Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      await context.sync();

      if (!selection) {
        throw new Error('Selection is invalid or not found.');
      }

      if (type === 'TABLE') {
        const parser = new DOMParser();
        const doc = parser.parseFromString(word, 'text/html');
        const tableElement = doc.querySelector('table');

        if (!tableElement) {
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
        selection.insertParagraph(word, Word.InsertLocation.before);
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


// async function fetchGeneralImages() {
//   if (imageList.length === 0) {
//     try {
//       const response = await fetch('https://plsdevapp.azurewebsites.net/api/image/general', {
//         method: 'GET',
//         headers: {
//           'Content-Type': 'application/json',
//           'Authorization': `Bearer ${jwt}`
//         },
//       });

//       if (!response.ok) {
//         throw new Error('Network response was not ok.');
//       }

//       const data = await response.json();
//       imageList = data['Data'].Image;
//       imageDisplay();
//     } catch (error) {
//       console.error('Error during login:', error);
//       // Handle login error (e.g., show an error message)
//     }
//   } else {
//     imageDisplay()
//   }


// // }

// function imageDisplay() {
//   document.getElementById('app-body').innerHTML = `
//       <div class="container mt-3">
//       <div class="card">
//         <div class="card-header">
//           <h5 class="card-title">Search Images</h5>
//         </div>
//         <div class="card-body">
//           <div class="form-group">
//             <input type="text" id="search-box" class="form-control" placeholder="Search Images..." autocomplete="off" />
//           </div>
//           <ul id="image-list" class="list-group mt-2"></ul>
//         </div>
//       </div>
//     </div>
//   `

//   const searchBox = document.getElementById('search-box');
//   const imageBox = document.getElementById('image-list');

//   // Function to filter and display suggestions
//   function updateSuggestions() {
//     const searchTerm = searchBox.value.toLowerCase();
//     imageBox.innerHTML = '';
//     // Filter mention list based on search term
//     const filteredImages = imageList.filter(image =>
//       image.ImageName.toLowerCase().includes(searchTerm)
//     );

//     // Render filtered suggestions
//     filteredImages.forEach(images => {
//       const listItem = document.createElement('li');
//       listItem.className = 'list-group-item list-group-item-action';
//       listItem.textContent = images.ImageName;
//       listItem.onclick = () => {
//         insertImageIntoWord(images.ImageData)
//         // Replace # with the selected value (adjust as needed)
//         // searchBox.value = '';
//         // suggestionList.innerHTML = '';
//         // replaceMention(images.EditorValue, images.ComponentKeyDataType)
//         // Clear suggestions after selection
//       };
//       imageBox.appendChild(listItem);
//     });
//   }

//   // Add input event listener to the search box
//   searchBox.addEventListener('input', updateSuggestions);
// }





// async function insertImageIntoWord(base64Image) {
//   await Word.run(async (context) => {
//     try {
//       const selection = context.document.getSelection();
//       await context.sync();
//       selection.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.before);
//       await context.sync();
//     } catch (err) {
//       console.log(err)
//     }
//   });
// }