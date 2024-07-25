/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// const layTerms=data.map(entry => ({
//   LayTerm: entry.LayTerm, // or entry.ClinicalTerm based on what you want to search
//   ClinicalTerm: entry.ClinicalTerm // Store the original term for reference
// }));

let layTerms
let isGlossaryMarked = false;
const jwt='eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJOYW1lIjoiYWtoaWxybyIsIk5hbWVJZGVudGlmaWVyIjoiMTQwOCIsIlVzZXJJRCI6IjE0MDgiLCJVc2VyTmFtZSI6ImFraGlscm8iLCJFbWFpbCI6ImFraGlsLmFAcGFjZXdpc2RvbS5jb20iLCJDbGllbnRJRCI6IjEwMDU5IiwiVEFBdXRoX0RCTmFtZSI6IlRBX0F1dGhfUGFjZURldiIsIkVycm9yTXNnIjoiXCJcIiIsIklzVmFsaWQiOiJUcnVlIiwiQXBwbGljYXRpb25Db2RlIjoiTElOSyIsImh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vd3MvMjAwOC8wNi9pZGVudGl0eS9jbGFpbXMvcm9sZSI6IlVzZXIiLCJBbm9ueW1pemF0aW9uX0RCTmFtZSI6IiIsIkJFQUNPTl9EQk5hbWUiOiIiLCJMSU5LX0RCTmFtZSI6IkxpbmtfTU1TIiwiUmVnaXN0cnlfREJOYW1lIjoiIiwiZXhwIjoxNzIyOTQyNDQ4LCJpc3MiOiJodHRwOi8vVHJpYWxBc3N1cmUuY29tIiwiYXVkIjoiaHR0cDovL1RyaWFsQXNzdXJlLmNvbSJ9.k-s_qBxCj-ArIIicYm_FRWv-332_0loUKLGuoPfmosw'

/* global document, Office, Word */

Office.onReady(async(info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    // document.getElementById("glossarycheck").onclick = checkGlossary;
    document.getElementById("Clear").onclick = clearGlossary;
    
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      handleSelectionChange
    );

    await fetchGlossaryData();
    await clearGlossary();
  }
});


function disableButtons() {
  document.getElementById("run")?.setAttribute("disabled", "true");
  // document.getElementById("glossarycheck")?.setAttribute("disabled", "true");
  document.getElementById("Clear")?.setAttribute("disabled", "true");
}
// Function to enable buttons
function enableButtons() {
  document.getElementById("run")?.removeAttribute("disabled");
  // document.getElementById("glossarycheck")?.removeAttribute("disabled");
  document.getElementById("Clear")?.removeAttribute("disabled");

}


export async function run() {
  try {
    await Word.run(async (context) => {
      document.getElementById("run")?.setAttribute("disabled", "true");
      document.getElementById('loader').style.display='block';

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
      document.getElementById("run")?.removeAttribute("disabled");

      document.getElementById('Clear').style.display='block';
      document.getElementById('run').style.display='none';
      document.getElementById('loader').style.display='none';

      isGlossaryMarked = true; // Set the flag when glossary is marked

      await context.sync();
    });

    // Optional: Notify user of completion
    console.log('Glossary applied successfully');
  } catch (error) {
    console.error('Error applying glossary:', error);
    // Optional: Notify user of error
    console.log('Error applying glossary. Please try again.');
  }
}

async function clearGlossary() {
  try {
    await Word.run(async (context) => {
      document.getElementById("Clear")?.setAttribute("disabled", "true");
      document.getElementById('loader').style.display='block';
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
      document.getElementById("Clear")?.removeAttribute("disabled");
      document.getElementById('Clear').style.display='none';
      document.getElementById('loader').style.display='none';

      document.getElementById('run').style.display='block';
      await context.sync();
      
      
      isGlossaryMarked = false;  // Clear the flag when glossary is cleared
      clearHighlightedText(); // Clear the highlighted text boxes
    });

    console.log('Glossary cleared successfully');
  } catch (error) {
    console.error('Error clearing glossary:', error);
  }
}

function clearHighlightedText() {
  const displayElement = document.getElementById('highlighted-text');
  if (displayElement) {
    displayElement.innerHTML = ''; // Clear the content of highlighted text
  }
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
      heading.textContent = clinicalTerm;
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


async function fetchGlossaryData() {
  document.getElementById('run').style.display='none';
  document.getElementById('Clear').style.display='none';

  disableButtons(); // Disable buttons before making the API call

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
    document.getElementById('loader').style.display='none';
    document.getElementById('run').style.display='block';
    layTerms = data.Data.GlossaryTemplateData;
   
      // alert('Glossary data loaded successfully.');
  } catch (error) {
    console.error('Error fetching glossary data:', error);
    // Optionally show an error message to the user
    // alert('Error fetching glossary data.');
  } finally {
    enableButtons(); // Re-enable buttons after the API call completes
  }

}

async function handleSelectionChange() {
  if (isGlossaryMarked) {
    await checkGlossary();
  }
}
