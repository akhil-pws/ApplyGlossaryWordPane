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

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("Clear").onclick = clearGlossary;

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      handleSelectionChange
    );

    await fetchGlossaryData();
    await clearGlossary();
  }
});

async function fetchGlossaryData() {
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
  } catch (error) {
    console.error('Error fetching glossary data:', error);
  }
}

async function handleSelectionChange() {
  if (isGlossaryMarked) {
    await checkGlossary();
  }
}

function disableButtons() {
  document.getElementById("run")?.setAttribute("disabled", "true");
  document.getElementById("Clear")?.setAttribute("disabled", "true");
}

function enableButtons() {
  document.getElementById("run")?.removeAttribute("disabled");
  document.getElementById("Clear")?.removeAttribute("disabled");
}

async function run() {
  try {
    disableButtons();
    await Word.run(async (context) => {
      document.getElementById('loader').style.display = 'block';

      const body = context.document.body;
      const termsRegex = layTerms.map(term => `\\b${term.ClinicalTerm}\\b`).join('|');
      const searchResults = body.search(termsRegex, {
        matchCase: true,
        matchWholeWord: true,
        matchWildcards: true
      });
      searchResults.load("items");

      await context.sync();

      searchResults.items.forEach(item => {
        item.font.highlightColor = "yellow";
      });

      await context.sync();

      document.getElementById('loader').style.display = 'none';
      isGlossaryMarked = true;
    });

    console.log('Glossary applied successfully');
  } catch (error) {
    console.error('Error applying glossary:', error);
  } finally {
    enableButtons();
  }
}

async function clearGlossary() {
  try {
    disableButtons();
    await Word.run(async (context) => {
      document.getElementById('loader').style.display = 'block';

      const body = context.document.body;
      const termsRegex = layTerms.map(term => `\\b${term.ClinicalTerm}\\b`).join('|');
      const searchResults = body.search(termsRegex, {
        matchCase: false,
        matchWholeWord: true,
        matchWildcards: true
      });
      searchResults.load("items");

      await context.sync();

      searchResults.items.forEach(item => {
        item.font.highlightColor = 'white';
      });

      await context.sync();

      isGlossaryMarked = false;
      clearHighlightedText();
    });

    console.log('Glossary cleared successfully');
  } catch (error) {
    console.error('Error clearing glossary:', error);
  } finally {
    enableButtons();
    document.getElementById('loader').style.display = 'none';
  }
}

function clearHighlightedText() {
  const displayElement = document.getElementById('highlighted-text');
  if (displayElement) {
    displayElement.innerHTML = '';
  }
}

async function checkGlossary() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");

      await context.sync();

      if (selection.text) {
        const termsRegex = layTerms.map(term => `\\b${term.ClinicalTerm}\\b`).join('|');
        const searchResults = selection.search(termsRegex, {
          matchCase: false,
          matchWholeWord: true,
          matchWildcards: true
        });
        searchResults.load("items");

        await context.sync();

        const selectedWords = searchResults.items.map(item => item.text);
        displayHighlightedText(selectedWords);
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
    const fragment = document.createDocumentFragment(); // Use a DocumentFragment

    const groupedTerms = {};

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

    Object.keys(groupedTerms).forEach(clinicalTerm => {
      const mainBox = document.createElement('div');
      mainBox.className = 'box';

      const heading = document.createElement('h3');
      heading.textContent = clinicalTerm;
      mainBox.appendChild(heading);

      groupedTerms[clinicalTerm].forEach(layTerm => {
        const subBox = document.createElement('div');
        subBox.className = 'sub-box';
        subBox.textContent = layTerm;

        subBox.addEventListener('click', async () => {
          await replaceClinicalTerm(clinicalTerm, layTerm);
          mainBox.remove();
        });

        mainBox.appendChild(subBox);
      });

      fragment.appendChild(mainBox);
    });

    displayElement.innerHTML = ''; // Clear the current content
    displayElement.appendChild(fragment); // Append the fragment in one go
  }
}

async function replaceClinicalTerm(clinicalTerm, layTerm) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load('text');
      await context.sync();

      if (selection.text.includes(clinicalTerm)) {
        selection.insertText(layTerm, Word.InsertLocation.replace);
        await context.sync();
      }
    });
  } catch (error) {
    console.error('Error replacing term:', error);
  }
}
