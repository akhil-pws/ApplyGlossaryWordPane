import { generateCheckboxHistory } from "./components/home";
import { theme } from "./taskpane";

export function insertLineWithHeadingStyle(range: Word.Range, line: string) {
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


export function removeQuotes(value: string): string {
  return value
    ? value
      .replace(/^"|"$/g, '')
      .replace(/\\n/g, '')
      .replace(/\*\*/g, '')
      .replace(/\\r/g, '')
    : '';
}


export async function insertSingleBookmark(text: any, DisplayName: any) {
  return Word.run(async (context) => {
    let range = context.document.getSelection();
    await context.sync(); // Ensure selection is ready

    // Replace spaces with underscores in DisplayName
    let cleanDisplayName = DisplayName.replace(/\s+/g, "_");

    let uniqueStr = new Date().getTime();
    let splitString = 'Split';
    let bookmarkName = `${cleanDisplayName}_${splitString}_${uniqueStr}`;

    // Insert text and get the range of inserted content
    let insertedTextRange = range.insertText(text, Word.InsertLocation.replace);

    await context.sync(); // Ensure text is inserted

    // Expand the range to cover the newly inserted text and apply bookmark
    insertedTextRange.insertBookmark(bookmarkName);

    await context.sync(); // Ensure bookmark is inserted
    console.log(`Single bookmark added: ${bookmarkName}`);
  });
}


export function copyText(text: string) {
  // Copy text to clipboard logic
  const tempTextArea = document.createElement('textarea');
  tempTextArea.value = text;
  document.body.appendChild(tempTextArea);
  tempTextArea.select();
  document.execCommand('copy');
  document.body.removeChild(tempTextArea);

}


export function switchToPromptBuilder() {
  // Remove active class from current tab
  document.querySelector('.nav-link.active')?.classList.remove('active');
  document.querySelector('.tab-pane.show.active')?.classList.remove('show', 'active');

  // Add active class to Prompt Builder tab
  document.getElementById('prompt-tab').classList.add('active');
  document.getElementById('add-prompt-template').classList.add('show', 'active');
}


export function switchToAddTag() {
  // Remove active class from current tab
  document.querySelector('.nav-link.active')?.classList.remove('active');
  document.querySelector('.tab-pane.show.active')?.classList.remove('show', 'active');

  // Add active class to Prompt Builder tab
  document.getElementById('tag-tab').classList.add('active');
  document.getElementById('add-tag-body').classList.add('show', 'active');
}

export function updateEditorFinalTable(data) {
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
  if (!jsonData || (Array.isArray(jsonData) && jsonData.length === 0)) {
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
          return typeof item === 'object'
            ? Object.entries(item).map(([k, v]) => `<strong>${k}:</strong> ${v}`).join('<br>')
            : item;
        }).join('<br>');
      } else {
        result[newKey] = value;
      }
    });
    return result;
  }

  if (!Array.isArray(jsonData)) {
    jsonData = Object.entries(jsonData).map(([key, value]) => ({ [key]: value }));
  }

  jsonData.forEach(item => {
    let flattenedItem = flattenObject(item);
    Object.keys(flattenedItem).forEach(key => headers.add(key));
    rows.push(flattenedItem);
  });

  let table = '<table border="1" cellspacing="0" cellpadding="5">';
  table += '<tr>' + [...headers].map(header => `<th>${header}</th>`).join('') + '</tr>';
  rows.forEach(row => {
    table += '<tr>' + [...headers].map(header => `<td>${row[header]}</td>`).join('') + '</tr>';
  });

  table += '</table>';
  return table;
}


export async function insertTagPrompt(tag: any) {
  return Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      await context.sync();

      if (!selection) {
        throw new Error('Selection is invalid or not found.');
      }

      if (tag.EditorValue === '') {
        selection.insertParagraph(`#${tag.DisplayName}#`, Word.InsertLocation.before);
      } else {
        if (tag.ComponentKeyDataType === 'TABLE') {
          const parser = new DOMParser();
          const doc = parser.parseFromString(tag.EditorValue, 'text/html');

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
        } else {
          let content = removeQuotes(tag.EditorValue);
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


export function generateChatHistoryHtml(chatList: any[]): string {
      const promptclass= theme==='Dark' ? 'bg-secondary text-light' : 'bg-white text-dark';
  
  return chatList.map((chat, index) =>
    `<div class="row chat-entry m-0 p-0">
            <div class="col-md-12 mt-2 p-2">
                <span class="float-end me-1">
                    <i class="fa fa-copy text-secondary c-pointer" title="Copy Prompt" id="copyPrompt-${index}"></i>
                </span>
                <span class="float-end w-75 me-2">
                    <div class="form-control h-34 d-flex align-items-center dynamic-height prompt-text ${promptclass}">
                        ${chat.Prompt}
                    </div>
                </span>
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
                    </div>
                </span>
            </div>
        </div>`
  ).join('');
}


export function chatfooter(tag: any) {
  const promptclass= theme==='Dark' ? 'bg-secondary text-light' : 'bg-white text-dark';
  const tooltipButton = tag.Sources && tag.Sources.length > 0
    ? `  <span class="tooltiptext">${tag.Sources}</span>`
    : '<span class="tooltiptext">Source</span>';
  return ` <textarea class="form-control ${promptclass}"
                      rows="5"
                      id="chatInput"
                      ></textarea>
            <div id="mention-dropdown" class="dropdown-menu"></div>
            <div class="d-flex flex-column align-self-end me-3">
              <button class="btn btn-secondary text-light ms-2 mb-2 ngb-tooltip" id="insertTagButton">
                <span class="tooltiptext">Insert</span>
                <i class="fa fa-plus text-light c-pointer"></i>
              </button>
              <button class="btn btn-secondary ms-2 mb-2 text-white ngb-tooltip" id="changeSourceButton">
              ${tooltipButton}
                <i class="fa fa-file-lines text-white"></i>
              </button>
              <button type="submit" class="btn btn-primary bg-primary-clr ms-2 text-white ngb-tooltip" id="sendPromptButton">
                <span class="tooltiptext">Send</span>
                <i class="fa fa-paper-plane text-white"></i>
              </button>
            </div>`
}

export function renderSelectedTags(selectedNames, availableKeys) {
  const badgeWrapper = document.getElementById('tag-badge-wrapper');
  badgeWrapper.innerHTML = '';

  // Filter out duplicates (case-insensitive)
  const uniqueNames = [...new Set(
    selectedNames.map(name => name.toLowerCase())
  )].map(lowerName => 
    selectedNames.find(name => name.toLowerCase() === lowerName)
  );

  uniqueNames.forEach(name => {
    const badge = document.createElement('span');
    badge.className = 'badge rounded-pill border bg-white text-dark px-3 py-2 shadow-sm d-flex align-items-center badge-clickable';
    badge.style.cursor = 'pointer';
    badge.innerHTML = `${name} <i class="fa-solid fa-robot ms-2 text-muted" aria-label="AI Suggested"></i>`;

    badge.addEventListener('click', () => {
      const aiTag = availableKeys.find(
        mention => mention.AIFlag === 1 && mention.DisplayName.toLowerCase() === name.toLowerCase()
      );

      if (aiTag) {
        const appBody = document.getElementById('app-body');
        appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';

        generateCheckboxHistory(aiTag).then(html => {
          appBody.innerHTML = html;
        });
      }
    });

    badgeWrapper.appendChild(badge);
  });
}


export function applyThemeClasses(theme) {
  const isDark = theme === 'Dark';
  const isLight = theme === 'Light';

  const safeApplyClass = (selector, darkClasses, lightClasses) => {
    const elements = document.querySelectorAll(selector);
    const darkClassList = darkClasses.split(' ');
    const lightClassList = lightClasses.split(' ');

    elements.forEach(elem => {
      if (!elem) return;
      // Remove all related theme classes
      elem.classList.remove(...darkClassList);
      elem.classList.remove(...lightClassList);
      // Add only the relevant set
      if (isDark) elem.classList.add(...darkClassList);
      if (isLight) elem.classList.add(...lightClassList);
    });
  };

  // Now use it for different elements
  safeApplyClass('#app-body', 'bg-dark text-light', 'bg-white text-dark');
  safeApplyClass('#search-box', 'bg-secondary text-light border-0', 'bg-white text-dark border');
  safeApplyClass('.dropdown-menu', 'bg-dark text-light border-light', 'bg-white text-dark border');
  safeApplyClass('.list-group-item', 'bg-dark text-light', 'bg-white text-dark');
  safeApplyClass('.dropdown-toggle', 'bg-dark text-light border-0', 'bg-white text-dark border');
  safeApplyClass('.dropdown-item', 'bg-dark text-light', 'bg-white text-dark');
  // container for the suggestion list
  safeApplyClass(
    '.list-group-item-action',
    'bg-dark text-light list-hover-dark',
    'bg-light text-dark list-hover-light'
  );

safeApplyClass('#close-ai-window', 'fa-solid fa-circle-xmark bg-dark text-light', 'fa-solid fa-circle-xmark bg-light text-dark');
safeApplyClass('#chatInput', 'bg-secondary text-light', 'bg-white text-dark');
safeApplyClass('.prompt-text', 'bg-secondary text-light', 'bg-white text-dark');


}

export function swicthThemeIcon(){
  const themeToggle = document.getElementById('theme-toggle');
  const icon = themeToggle.querySelector('i');

  if (theme === 'Dark') {
    icon.classList.remove('fa-moon');
    icon.classList.add('fa-sun');
  } else if (theme === 'Light') {
    icon.classList.remove('fa-sun');
    icon.classList.add('fa-moon');
  }
}