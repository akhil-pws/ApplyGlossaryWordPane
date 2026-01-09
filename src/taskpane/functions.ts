import { toaster } from "./components/bodyelements";
import { generateCheckboxHistory } from "./components/home";
import { wordTableStyleLocales } from "./components/tablestyles";
import { colorPallete, tableStyle, theme, UserRole } from "./taskpane";

export function insertLineWithHeadingStyle(
  paragraph: Word.Paragraph,
  line: string
) {
  let builtInStyle: Word.BuiltInStyleName = Word.BuiltInStyleName.normal;
  let text = line;

  if (line.startsWith("###### ")) {
    builtInStyle = Word.BuiltInStyleName.heading6;
    text = line.substring(7).trim();
  } else if (line.startsWith("##### ")) {
    builtInStyle = Word.BuiltInStyleName.heading5;
    text = line.substring(6).trim();
  } else if (line.startsWith("#### ")) {
    builtInStyle = Word.BuiltInStyleName.heading4;
    text = line.substring(5).trim();
  } else if (line.startsWith("### ")) {
    builtInStyle = Word.BuiltInStyleName.heading3;
    text = line.substring(4).trim();
  } else if (line.startsWith("## ")) {
    builtInStyle = Word.BuiltInStyleName.heading2;
    text = line.substring(3).trim();
  } else if (line.startsWith("# ")) {
    builtInStyle = Word.BuiltInStyleName.heading1;
    text = line.substring(2).trim();
  }

  // ✅ THIS is the important line
  paragraph.styleBuiltIn = builtInStyle;

  const regex = /(\*\*(.+?)\*\*)|(\*(.+?)\*)|(_(.+?)_)/g;
  let lastIndex = 0;
  let match;

  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      paragraph.insertText(
        text.substring(lastIndex, match.index),
        Word.InsertLocation.end
      );
    }

    let content = "";
    let bold = false;
    let italic = false;
    let underline = false;

    if (match[1]) {
      content = match[2];
      bold = true;
    } else if (match[3]) {
      content = match[4];
      italic = true;
    } else if (match[5]) {
      content = match[6];
      underline = true;
    }

    const r = paragraph.insertText(content, Word.InsertLocation.end);
    r.font.bold = bold;
    r.font.italic = italic;
    r.font.underline = underline
      ? Word.UnderlineType.single
      : Word.UnderlineType.none;

    lastIndex = regex.lastIndex;
  }

  if (lastIndex < text.length) {
    paragraph.insertText(
      text.substring(lastIndex),
      Word.InsertLocation.end
    );
  }
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



export function copyText(text: string) {
  // Copy text to clipboard logic
  const tempTextArea = document.createElement('textarea');
  tempTextArea.value = text;
  document.body.appendChild(tempTextArea);
  tempTextArea.select();
  document.execCommand('copy');
  document.body.removeChild(tempTextArea);
  toaster('Copied to clipboard successfully!', 'success')

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



export function generateChatHistoryHtml(chatList: any[]): string {
  const promptclass = theme === 'Dark' ? 'bg-secondary text-light' : 'bg-white text-dark';
  const globalPromptUpdate = UserRole.UserRoleEntityAccessList.find(
    (item: any) => item.UserRoleEntity === 'Global Prompt Update'
  );

  return chatList.map((chat, index) => {
    const includeSaveIcon = globalPromptUpdate?.UserRoleAccessID === 3;
    const includeReferenceIcon = chat.Evidences && chat.Evidences.length > 0;

    return `
      <div class="row chat-entry m-0 p-0">
        <div class="col-md-12 mt-2 p-2">
          <div class="d-flex justify-content-between align-items-start">
            <!-- Prompt Box -->
            <div class="form-control h-34 d-flex align-items-center dynamic-height prompt-text ${promptclass}" style="width: 95%;">
              ${chat.Prompt}
            </div>

            <!-- Icons Stack -->
            <div class="d-flex flex-column align-items-center ms-2">
              <i class="fa fa-copy text-secondary c-pointer mb-2" title="Copy Prompt" id="copyPrompt-${index}"></i>
              ${includeSaveIcon ? `<i class="fa fa-save text-secondary c-pointer" title="Save Prompt" id="savePrompt-${index}"></i>` : ''}
            </div>
          </div>
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
              ${includeReferenceIcon
        ? `<i class="fa fa-folder-gear text-secondary c-pointer ms-2"
                    title="Open Reference"
                    id="openRefferance-${index}"></i>`
        : ''}
            </div>
          </span>
        </div>
      </div>`;
  }).join('');
}



export function chatfooter(tag: any) {
  const promptclass = theme === 'Dark' ? 'bg-secondary text-light' : 'bg-white text-dark';
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
             <div class="ngb-tooltip d-inline-block">
                 ${tooltipButton}
                <button class="btn btn-secondary ms-2 mb-2 text-white" id="changeSourceButton">
                  <i class="fa fa-file-lines text-white"></i>
                </button>
              </div>

              <button type="submit" class="btn btn-primary bg-primary-clr ms-2 text-white ngb-tooltip" id="sendPromptButton">
                <span class="tooltiptext">Send</span>
                <i class="fa fa-paper-plane text-white"></i>
              </button>
            </div>`
}

export function renderSelectedTags(selectedNames, availableKeys) {
  const badgeWrapper = document.getElementById('tag-badge-wrapper');
  if (badgeWrapper) {
    badgeWrapper.innerHTML = '';
    // Filter out duplicates (case-insensitive)
    const uniqueNames = [...new Set(
      selectedNames.map(name => name.toLowerCase())
    )].map(lowerName =>
      selectedNames.find(name => name.toLowerCase() === lowerName)
    );


    uniqueNames.forEach(name => {
      let aiTag;

      if (/^ID\d+$/i.test(name)) {
        aiTag = availableKeys.find(
          mention => mention.AIFlag === 1 && `id${mention.ID}`.toLowerCase() === name.toLowerCase()
        );
      } else {
        aiTag = availableKeys.find(
          mention => mention.AIFlag === 1 && mention.DisplayName.toLowerCase() === name.toLowerCase()
        );
      }
      if (aiTag?.DisplayName) {

        const badge = document.createElement('span');
        badge.className = 'badge rounded-pill border bg-white text-dark px-3 py-2 shadow-sm d-flex align-items-center badge-clickable';
        badge.style.cursor = 'pointer';
        badge.innerHTML = `${aiTag.DisplayName} <i class="fa-solid fa-microchip-ai ms-2 text-muted" aria-label="AI Suggested"></i>`;
        badge.addEventListener('click', async () => {
          await selectMatchingBookmarkFromSelection(name);

          if (aiTag) {
            const appBody = document.getElementById('app-body');
            appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';

            generateCheckboxHistory(aiTag).then(html => {
              appBody.innerHTML = html;
            });
          }
        });
        badgeWrapper.appendChild(badge);
      }
    });
  }

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

export function swicthThemeIcon() {
  const themeToggle = document.getElementById('theme-toggle');
  const icon = themeToggle.querySelector('i');

  if (theme === 'Dark') {
    icon.classList.remove('fa-moon');
    icon.classList.add('fa-sun');
    sessionStorage.setItem('theme', 'Dark');
  } else if (theme === 'Light') {
    icon.classList.remove('fa-sun');
    icon.classList.add('fa-moon');
    sessionStorage.setItem('theme', 'Light');
  }
}

export async function selectMatchingBookmarkFromSelection(displayName) {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    const bookmarks = selection.getBookmarks(); // ClientResult<string[]>
    await context.sync();

    const targetBookmarkName = bookmarks.value.find(bookmark => {
      const cleanName = bookmark.split('_Split_')[0].replace(/_/g, ' ');
      return cleanName.toLowerCase() === displayName.toLowerCase();
    });

    if (targetBookmarkName) {
      const range = context.document.getBookmarkRangeOrNullObject(targetBookmarkName);
      range.load('isNullObject');
      await context.sync();

      if (!range.isNullObject) {
        range.select(); // Select the entire bookmark
      }
    }
  });
}

export async function colorTable(table: any, rows: any, context: any) {
  // Copy cell values from DOM table to Word table
  rows.forEach((row, rowIndex) => {
    const cells = Array.from(row.querySelectorAll("td, th"));
    let cellIndex = 0;

    cells.forEach((cell) => {
      const text = cell.innerText.trim();
      table.getCell(rowIndex, cellIndex).value = text;
      cellIndex++;
    });
  });
  await context.sync();

  // Load rows for formatting
  table.rows.load("items");
  await context.sync();

  // Helper to check if color is dark
  const isColorDark = (hexColor: string): boolean => {
    const hex = hexColor.replace("#", "");
    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return luminance < 0.5;
  };

  const applyBoldIfNeeded = (cell: any, rowIndex: number, cellIndex: number) => {
    let weight = "Normal";

    if (colorPallete.IsHeaderBold && rowIndex === 0) weight = "Bold";
    if (colorPallete.IsSideHeaderBold && cellIndex === 0) weight = "Bold";

    cell.body.font.bold = weight === "Bold";
  };

  // Helper to apply shading, font color, and border to a row or cell
  const applyColor = (cellOrRow: any, bgColor: string) => {
    cellOrRow.shadingColor = bgColor;
    try {
      cellOrRow.font.color = isColorDark(bgColor) ? "#FFFFFF" : "#000000";

      // Apply thin light grey border for all sides
      const borderColor = "#D3D3D3"; // light grey
      if (cellOrRow.getBorder) {
        ["Top", "Bottom", "Left", "Right"].forEach((side) => {
          cellOrRow.getBorder(side).type = "Single";
          cellOrRow.getBorder(side).color = borderColor;
          cellOrRow.getBorder(side).width = 1; // 1pt thin
        });
      }
    } catch (e) {
      // row objects might not have font directly
    }
  };

  // Determine base table type
  const base = tableStyle.split(" - ")[0].trim();
  if (base === "Plain Table 3") {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();

    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = colorPallete.Primary;

        if (rowIndex === 0 || cellIndex === 0) {
          bgColor = colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1
            ? colorPallete.Primary
            : colorPallete.Secondary;
        }

        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);

      });
    });
  }

  else if (base === "Plain Table 2") {

    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();

    const rowCount = table.rows.items.length;
    const firstGapIndex = 0;             // gap between row 0 and row 1
    const lastGapIndex = rowCount - 1;   // gap between last-2 and last row

    // --------------------------------------------
    // STEP 2 — Turn OFF all gaps EXCEPT first + last
    // --------------------------------------------
    table.rows.items.forEach((row, rowIndex) => {
      // rowIndex = gap BELOW this row
      if (rowIndex !== firstGapIndex && rowIndex !== lastGapIndex) {
        const bottom = row.getBorder(Word.BorderLocation.bottom);
        bottom.type = Word.BorderType.none;
      }
    });

    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = colorPallete.Primary;

        if (rowIndex === 0) {
          bgColor = colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1
            ? colorPallete.Primary
            : colorPallete.Secondary;
        }

        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });

    await context.sync();
  }
  else if (base === "Plain Table 5") {
    table.getBorder(Word.BorderLocation.insideVertical).type = Word.BorderType.none;

    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();

    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = colorPallete.Primary;

        if (rowIndex === 0) {
          bgColor = colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1
            ? colorPallete.Primary
            : colorPallete.Secondary;
        }

        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
  }


  // ------------------------------------------------------------
  // 3) Plain Table (your original logic — UNCHANGED)
  // ------------------------------------------------------------
  else if (base === "Table Grid") {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, i) => {
      const bg = i % 2 === 0 ? colorPallete.Header : colorPallete.Primary;
      applyColor(row, bg);
      row.cells.items.forEach((cell, cellIndex) => {
        applyBoldIfNeeded(cell, i, cellIndex);
      });
    });
  }

  // ------------------------------------------------------------
  // 4) Grid Table 4 (UNCHANGED)
  // ------------------------------------------------------------
  else if (base.startsWith("Grid Table 4")) {
    const headerRow = table.rows.items[0];
    applyColor(headerRow, colorPallete.Header);
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, i) => {
      if (i > 0) {
        const bg = i % 2 === 0 ? colorPallete.Secondary : colorPallete.Primary;
        applyColor(row, bg);
      }

      row.cells.items.forEach((cell, cellIndex) => {
        applyBoldIfNeeded(cell, i, cellIndex);
      });
    });
  }
  else if (base.startsWith("Grid Table 5 Dark")) {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();

    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = colorPallete.Primary;

        if (rowIndex === 0 || cellIndex === 0) {
          bgColor = colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1
            ? colorPallete.Primary
            : colorPallete.Secondary;
        }

        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
  }
  else if (base.startsWith("List Table 3")) {
    const headerRow = table.rows.items[0];
    applyColor(headerRow, colorPallete.Header);
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, i) => {
      if (i > 0) applyColor(row, colorPallete.Primary);

      row.cells.items.forEach((cell, cellIndex) => {
        applyBoldIfNeeded(cell, i, cellIndex);
      });
    });
  }

  else if (base.startsWith("List Table 2")) {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();

    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = colorPallete.Primary;

        if (rowIndex === 0) {
          bgColor = colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1
            ? colorPallete.Primary
            : colorPallete.Secondary;
        }

        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
  }
  else {
    table.rows.items.forEach((row) => applyColor(row, colorPallete.Primary));
  }

  await context.sync();
}

export function mapImagesToComponentObjects(input: any): any[] {
  if (!input) return [];

  // 1️⃣ Flatten ALL three arrays into a single list
  const flatImages = [
    ...(input.Flowchart || []),
    ...(input.Graph || []),
    ...(input.Image || [])
  ];

  // 2️⃣ Map to required structure
  return flatImages.map(img => ({
    Name: img.ImageName,
    DisplayName: img.ImageName,
    EditorValue: img.ImageData,
    UserValue: img.ImageData,
    ComponentKeyDataType: "IMAGE",
    AIFlag: 0,
    IsImage: true
  }));
}

export async function svgBase64ToPngBase64(svgBase64: string): Promise<string> {
  return new Promise((resolve, reject) => {
    try {
      const svgBlob = new Blob(
        [atob(svgBase64.split(',')[1])],
        { type: "image/svg+xml" }
      );
      const url = URL.createObjectURL(svgBlob);

      const img = new Image();
      img.crossOrigin = "anonymous";

      img.onload = function () {
        const canvas = document.createElement("canvas");
        canvas.width = img.width || 1200;
        canvas.height = img.height || 800;

        const ctx = canvas.getContext("2d");
        if (!ctx) return reject("Canvas unsupported.");

        ctx.drawImage(img, 0, 0);

        const pngBase64 = canvas.toDataURL("image/png").split(",")[1];
        resolve(pngBase64);

        URL.revokeObjectURL(url);
      };

      img.onerror = () => reject("SVG load failed.");
      img.src = url;

    } catch (err) {
      reject(err);
    }
  });
}

export function resolveWordTableStyle(englishStyle: string): string {
  const lang =
    Office.context.displayLanguage?.toLowerCase().slice(0, 2) || "en";

  return (
    wordTableStyleLocales[lang]?.[englishStyle] ??
    wordTableStyleLocales.en[englishStyle] ??
    'none'
  );
}
