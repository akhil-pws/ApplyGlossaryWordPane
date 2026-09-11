import { toaster, Confirmationpopup } from "../components/bodyelements";
import { generateCheckboxHistory } from "./home";
import { wordTableStyleLocales } from "../components/tablestyles";
import { StoreService } from "../services/store.service";
import { DocStorage } from "../utils/doc-storage";
import { faFolderGear, faMicrochipAi, getIconSvg } from "../utils/fontawesome-icons";

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

  // Apply configured default text style if it's the normal style
  try {
    const store = StoreService.getInstance();
    if (builtInStyle === Word.BuiltInStyleName.normal) {
      const styleName = store.defaultTextStyle || "Normal";
      paragraph.style = styleName;
    } else {
      paragraph.styleBuiltIn = builtInStyle;
    }

    // Apply customized style properties if configured for normal text
    if (builtInStyle === Word.BuiltInStyleName.normal && store.customizedTextStyle && store.customizedTextStyle.properties) {
      const props = store.customizedTextStyle.properties;
      paragraph.font.bold = props.bold;
      paragraph.font.italic = props.italic;
      paragraph.font.underline = props.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
      if (props.fontFamily) {
        paragraph.font.name = props.fontFamily;
      }
      if (props.size) {
        const numericSize = parseFloat(props.size);
        if (!isNaN(numericSize)) {
          paragraph.font.size = numericSize;
        }
      }
      if (props.fontColor) {
        paragraph.font.color = props.fontColor;
      }
      if (props.backgroundColor && props.backgroundColor !== "transparent") {
        try {
          paragraph.font.highlightColor = props.backgroundColor;
        } catch (e) {}
      }
    }
  } catch (e) {
    console.error("Error setting paragraph style, falling back to normal:", e);
    try {
      paragraph.styleBuiltIn = Word.BuiltInStyleName.normal;
    } catch (err) {}
  }

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

export function confirmSwitchChatHistory(onConfirm: () => void, onCancel?: () => void) {
  const chatInput = document.getElementById("chatInput") as HTMLTextAreaElement;
  if (chatInput && chatInput.value.trim().length > 0) {
    const container = document.getElementById('confirmation-popup');
    if (container) {
      container.innerHTML = Confirmationpopup('Are you sure you want to discard changes and proceed?');

      setTimeout(() => {
        const cancelBtn = document.getElementById('confirmation-popup-cancel');
        const confirmBtn = document.getElementById('confirmation-popup-confirm');

        cancelBtn?.addEventListener('click', () => {
          container.innerHTML = '';
          if (onCancel) onCancel();
        });

        confirmBtn?.addEventListener('click', () => {
          container.innerHTML = '';
          onConfirm();
        });
      }, 0);
    } else {
      if (confirm('You have unsaved changes in your chat input. Do you want to switch?')) {
        onConfirm();
      } else {
        if (onCancel) onCancel();
      }
    }
  } else {
    onConfirm();
  }
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

  let normalizedData = Array.isArray(jsonData) ? jsonData : [jsonData];

  // Narrative summary tables come back as [{ Label: Value }, ...].
  // Render them as key/value rows instead of one very wide row.
  if (normalizedData.length > 1 && normalizedData.every(item =>
    typeof item === 'object' && item !== null && Object.keys(item).length === 1
  )) {
    const keyValueRows = normalizedData
      .map(item => {
        const [key, value] = Object.entries(flattenObject(item))[0] || ["", ""];
        return { key, value };
      })
      .filter(({ key, value }) => {
        const cellValue = value === undefined || value === null ? "" : String(value);
        return key.trim() !== "" || cellValue.trim() !== "";
      });

    if (keyValueRows.length === 0) return '<p>No data available</p>';

    let table = '<table border="1" cellspacing="0" cellpadding="5">';
    keyValueRows.forEach(({ key, value }) => {
      const cellValue = value === undefined || value === null ? "" : value;
      table += `<tr><th>${key}</th><td>${cellValue}</td></tr>`;
    });
    table += '</table>';
    return table;
  }

  normalizedData.forEach(item => {
    let flattenedItem = flattenObject(item);
    Object.keys(flattenedItem).forEach(key => headers.add(key));
    rows.push(flattenedItem);
  });

  let table = '<table border="1" cellspacing="0" cellpadding="5">';
  table += '<tr>' + [...headers].map(header => `<th>${header}</th>`).join('') + '</tr>';
  rows.forEach(row => {
    table += '<tr>' + [...headers].map(header => `<td>${(row[header] === undefined || row[header] === null) ? "" : row[header]}</td>`).join('') + '</tr>';
  });

  table += '</table>';
  return table;
}



function formatChatDate(dateStr: any): string {
  if (!dateStr) return '';
  try {
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return dateStr;
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const yyyy = date.getFullYear();
    const mmm = months[date.getMonth()];
    const dd = String(date.getDate()).padStart(2, '0');
    let hours = date.getHours();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12; // the hour '0' should be '12'
    const hh = String(hours).padStart(2, '0');
    const min = String(date.getMinutes()).padStart(2, '0');
    
    // Calculate GMT offset
    const offset = -date.getTimezoneOffset();
    const diff = offset >= 0 ? '+' : '-';
    const absOffset = Math.abs(offset);
    const offsetHours = Math.floor(absOffset / 60);
    const offsetMinutes = absOffset % 60;
    const pad = (num: number) => String(num).padStart(2, '0');
    const gmtStr = `(GMT${diff}${pad(offsetHours)}:${pad(offsetMinutes)})`;

    return `${dd} ${mmm} ${yyyy} ${hh}:${min} ${ampm} ${gmtStr}`;
  } catch (e) {
    return dateStr;
  }
}

export function generateChatHistoryHtml(chatList: any[]): string {
  const store = StoreService.getInstance();
  const promptclass = store.theme === 'Dark' ? 'bg-secondary text-light' : 'bg-white text-dark';
  const globalPromptUpdate = store.UserRole.UserRoleEntityAccessList.find(
    (item: any) => item.UserRoleEntity === 'Global Prompt Update'
  );

  return chatList.map((chat, index) => {
    const includeSaveIcon = globalPromptUpdate?.UserRoleAccessID === 3;
    const includeReferenceIcon = true;
    const formattedDate = formatChatDate(chat.CreatedDate);

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
              ${includeSaveIcon ? `<i class="fa fa-save text-secondary c-pointer mb-2" title="Save Prompt" id="savePrompt-${index}"></i>` : ''}
              <div class="ngb-tooltip d-inline-block">
                <span class="tooltiptext info-tooltip-text">${chat.DocumentInstruction ? `<strong>Document Instruction:</strong> ${chat.DocumentInstruction}<br>` : ''}<strong>Created By:</strong> ${chat.CreatedByName || 'Unknown User'}<br><strong>Date:</strong> ${formattedDate || 'N/A'}</span>
                <i class="fa-solid fa-circle-info text-secondary c-pointer" id="infoPrompt-${index}"></i>
              </div>
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
        ? getIconSvg(faFolderGear, 'text-secondary c-pointer ms-2', '', `title="Open Reference" id="openRefferance-${index}"`)
        : ''}
            </div>
          </span>
        </div>
      </div>`;
  }).join('');
}



export function chatfooter(tag: any) {
  const store = StoreService.getInstance();
  const promptclass = store.theme === 'Dark' ? 'bg-secondary text-light' : 'bg-white text-dark';
  const tooltipButton = tag.Sources && tag.Sources.length > 0
    ? `  <span class="tooltiptext">${tag.Sources}</span>`
    : '<span class="tooltiptext">Source</span>';
  return ` <textarea class="form-control ${promptclass}"
                      rows="7"
                      id="chatInput"
                      ></textarea>
            <div id="mention-dropdown" class="dropdown-menu"></div>
            <div class="d-flex flex-column align-self-end me-3">
              <button class="btn btn-secondary text-light ms-2 mb-2 ngb-tooltip" id="insertTagButton">
                <span class="tooltiptext">Insert</span>
                <i class="fa fa-plus text-light c-pointer"></i>
              </button>
              <button class="btn btn-secondary text-light ms-2 mb-2 ngb-tooltip" id="promptBuilderButton">
                <span class="tooltiptext">Prompt Builder</span>
                <i class="fa fa-keyboard text-light c-pointer"></i>
              </button>
              <button class="btn btn-secondary text-light ms-2 mb-2 ngb-tooltip" id="changeSourceButton">
                ${tooltipButton}
                <i class="fa fa-file-lines text-light c-pointer"></i>
              </button>

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

    const store = StoreService.getInstance();

    uniqueNames.forEach(name => {
      if (store.mode === 'Summary') {
        let summaryTag;
        if (/^SM\d+$/i.test(name)) {
          summaryTag = store.summaryTagList?.find(
            k => `sm${k.ID || k.ReportHeadSummaryTagID}`.toLowerCase() === name.toLowerCase()
          );
        } else {
          summaryTag = store.summaryTagList?.find(
            k => k.Name?.toLowerCase() === name.toLowerCase()
          );
        }

        if (summaryTag?.Name) {
          const badge = document.createElement('span');
          badge.className = 'badge rounded-pill border bg-white text-dark px-3 py-2 shadow-sm d-flex align-items-center badge-clickable';
          badge.style.cursor = 'pointer';
          badge.innerHTML = `${summaryTag.Name} <i class="fa-solid fa-wand-magic-sparkles ms-2 text-muted" aria-label="Summary Tag"></i>`;
          badge.addEventListener('click', async () => {
            const tagId = summaryTag.ID || summaryTag.ReportHeadSummaryTagID;
            if (store.currentChatTagId !== -1 && store.currentChatTagId !== undefined && store.currentChatTagId !== null && String(store.currentChatTagId) === String(tagId)) {
              return;
            }
            confirmSwitchChatHistory(async () => {
              const targetChatId = await selectMatchingBookmarkFromSelection(name);

              if (summaryTag) {
                const appBody = document.getElementById('app-body');
                appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';

                generateCheckboxHistory(summaryTag, "Summary", targetChatId || undefined).then(html => {
                  appBody.innerHTML = html;
                });
              }
            });
          });
          badgeWrapper.appendChild(badge);
        }
      } else {
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
          badge.innerHTML = `${aiTag.DisplayName} ${getIconSvg(faMicrochipAi, 'ms-2 text-muted', '', 'aria-label="AI Suggested"')}`;
          badge.addEventListener('click', async () => {
            const tagId = aiTag.ID || aiTag.ReportHeadSummaryTagID;
            if (store.currentChatTagId !== -1 && store.currentChatTagId !== undefined && store.currentChatTagId !== null && String(store.currentChatTagId) === String(tagId)) {
              return;
            }
            confirmSwitchChatHistory(async () => {
              const targetChatId = await selectMatchingBookmarkFromSelection(name);

              if (aiTag) {
                const appBody = document.getElementById('app-body');
                appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';

                generateCheckboxHistory(aiTag, "AITag", targetChatId || undefined).then(html => {
                  appBody.innerHTML = html;
                });
              }
            });
          });
          badgeWrapper.appendChild(badge);
        }
      }
    });
  }
}



// applyThemeClasses and swicthThemeIcon moved to UIService

const modeIconMap = {
  Home: "fa-home",
  Summary: "fa-wand-magic-sparkles",
  Review: "fa-magnifying-glass",
} as const;

export function switchModeIcon() {
  const icon = document.getElementById("modeDropdown");
  if (!icon) return;

  const store = StoreService.getInstance();
  icon.classList.remove(...Object.values(modeIconMap));
  icon.classList.add(modeIconMap[store.mode]);

  DocStorage.setItem("mode", store.mode);
}

/**
 * Extracts chat/response ID from bookmark format like ID455997_Split_20260903_105201_101 or SM12_Split_20260903_105201_201
 */
export function extractChatIdFromBookmark(bookmarkName: string): string | null {
  if (!bookmarkName || !bookmarkName.includes('_Split_')) {
    return null;
  }
  const splitParts = bookmarkName.split('_Split_');
  if (splitParts.length > 1) {
    const afterSplit = splitParts[1];
    const parts = afterSplit.split('_');
    // Format: [YYYYMMDD, HHMMSS, chatId] or [YYYYMMDD, HHMMSS]
    if (parts.length >= 3) {
      return parts.slice(2).join('_');
    }
  }
  return null;
}

export async function selectMatchingBookmarkFromSelection(displayName): Promise<string | null> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    const bookmarks = selection.getBookmarks(); // ClientResult<string[]>
    await context.sync();

    const targetBookmarkName = (bookmarks.value || []).find(bookmark => {
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
      return extractChatIdFromBookmark(targetBookmarkName);
    }
    return null;
  });
}


export function applyCustomTextStyleToCell(cell: any, store: any) {
  try {
    const styleName = store.defaultTextStyle || "Normal";
    cell.body.style = styleName;
  } catch (e) {
    console.error("Error setting cell body style:", e);
  }

  if (store.customizedTextStyle && store.customizedTextStyle.properties) {
    const props = store.customizedTextStyle.properties;
    const font = cell.body.font;
    font.bold = props.bold;
    font.italic = props.italic;
    font.underline = props.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
    if (props.fontFamily) {
      font.name = props.fontFamily;
    }
    if (props.size) {
      const numericSize = parseFloat(props.size);
      if (!isNaN(numericSize)) {
        font.size = numericSize;
      }
    }
    if (props.fontColor) {
      font.color = props.fontColor;
    }
    if (props.backgroundColor && props.backgroundColor !== "transparent") {
      try {
        font.highlightColor = props.backgroundColor;
      } catch (e) {}
    }
  }
}

export async function colorTable(table: any, rows: any, context: any, isReversed: boolean = false) {
  const store = StoreService.getInstance();

  // ------------------------------------------------------------
  // 1) Copy cell values DOM -> Word + MERGE FIRST COLUMN GROUPS
  // Skip this phase if reversed, as the grid population is handled externally 
  // and merging is disabled for transposed tables.
  // ------------------------------------------------------------
  if (!isReversed) {
    let lastParamRowIndex = -1; // last row index where 1st col had value

    rows.forEach((row, rowIndex) => {
      const cells = Array.from(row.querySelectorAll("td, th"));
      let cellIndex = 0;

      let firstColText = "";

      cells.forEach((cell) => {
        const text = cell.innerText.trim();

        // capture first column text
        if (cellIndex === 0) {
          firstColText = text;
        }

        const tableCell = table.getCell(rowIndex, cellIndex);
        tableCell.value = text;
        applyCustomTextStyleToCell(tableCell, store);
        cellIndex++;
      });

      // ✅ Merge first column when empty (skip header row)
      if (rowIndex > 0) {
        if (firstColText) {
          // new group starts
          lastParamRowIndex = rowIndex;
        } else {
          // empty first col = merge with last non-empty parameter row
          if (lastParamRowIndex !== -1) {
            const topCell = table.getCell(lastParamRowIndex, 0);
            const bottomCell = table.getCell(rowIndex, 0);

            topCell.merge(bottomCell);

            // ✅ center align merged parameter cell
            try {
              topCell.verticalAlignment = Word.VerticalAlignment.center;
              topCell.body.paragraphs.getFirst().alignment = Word.Alignment.center;
            } catch (e) {
              // ignore
            }
          }
        }
      }
    });

    await context.sync();
  }

  // ------------------------------------------------------------
  // 2) Load rows for formatting
  // ------------------------------------------------------------
  table.rows.load("items");
  await context.sync();

  table.rows.items.forEach(row => row.cells.load("items"));
  await context.sync();

  table.rows.items.forEach((row) => {
    row.cells.items.forEach((cell) => {
      applyCustomTextStyleToCell(cell, store);
    });
  });

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

    if (store.colorPallete.IsHeaderBold && rowIndex === 0) weight = "Bold";
    if (store.colorPallete.IsSideHeaderBold && cellIndex === 0) weight = "Bold";

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
  const base = store.tableStyle.split(" - ")[0].trim();

  // ------------------------------------------------------------
  // Plain Table 3
  // ------------------------------------------------------------
  if (base === "Plain Table 3") {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();

    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = store.colorPallete.Primary;

        if (rowIndex === 0 || cellIndex === 0) {
          bgColor = store.colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1
            ? store.colorPallete.Primary
            : store.colorPallete.Secondary;
        }

        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
  }

  // ------------------------------------------------------------
  // Plain Table 2
  // ------------------------------------------------------------
  else if (base === "Plain Table 2") {

    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();

    const rowCount = table.rows.items.length;
    const firstGapIndex = 0;             // gap between row 0 and row 1
    const lastGapIndex = rowCount - 1;   // gap between last-2 and last row

    // STEP 2 — Turn OFF all gaps EXCEPT first + last
    table.rows.items.forEach((row, rowIndex) => {
      // rowIndex = gap BELOW this row
      if (rowIndex !== firstGapIndex && rowIndex !== lastGapIndex) {
        const bottom = row.getBorder(Word.BorderLocation.bottom);
        bottom.type = Word.BorderType.none;
      }
    });

    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = store.colorPallete.Primary;

        if (rowIndex === 0) {
          bgColor = store.colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1
            ? store.colorPallete.Primary
            : store.colorPallete.Secondary;
        }

        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });

    await context.sync();
  }

  // ------------------------------------------------------------
  // Plain Table 5
  // ------------------------------------------------------------
  else if (base === "Plain Table 5") {
    table.getBorder(Word.BorderLocation.insideVertical).type = Word.BorderType.none;

    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();

    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = store.colorPallete.Primary;

        if (rowIndex === 0) {
          bgColor = store.colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1
            ? store.colorPallete.Primary
            : store.colorPallete.Secondary;
        }

        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
  }

  // ------------------------------------------------------------
  // Table Grid
  // ------------------------------------------------------------
  else if (base === "Table Grid") {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, i) => {
      const bg = i % 2 === 0 ? store.colorPallete.Header : store.colorPallete.Primary;
      applyColor(row, bg);
      row.cells.items.forEach((cell, cellIndex) => {
        applyBoldIfNeeded(cell, i, cellIndex);
      });

    });
  } else if (base === "Table Grid 2") {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();

    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = store.colorPallete.Primary;

        if (cellIndex === 0) {
          // Side header only
          bgColor = store.colorPallete.Header;
        } else {
          // Rest of the cells follow alternating Primary/Secondary logic
          bgColor = rowIndex % 2 === 0
            ? store.colorPallete.Primary
            : store.colorPallete.Secondary;
        }

        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
  }
  else if (base === "Table Grid 1") {
    table.rows.items.forEach(r => r.cells.load("items"));
    await context.sync();

    table.rows.items.forEach((row, i) => {
      const bg = i % 2 === 0 ? store.colorPallete.Header : store.colorPallete.Primary;
      applyColor(row, bg);

      row.cells.items.forEach((cell, cellIndex) => {
        applyBoldIfNeeded(cell, i, cellIndex);
      });
    });

    // ✅ ONE full width double border under header row
    const headerRow = table.rows.items[0];
    const bottom = headerRow.getBorder(Word.BorderLocation.bottom);
    bottom.type = Word.BorderType.double;
    bottom.width = 0.5;
    bottom.color = "000000";

    await context.sync();
  }


  // ------------------------------------------------------------
  // Grid Table 4
  // ------------------------------------------------------------
  else if (base.startsWith("Grid Table 4")) {
    const headerRow = table.rows.items[0];
    applyColor(headerRow, store.colorPallete.Header);
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, i) => {
      if (i > 0) {
        const bg = i % 2 === 0 ? store.colorPallete.Secondary : store.colorPallete.Primary;
        applyColor(row, bg);
      }

      row.cells.items.forEach((cell, cellIndex) => {
        applyBoldIfNeeded(cell, i, cellIndex);
      });
    });
  }

  // ------------------------------------------------------------
  // Grid Table 5 Dark
  // ------------------------------------------------------------
  else if (base.startsWith("Grid Table 5 Dark")) {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();

    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = store.colorPallete.Primary;

        if (rowIndex === 0 || cellIndex === 0) {
          bgColor = store.colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1
            ? store.colorPallete.Primary
            : store.colorPallete.Secondary;
        }

        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
  }

  // ------------------------------------------------------------
  // List Table 3
  // ------------------------------------------------------------
  else if (base.startsWith("List Table 3")) {
    const headerRow = table.rows.items[0];
    applyColor(headerRow, store.colorPallete.Header);
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();
    table.rows.items.forEach((row, i) => {
      if (i > 0) applyColor(row, store.colorPallete.Primary);

      row.cells.items.forEach((cell, cellIndex) => {
        applyBoldIfNeeded(cell, i, cellIndex);
      });
    });
  }

  // ------------------------------------------------------------
  // List Table 2
  // ------------------------------------------------------------
  else if (base.startsWith("List Table 2")) {
    table.rows.items.forEach(row => row.cells.load("items"));
    await context.sync();

    table.rows.items.forEach((row, rowIndex) => {
      row.cells.items.forEach((cell, cellIndex) => {
        let bgColor = store.colorPallete.Primary;

        if (rowIndex === 0) {
          bgColor = store.colorPallete.Header;
        } else {
          bgColor = rowIndex % 2 === 1
            ? store.colorPallete.Primary
            : store.colorPallete.Secondary;
        }

        applyColor(cell, bgColor);
        applyBoldIfNeeded(cell, rowIndex, cellIndex);
      });
    });
  }

  // ------------------------------------------------------------
  // default
  // ------------------------------------------------------------
  else {
    table.rows.items.forEach((row) => applyColor(row, store.colorPallete.Primary));
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

/**
 * Converts an array of HTML table rows into a 2D string grid.
 * Correctly accounts for colspan and rowspan by filling the grid cells.
 */
export function parseHtmlTableToGrid(rows: HTMLElement[]): string[][] {
  if (rows.length === 0) return [];

  // 1. Calculate max columns considering colspans
  let maxCols = 0;
  rows.forEach((row) => {
    let colsInRow = 0;
    Array.from(row.querySelectorAll("td, th")).forEach((cell) => {
      colsInRow += parseInt(cell.getAttribute("colspan") || "1", 10);
    });
    if (colsInRow > maxCols) maxCols = colsInRow;
  });

  const rowCount = rows.length;
  const grid: string[][] = Array.from({ length: rowCount }, () => new Array(maxCols).fill(""));
  const occupied = Array.from({ length: rowCount }, () => new Array(maxCols).fill(false));

  rows.forEach((row, rowIndex) => {
    const cells = Array.from(row.querySelectorAll("td, th"));
    let colIndex = 0;

    cells.forEach((cell) => {
      // Find the next available column in the grid
      while (colIndex < maxCols && occupied[rowIndex][colIndex]) {
        colIndex++;
      }

      if (colIndex >= maxCols) return;

      const cellText = Array.from(cell.childNodes)
        .map((node) => {
          if (node.nodeType === Node.TEXT_NODE) {
            return node.textContent?.trim() || "";
          } else if (node.nodeType === Node.ELEMENT_NODE) {
            return (node as HTMLElement).innerText.trim();
          }
          return "";
        })
        .filter((text) => text.length > 0)
        .join(" ");

      const colspan = parseInt(cell.getAttribute("colspan") || "1", 10);
      const rowspan = parseInt(cell.getAttribute("rowspan") || "1", 10);

      // Fill the grid cells covered by this cell's rowspan/colspan
      for (let r = 0; r < rowspan; r++) {
        for (let c = 0; c < colspan; c++) {
          const targetRow = rowIndex + r;
          const targetCol = colIndex + c;

          if (targetRow < rowCount && targetCol < maxCols) {
            // Only put text in the primary cell; others stay empty strings for merging logic
            grid[targetRow][targetCol] = (r === 0 && c === 0) ? cellText : "";
            occupied[targetRow][targetCol] = true;
          }
        }
      }

      colIndex += colspan;
    });
  });

  return grid;
}

/**
 * Flips the 2D grid (rows become columns and vice versa).
 */
export function transposeGrid(grid: string[][]): string[][] {
  if (grid.length === 0) return [];
  const rows = grid.length;
  const cols = grid[0].length;
  const transposed = Array.from({ length: cols }, () => new Array(rows).fill(""));

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      transposed[c][r] = grid[r][c];
    }
  }

  return transposed;
}


type TableCase = "CASE_1" | "CASE_2" | "CASE_3" | "UNKNOWN";

export function detectTableCase(data: any[][]): TableCase {
  if (!Array.isArray(data) || data.length === 0) return "UNKNOWN";

  // CASE 1 → every row has exactly 2 elements (key-value)
  if (data.every(row => Array.isArray(row) && row.length === 2)) {
    return "CASE_1";
  }

  // Must have at least header
  if (!Array.isArray(data[0])) return "UNKNOWN";

  const headerLength = data[0].length;

  // all rows same length as header
  const validTable = data.every(row => Array.isArray(row) && row.length === headerLength);

  if (!validTable) return "UNKNOWN";

  // CASE 2 → header + single row
  if (data.length === 2) return "CASE_2";

  // CASE 3 → header + multiple rows
  if (data.length > 2) return "CASE_3";

  return "UNKNOWN";
}
