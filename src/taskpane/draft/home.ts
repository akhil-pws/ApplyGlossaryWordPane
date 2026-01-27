import { getPromptTemplateById, updateGroupKey, updateAiHistory, updatePromptTemplate } from "./draft.api";
import { chatfooter, copyText, generateChatHistoryHtml, insertLineWithHeadingStyle, removeQuotes, switchToAddTag, updateEditorFinalTable, colorTable, svgBase64ToPngBase64, resolveWordTableStyle, renderSelectedTags } from "./draft-functions";
import { addGenAITags, applyTagFn, createMultiSelectDropdown, customizeTable, mentionDropdownFn } from "../taskpane";
import { StoreService } from "../services/store.service";
import { AIService } from "../services/ai.service";
import { Confirmationpopup, DataModalPopup, toaster } from "../components/bodyelements";
import { loadSummarypage } from "../summary/summary";
import { summaryService } from "../services/summary.service";
import { updateSummaryHistory } from "../summary/summary.api";

let preview = '';


export function loadHomepage(availableKeys) {
    const store = StoreService.getInstance();
    const searchBoxClass = store.theme === 'Dark' ? 'bg-secondary text-light' : 'bg-white text-dark';

    document.getElementById('app-body').innerHTML = `
    <div class="container pt-3">
        <div class="d-flex justify-content-end px-2">
            <div class="dropdown">
                <button class="btn btn-default dropdown-toggle" type="button" data-bs-toggle="dropdown" aria-expanded="false">
                    Action
                </button>
                <ul class="dropdown-menu">
                    <li>
                        <a class="dropdown-item" href="#" id="add-btn-tag">
                            <i class="fa fa-plus me-2" aria-hidden="true"></i> Add
                        </a>
                    </li>
                    <li>
                        <a class="dropdown-item" href="#" id="apply-btn-tag">
                            <i class="fa-solid fa-circle-check me-2"></i> Apply
                        </a>
                    </li>

                     <li>
                        <a class="dropdown-item" href="#" id="predefined-table">
                            <i class="fa fa-table me-2" aria-hidden="true"></i> Default Tables
                        </a>
                    </li>
                    <li>
                        <a class="dropdown-item" href="#" id="customized-table">
                            <i class="fa fa-brush me-2" aria-hidden="true"></i> Customized Tables
                        </a>
                    </li>
                </ul>
            </div>
        </div>

        <div class="form-group px-2 pt-2">
            <input type="text" id="search-box" class="form-control ${searchBoxClass}" placeholder="Search Tags..." autocomplete="off" />
        </div>

        <ul id="suggestion-list" class="list-group mt-2 px-2"></ul>
        
        <div id="tags-in-selected-text" class="mt-2 px-2 selected-text-box d-none">
            <label class="form-label mb-2 fw-bold">Tags in Selected Text</label>
            <div class="tag-panel d-flex flex-wrap gap-2" id="tag-badge-wrapper"></div>
        </div>
    </div>`;

    const searchBox = document.getElementById('search-box') as HTMLInputElement;
    const suggestionList = document.getElementById('suggestion-list');

    function updateSuggestions() {
        const searchTerm = searchBox.value.trim().toLowerCase();
        suggestionList.replaceChildren();

        if (searchTerm === '') {
            suggestionList.innerHTML = '';
            return;
        }

        const filteredMentions = availableKeys.filter(m =>
            m.DisplayName.toLowerCase().includes(searchTerm)
        );

        // Split groups
        const nonAITags = filteredMentions.filter(m => m.AIFlag === 0);
        const aiTags = filteredMentions.filter(m => m.AIFlag === 1);

        // Further split non-AI tags into: TEXT + IMAGE
        const propertiesTags = nonAITags.filter(m => m.ComponentKeyDataType === "TEXT" || m.ComponentKeyDataType === "TABLE");
        const imageTags = nonAITags.filter(m => m.ComponentKeyDataType === "IMAGE" && m.IsImage);

        const createSection = (labelText, mentions, isAISection = false, isImageSection = false) => {
            if (mentions.length === 0) return;

            const themeClasses = store.theme === 'Dark'
                ? { itemClass: 'bg-dark text-light list-hover-dark', labelClass: 'bg-dark text-light' }
                : { itemClass: 'bg-light text-dark list-hover-light', labelClass: 'bg-light text-dark' };

            const label = document.createElement('li');
            label.className = `list-group-item fw-bold text-secondary ${themeClasses.labelClass}`;
            label.textContent = labelText;
            suggestionList.appendChild(label);

            mentions.forEach(mention => {
                const listItem = document.createElement('li');
                listItem.className = `list-group-item list-group-item-action ${themeClasses.itemClass}`;

                // ICON LOGIC
                let icon = `<i class="fa-solid fa-layer-group text-muted me-2"></i>`; // default (TEXT)
                if (isAISection) icon = `<i class="fa-solid fa-microchip-ai text-muted me-2"></i>`;
                if (isImageSection) icon = `<i class="fa-solid fa-image text-muted me-2"></i>`;

                listItem.innerHTML = `${icon} ${mention.DisplayName}`;

                listItem.onclick = () => {
                    if (isAISection) {
                        const appBody = document.getElementById('app-body');
                        appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';
                        generateCheckboxHistory(mention, "AITag")
                            .catch(() => appBody.innerHTML = '<div class="text-danger p-2">Error loading data</div>')
                            .then(html => { appBody.innerHTML = html; });
                    } else {
                        // Properties + Images behave same
                        replaceMention(mention, mention.ComponentKeyDataType);
                        suggestionList.replaceChildren();
                    }
                };

                suggestionList.appendChild(listItem);
            });
        };

        // Render in desired order
        createSection('Properties', propertiesTags);
        createSection('AI Tags', aiTags, true);
        createSection('Images', imageTags, false, true);
    }

    if (store.selectedNames.length > 0) {
        const badgeWrapper = document.getElementById('tags-in-selected-text');
        badgeWrapper.classList.remove('d-none');
        badgeWrapper.classList.add('d-block');
        renderSelectedTags(store.selectedNames, availableKeys);
    }

    // Add input event listener to the search box
    let debounceTimeout;
    searchBox.addEventListener('input', () => {
        clearTimeout(debounceTimeout);
        debounceTimeout = setTimeout(updateSuggestions, 300); // Delay input handling by 300ms
    });

    document.getElementById('add-btn-tag').addEventListener('click', () => {
        if (!store.isPendingResponse) {
            addGenAITags();
        }
    });

    document.getElementById('customized-table').addEventListener('click', () => {
        if (!store.isPendingResponse) {
            customizeTable('Custom');
        }
    })

    document.getElementById('predefined-table').addEventListener('click', () => {
        if (!store.isPendingResponse) {
            customizeTable('Pre');
        }
    })

    document.getElementById('apply-btn-tag').addEventListener('click', () => {
        if (!store.isPendingResponse) {
            applyTagFn();
        }
    });
}



export async function replaceMention(word: any, type: any) {
    return Word.run(async (context) => {
        try {
            const selection = context.document.getSelection();
            await context.sync();

            if (!selection) {
                throw new Error('Selection is invalid or not found.');
            }

            let newSelection = selection;

            if (type === 'TABLE') {
                const parser = new DOMParser();
                const doc = parser.parseFromString(word.EditorValue, 'text/html');
                const bodyNodes = Array.from(doc.body.childNodes);

                await context.sync();

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
                            const store = StoreService.getInstance();
                            const resolvedTableStyle = resolveWordTableStyle(store.tableStyle);
                            if (resolvedTableStyle !== 'none') {
                                table.style = resolvedTableStyle;
                            }  // Apply built-in Word table style

                            await context.sync();
                            if (store.colorPallete.Customize) {
                                await colorTable(table, rows, context);
                            }
                            else {
                                const rowspanTracker: number[] = new Array(maxCols).fill(0);
                                let lastParamRowIndex = -1; // row index of last non-empty Parameter cell

                                rows.forEach((row, rowIndex) => {
                                    const cells = Array.from(row.querySelectorAll('td, th'));
                                    let cellIndex = 0;

                                    // track first column text in HTML row
                                    let firstColText = "";

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

                                        // save first column text if this is first column
                                        if (cellIndex === 0) {
                                            firstColText = cellText.trim();
                                        }

                                        table.getCell(rowIndex, cellIndex).value = cellText;

                                        // colspan blank filling
                                        for (let i = 1; i < colspan; i++) {
                                            if (cellIndex + i < maxCols) {
                                                table.getCell(rowIndex, cellIndex + i).value = "";
                                            }
                                        }

                                        // rowspan tracking
                                        if (rowspan > 1) {
                                            for (let i = 0; i < colspan; i++) {
                                                if (cellIndex + i < maxCols) {
                                                    rowspanTracker[cellIndex + i] = rowspan - 1;
                                                }
                                            }
                                        }

                                        cellIndex += colspan;
                                    });

                                    // ✅ AFTER filling the row → apply merge logic for 1st column
                                    if (rowIndex === 0) {
                                        // header row typically
                                        return;
                                    }

                                    if (firstColText) {
                                        // new parameter starts
                                        lastParamRowIndex = rowIndex;
                                    } else {
                                        // empty parameter row → merge vertically with previous parameter row
                                        if (lastParamRowIndex !== -1) {
                                            const topCell = table.getCell(lastParamRowIndex, 0);
                                            const bottomCell = table.getCell(rowIndex, 0);

                                            // merge bottom into top
                                            topCell.merge(bottomCell);

                                            // ✅ center align merged cell
                                            try {
                                                topCell.verticalAlignment = Word.VerticalAlignment.center;
                                                topCell.body.paragraphs.getFirst().alignment = Word.Alignment.center;
                                            } catch (e) {
                                                // ignore (safe fallback)
                                            }
                                        }
                                    }
                                });

                            }

                            newSelection = table.getCell(0, 0); // Set the cursor to the start of the table
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
                            newSelection = selection; // If it's not a table, just use the existing selection.
                        }
                    }
                }
            }
            else if (type === "IMAGE") {
                let base64Image: string = word.EditorValue;

                if (base64Image.startsWith("data:image/svg+xml")) {
                    // Convert SVG → PNG
                    base64Image = await svgBase64ToPngBase64(base64Image);
                } else if (base64Image.startsWith("data:image")) {
                    base64Image = base64Image.split(",")[1]; // strip prefix
                }

                selection.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.replace);
                newSelection = selection;
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
                newSelection = selection; // After inserting the text, set selection to it.
            }

            // Move the cursor to the next line after content insertion
            const nextLineParagraph = selection.insertParagraph("", Word.InsertLocation.after);
            await context.sync();

            // Set the new cursor position after content
            newSelection = nextLineParagraph;
            selection.select(); // Select the new paragraph where the cursor will be
            await context.sync();

        } catch (error) {
            console.error('Detailed error:', error);
        }
    });
}


export async function openAITag(tag) {
    tag.ReportHeadAIHistoryList.forEach((historyList) => {
        historyList.Response = removeQuotes(historyList.Response);
        tag.FilteredReportHeadAIHistoryList.unshift(historyList);
    });


}

export async function generateCheckboxHistory(tag, type: "Summary" | "AITag") {
    var skipFetch = false;
    if ((!tag.FilteredReportHeadAIHistoryList || tag.FilteredReportHeadAIHistoryList.length === 0) && !skipFetch) {
        if (type !== 'Summary') {
            await AIService.fetchAIHistory(tag);
        } else {
            await summaryService.fetchSummaryAIHistory(tag);
        }
    }

    const history = tag.FilteredReportHeadAIHistoryList;

    if (history.length === 0) {
        return '<div>No AI history available.</div>';
    }

    // Check current theme
    const store = StoreService.getInstance();
    const isDark = store.theme === 'Dark';
    const closeBtnClass = isDark
        ? 'fa-solid fa-circle-xmark bg-dark text-light'
        : 'fa-solid fa-circle-xmark bg-light text-dark';

    const headerBgClass = isDark ? 'bg-dark text-light' : 'bg-white text-dark';
    const DisplayName = type === 'Summary' ? tag.Name : tag.DisplayName;
    const closeBar = `
    <div class="chat-header sticky-top ${headerBgClass} z-3">
        <div class="d-flex justify-content-between align-items-center px-2 pt-3">
            <div class="d-flex align-items-center ms-3">
                <i class="fa fa-microchip-ai text-muted me-2"></i>
                <span class="fw-bold">${DisplayName}</span>
            </div>
            <div class="d-flex justify-content-center align-items-center me-3 c-pointer" id="close-btn-tag">
                <i class="${closeBtnClass}" id="close-ai-window"></i>
            </div>
        </div>
        <hr class="mt-2 mb-1 mx-3">
        </div>
    `;

    const chatBody = `
        <div class="chat-body flex-grow-1 overflow-auto">
            ${generateChatHistoryHtml(history)}
        </div>
    `;

    const chatFooterHtml = `
        <div class="d-flex align-items-end justify-content-end chatbox p-2" id="chatFooter">
            ${chatfooter(tag)}
        </div>
    `;

    initializeAIHistoryEvents(tag, store.jwt, store.availableKeys, type);

    return `${closeBar}${chatBody}${chatFooterHtml}`;
}




export async function setupPromptBuilderUI(container, promptBuilderList) {

    // Static template and field definitions
    let preview = '';
    let templateText = '';

    // Field configs (can be extended)
    let fieldsList = [];

    // Create the form container
    // Create the form container
    container.innerHTML = `
  <div class="form-group mb-3 p-3 pt-0">
    <label class='form-label'><span class="text-danger">*</span> Prompt Builder Template</label>
    <select id="promptBuilderTemplate" class="form-control">
      <option value="" disabled selected>Select a template</option>
    </select>
    <div id="templateError" class="invalid-feedback d-none">Type is required.</div>
  </div>

  <div id="fieldsContainer"></div>

  <div class="form-group mb-3 p-3 pt-0" id="previewContainer" style="display: none;">
    <label class="mb-2">Preview</label>
    <div id="preview" class="form-control"></div>
  </div>

  <div class="d-flex justify-content-between px-3 align-items-center mt-3">
    <span id="resetBtn" class="text-primary fw-bold" style="cursor: pointer;">Reset</span>
    <button id="applyBtn" class="btn btn-primary text-white" disabled>Apply Prompt</button>
  </div>
`;

    // Element references
    const templateSelect = container.querySelector('#promptBuilderTemplate') as HTMLSelectElement;
    const applyBtn = container.querySelector('#applyBtn') as HTMLButtonElement;
    const resetBtn = container.querySelector('#resetBtn') as HTMLSpanElement;
    const previewDiv = container.querySelector('#preview') as HTMLDivElement;
    const fieldsContainer = container.querySelector('#fieldsContainer') as HTMLDivElement;
    const previewContainer = container.querySelector('#previewContainer') as HTMLDivElement;
    const templateError = container.querySelector('#templateError') as HTMLDivElement;

    // Populate template dropdown
    promptBuilderList.forEach((item) => {
        const option = document.createElement('option');
        option.value = item.ID.toString();
        option.textContent = item.Name;
        templateSelect.appendChild(option);
    });

    templateSelect.addEventListener('change', async () => {
        const templateId = templateSelect.value;
        const jwt = sessionStorage.getItem('token') || '';

        const data = await getPromptTemplateById(templateId, jwt);
        if (data.Status && data.Data) {
            fieldsList = data.Data;
            preview = promptBuilderList.find((item) => item.ID.toString() === templateId).Template;

            templateText = promptBuilderList.find((item) => item.ID.toString() === templateId).Template;
        }
        if (!templateId) {
            templateError.classList.remove('d-none');
            return;
        }

        templateError.classList.add('d-none');

        renderFields();
        updatePreview();
    });

    function renderFields() {
        fieldsContainer.innerHTML = '';

        fieldsList.forEach((field) => {
            const div = document.createElement('div');
            div.className = 'form-group mb-3 p-3 pt-0';

            const label = document.createElement('label');
            label.textContent = field.Label;
            div.appendChild(label);

            if (field.Type === 1) {
                const input = document.createElement('input');
                input.type = 'text';
                input.className = 'form-control';
                input.id = field.Label;
                input.addEventListener('input', replaceKeywordsManually);
                div.appendChild(input);
            } else if (field.Type === 2) {
                const select = document.createElement('select');
                select.className = 'form-control';
                select.id = field.Label;
                field.PromptTemplateOptionList.forEach((opt: any) => {
                    const option = document.createElement('option');
                    option.value = opt.Text;
                    option.textContent = opt.Option;
                    select.appendChild(option);
                });
                select.addEventListener('change', replaceKeywordsManually);
                div.appendChild(select);
            }

            fieldsContainer.appendChild(div);
        });
    }

    function replaceKeywordsManually() {
        const keywordMap: { [key: string]: string } = {};

        fieldsList.forEach((field) => {
            const id = field.Label;
            const keyword = `#${id}#`;

            let value = '';
            const element = document.getElementById(id) as HTMLInputElement | HTMLSelectElement;

            if (element) {
                value = (element instanceof HTMLInputElement || element instanceof HTMLSelectElement)
                    ? element.value
                    : '';
            }

            keywordMap[keyword] = value ? value : keyword;
        });

        let insertValue = templateText;
        for (const [keyword, value] of Object.entries(keywordMap)) {
            insertValue = insertValue.replace(new RegExp(keyword, 'g'), value);
        }

        preview = insertValue;
        previewDiv.textContent = preview;
        previewContainer.style.display = preview ? 'block' : 'none';
        applyBtn.disabled = preview === '';
    }

    function updatePreview() {
        replaceKeywordsManually();
    }

    function resetForm() {
        // Reset only the dynamic field values
        fieldsList.forEach((field) => {
            const element = document.getElementById(field.Label) as HTMLInputElement | HTMLSelectElement;
            if (element) {
                if (element instanceof HTMLInputElement) {
                    element.value = '';
                } else if (element instanceof HTMLSelectElement) {
                    element.selectedIndex = 0; // optional: reset to first option
                }
            }
        });

        // Clear preview
        previewDiv.textContent = templateText;
        preview = templateText;
    }


    function applyPrompt() {
        if (!preview) return;

        const promptTextarea = document.getElementById('prompt') as HTMLTextAreaElement;
        if (promptTextarea) {
            promptTextarea.value = preview;
            switchToAddTag()
        }

    }

    resetBtn.addEventListener('click', resetForm);
    applyBtn.addEventListener('click', applyPrompt);
}


export async function insertTagPrompt(tag) {
    return Word.run(async (context) => {
        try {
            const selection = context.document.getSelection();
            await context.sync();

            if (!selection) throw new Error("Invalid selection");

            /* --------------------------------------------------
               1️⃣ Create invisible anchor at cursor
            -------------------------------------------------- */
            const anchorChar = selection.insertText(
                "\u200B", // zero-width space
                Word.InsertLocation.replace
            );
            await context.sync();

            let cursor = anchorChar.getRange();

            let bookmarkStart: Word.Range | null = null;
            let bookmarkEnd: Word.Range | null = null;

            const include = (r: Word.Range) => {
                if (!bookmarkStart) {
                    bookmarkStart = r.getRange("Start");
                }
                bookmarkEnd = r.getRange("End");
            };

            /* --------------------------------------------------
               2️⃣ Insert content
            -------------------------------------------------- */
            if (tag.ComponentKeyDataType === "TABLE") {
                const parser = new DOMParser();
                const doc = parser.parseFromString(tag.EditorValue, "text/html");
                const bodyNodes = Array.from(doc.body.childNodes);

                for (const node of bodyNodes) {

                    // TEXT NODE
                    if (node.nodeType === Node.TEXT_NODE) {
                        let txt = node.textContent?.trim();
                        if (!txt) continue;

                        txt = txt.replace(/\n- /g, "\n• ");
                        for (const line of txt.split("\n")) {
                            if (!line.trim()) continue;

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

                            const maxCols = Math.max(
                                ...rows.map(r =>
                                    Array.from(r.querySelectorAll("td, th"))
                                        .reduce(
                                            (s, c) => s + parseInt(c.getAttribute("colspan") || "1"),
                                            0
                                        )
                                )
                            );

                            const p = cursor.insertParagraph("", Word.InsertLocation.after);
                            const table = p.insertTable(
                                rows.length,
                                maxCols,
                                Word.InsertLocation.after
                            );

                            const store = StoreService.getInstance();
                            const resolvedTableStyle = resolveWordTableStyle(store.tableStyle);
                            if (resolvedTableStyle !== 'none') {
                                table.style = resolvedTableStyle;
                            }

                            if (store.colorPallete.Customize) {
                                await colorTable(table, rows, context);
                            } else {
                                const rowspanTracker: number[] = new Array(maxCols).fill(0);
                                let lastParamRowIndex = -1; // row index of last non-empty Parameter cell

                                rows.forEach((row, rowIndex) => {
                                    const cells = Array.from(row.querySelectorAll('td, th'));
                                    let cellIndex = 0;

                                    // track first column text in HTML row
                                    let firstColText = "";

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

                                        // save first column text if this is first column
                                        if (cellIndex === 0) {
                                            firstColText = cellText.trim();
                                        }

                                        table.getCell(rowIndex, cellIndex).value = cellText;

                                        // colspan blank filling
                                        for (let i = 1; i < colspan; i++) {
                                            if (cellIndex + i < maxCols) {
                                                table.getCell(rowIndex, cellIndex + i).value = "";
                                            }
                                        }

                                        // rowspan tracking
                                        if (rowspan > 1) {
                                            for (let i = 0; i < colspan; i++) {
                                                if (cellIndex + i < maxCols) {
                                                    rowspanTracker[cellIndex + i] = rowspan - 1;
                                                }
                                            }
                                        }

                                        cellIndex += colspan;
                                    });

                                    // ✅ AFTER filling the row → apply merge logic for 1st column
                                    if (rowIndex === 0) {
                                        // header row typically
                                        return;
                                    }

                                    if (firstColText) {
                                        // new parameter starts
                                        lastParamRowIndex = rowIndex;
                                    } else {
                                        // empty parameter row → merge vertically with previous parameter row
                                        // empty parameter row → merge vertically with previous parameter row
                                        if (lastParamRowIndex !== -1) {
                                            const topCell = table.getCell(lastParamRowIndex, 0);
                                            const bottomCell = table.getCell(rowIndex, 0);

                                            // merge bottom into top
                                            topCell.merge(bottomCell);

                                            // ✅ center align merged cell
                                            try {
                                                topCell.verticalAlignment = Word.VerticalAlignment.center;
                                                topCell.body.paragraphs.getFirst().alignment = Word.Alignment.center;
                                            } catch (e) {
                                                // ignore (safe fallback)
                                            }
                                        }

                                    }
                                });

                            }

                            include(table.getRange());
                            cursor = table.getRange();
                        }

                        // OTHER HTML ELEMENTS
                        else {
                            let txt = el.innerText?.trim();
                            if (!txt) continue;

                            txt = txt.replace(/\n- /g, "\n• ");
                            for (const line of txt.split("\n")) {
                                if (!line.trim()) continue;

                                const p = cursor.insertParagraph("", Word.InsertLocation.after);
                                insertLineWithHeadingStyle(p, line);
                                include(p.getRange());
                                cursor = p.getRange();
                            }
                        }
                    }
                }
            }

            // NON-TABLE CONTENT
            else {
                const txt = tag.EditorValue.replace(/\n- /g, "\n• ").trim();

                for (const line of txt.split("\n")) {
                    if (!line.trim()) continue;

                    const p = cursor.insertParagraph("", Word.InsertLocation.after);
                    insertLineWithHeadingStyle(p, line);
                    include(p.getRange());
                    cursor = p.getRange();
                }
            }

            await context.sync();

            /* --------------------------------------------------
               3️⃣ Create ONE bookmark covering everything
            -------------------------------------------------- */
            if (bookmarkStart && bookmarkEnd) {
                const bookmarkName =
                    `ID${tag.ID}_Split_${getDateTimeStamp()}`;

                bookmarkStart
                    .expandTo(bookmarkEnd)
                    .insertBookmark(bookmarkName);
            }

            /* --------------------------------------------------
               4️⃣ Remove invisible anchor
            -------------------------------------------------- */
            anchorChar.delete();

            await context.sync();
            toaster("Inserted successfully", "success");

        } catch (err) {
            console.error(err);
            toaster("Something went wrong", "error");
        }
    });
}

export function getDateTimeStamp() {
    const d = new Date();

    const pad = (n) => n.toString().padStart(2, "0");

    return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_` +
        `${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}




export function initializeAIHistoryEvents(tag: any, jwt: string, availableKeys: any, type: "Summary" | "AITag") {
    setTimeout(() => {
        tag.FilteredReportHeadAIHistoryList.forEach((chat: any, index: number) => {
            // Copy buttons
            if (tag.textareavalue) {
                (document.getElementById(`chatInput`) as HTMLTextAreaElement).value = tag.textareavalue;
                delete (tag.textareavalue)
            }

            // After initializing buttons inside setTimeout
            const chatInput = document.getElementById("chatInput") as HTMLTextAreaElement;
            const changeSourceButton = document.getElementById("changeSourceButton") as HTMLButtonElement;

            if (chatInput && changeSourceButton) {
                // Run once on load
                changeSourceButton.disabled = chatInput.value.trim().length === 0;

                // Listen for changes in textarea
                chatInput.addEventListener("input", () => {
                    if (chatInput.value.trim().length > 0) {
                        changeSourceButton.disabled = false;
                    } else {
                        changeSourceButton.disabled = true;
                    }
                });
            };
            document.getElementById(`copyPrompt-${index}`)?.addEventListener('click', () => copyText(chat.Prompt));
            const savePromptele = document.getElementById(`savePrompt-${index}`);
            if (savePromptele) {
                document.getElementById(`savePrompt-${index}`)?.addEventListener('click', () => {
                    const container = document.getElementById('confirmation-popup');
                    if (container) {
                        container.innerHTML = Confirmationpopup('Do you want to save the current prompt as a global default?');

                        // Wait for DOM to update and then attach cancel button listener
                        setTimeout(() => {
                            document.getElementById('confirmation-popup-cancel')?.addEventListener('click', () => {
                                container.innerHTML = '';
                            });

                            document.getElementById('confirmation-popup-confirm')?.addEventListener('click', async () => {
                                try {
                                    document.getElementById('confirmation-popup-cancel')?.setAttribute('disabled', 'true');
                                    document.getElementById('confirmation-popup-confirm')?.setAttribute('disabled', 'true');
                                    let updatedTag = JSON.parse(JSON.stringify(tag));
                                    updatedTag.Prompt = chat.Prompt;
                                    const data = await updatePromptTemplate(updatedTag, jwt);
                                    if (data['Status']) {
                                        toaster('Updated Succesfully', 'success');
                                        container.innerHTML = '';
                                    } else {
                                        document.getElementById('confirmation-popup-cancel')?.setAttribute('disabled', 'false');
                                        document.getElementById('confirmation-popup-confirm')?.setAttribute('disabled', 'false');
                                        toaster('Something went wrong', 'error');


                                    }
                                } catch (error) {
                                    document.getElementById('confirmation-popup-cancel')?.setAttribute('disabled', 'false');
                                    document.getElementById('confirmation-popup-confirm')?.setAttribute('disabled', 'false');
                                    toaster('Something went wrong', 'error');
                                }
                            });
                        }, 0);
                    }
                });
            }

            const openRefferance = document.getElementById(`openRefferance-${index}`);
            if (openRefferance) {
                document.getElementById(`openRefferance-${index}`)?.addEventListener('click', () => {
                    const container = document.getElementById('confirmation-popup');
                    if (container) {
                        const store = StoreService.getInstance();
                        const sourceList = type === 'Summary' ? store.sourceSummaryList : store.sourceList;
                        const rawSources = type === 'Summary' ? chat.SourceVector : chat.SourceValue;
                        const sourceIds = Array.isArray(rawSources) ? rawSources : (rawSources ? String(rawSources).split(',') : []);

                        const chatSources = sourceIds.map((item: any) => {
                            if (type === 'Summary') {
                                return sourceList.find(
                                    (source: any) => String(item) === String(source.VectorID)
                                );
                            } else {
                                return sourceList.find(
                                    (source: any) => Number(item) === source.VectorID
                                );
                            }
                        });

                        const sources = chatSources.filter((src: any) => !!src);
                        const popupData = {
                            Data: chat.Evidences,
                            Name: type === 'Summary' ? tag.Name : tag.DisplayName,
                            UserValue: chat.Response,
                            Sources: sources
                        }

                        container.innerHTML = DataModalPopup(popupData);

                        // Wait for DOM to update and then attach cancel button listener
                        setTimeout(() => {
                            document.getElementById('confirmation-popup-cancel')?.addEventListener('click', () => {
                                container.innerHTML = '';
                            });

                            document.getElementById('confirmation-popup-confirm')?.addEventListener('click', async () => {
                                try {
                                    document.getElementById('confirmation-popup-cancel')?.setAttribute('disabled', 'true');
                                    document.getElementById('confirmation-popup-confirm')?.setAttribute('disabled', 'true');
                                    let updatedTag = JSON.parse(JSON.stringify(tag));
                                    updatedTag.Prompt = chat.Prompt;
                                    const data = await updatePromptTemplate(updatedTag, jwt);
                                    if (data['Status']) {
                                        toaster('Updated Succesfully', 'success');
                                        container.innerHTML = '';
                                    } else {
                                        document.getElementById('confirmation-popup-cancel')?.setAttribute('disabled', 'false');
                                        document.getElementById('confirmation-popup-confirm')?.setAttribute('disabled', 'false');
                                        toaster('Something went wrong', 'error');


                                    }
                                } catch (error) {
                                    document.getElementById('confirmation-popup-cancel')?.setAttribute('disabled', 'false');
                                    document.getElementById('confirmation-popup-confirm')?.setAttribute('disabled', 'false');
                                    toaster('Something went wrong', 'error');
                                }
                            });

                            document.getElementById('datamodel-popup-ok')?.addEventListener('click', async () => {
                                container.innerHTML = ''
                            })
                        }, 0);
                    }
                });
            }


            document.getElementById(`copyResponse-${index}`)?.addEventListener('click', () => copyText(chat.Response));

            // Checkbox logic
            const checkbox = document.getElementById(`checkbox-${index}`) as HTMLInputElement;
            if (checkbox) {
                checkbox.addEventListener('change', async (event: Event) => {
                    const isChecked = (event.target as HTMLInputElement).checked;

                    // Reset all
                    tag.FilteredReportHeadAIHistoryList.forEach((_: any, otherIndex: number) => {
                        const otherCheckbox = document.getElementById(`checkbox-${otherIndex}`) as HTMLInputElement;
                        const responseContainer = document.getElementById(`responseContainer-${otherIndex}`);
                        if (otherCheckbox) otherCheckbox.checked = false;
                        if (responseContainer) {
                            responseContainer.classList.remove('ai-selected-response');
                            responseContainer.classList.add('bg-light');
                        }
                        tag.FilteredReportHeadAIHistoryList[otherIndex].Selected = 0;
                    });

                    // Set selected
                    if (isChecked) {
                        checkbox.checked = true;
                        const responseContainer = document.getElementById(`responseContainer-${index}`);
                        if (responseContainer) {
                            responseContainer.classList.add('ai-selected-response');
                            responseContainer.classList.remove('bg-light');
                        }
                        chat.Selected = 1;
                    } else {
                        chat.Selected = 0;
                    }

                    try {
                        const data = type === 'Summary' ? await updateSummaryHistory(chat, jwt) : await updateAiHistory(chat, jwt);
                        if (data['Data']) {
                            tag.ReportHeadAIHistoryList = JSON.parse(JSON.stringify(data['Data']));
                            tag.FilteredReportHeadAIHistoryList = [];

                            tag.ReportHeadAIHistoryList.forEach((historyList: any) => {
                                historyList.Response = removeQuotes(historyList.Response);
                                tag.FilteredReportHeadAIHistoryList.unshift(historyList);
                            });

                            const finalResponse = chat.FormattedResponse
                                ? '\n' + updateEditorFinalTable(chat.FormattedResponse)
                                : chat.Response;

                            tag.ComponentKeyDataType = chat.FormattedResponse ? 'TABLE' : 'TEXT';
                            tag.UserValue = finalResponse;
                            tag.EditorValue = finalResponse;
                            tag.text = finalResponse;

                            const currentlySelected = tag.FilteredReportHeadAIHistoryList.some((item: any) => item.Selected === 1);
                            tag.IsApplied = !currentlySelected;
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

                            const store = StoreService.getInstance();
                            store.aiTagList.forEach(currentTag => {
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
                        }
                    } catch (err) {
                        console.error('Failed to update AI history:', err);
                    }
                });
            }
        });

        // Close button
        document.getElementById(`close-btn-tag`)?.addEventListener('click', () => {
            const store = StoreService.getInstance();
            if (store.mode === "Home") {
                loadHomepage(availableKeys)
            } else if (store.mode === "Summary") {
                loadSummarypage(availableKeys);
            }
        });

        // Button: Insert Tag
        document.getElementById(`insertTagButton`)?.addEventListener('click', () => {
            if (!tag.IsApplied) {
                insertTagPrompt(tag);
            }
        });

        // Button: Send Prompt
        document.getElementById(`sendPromptButton`)?.addEventListener('click', () => {
            const textareaValue = (document.getElementById(`chatInput`) as HTMLTextAreaElement).value;
            AIService.sendPrompt(tag, textareaValue, type);
        });

        // Button: Change Source
        document.getElementById(`changeSourceButton`)?.addEventListener('click', () => {
            const textareaValue = (document.getElementById(`chatInput`) as HTMLTextAreaElement).value;
            tag.textareavalue = textareaValue;
            createMultiSelectDropdown(tag, type);
        });

        // Mention dropdown
        mentionDropdownFn(`chatInput`, `mention-dropdown`, 'edit');
    }, 0);
}

