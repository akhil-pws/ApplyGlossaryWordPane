import { getPromptTemplateById, updateGroupKey, updateAiHistory, updatePromptTemplate } from "../api";
import { chatfooter, copyText, generateChatHistoryHtml, insertLineWithHeadingStyle, removeQuotes, renderSelectedTags, switchToAddTag, updateEditorFinalTable } from "../functions";
import { addGenAITags, aiTagList, applyAITagFn, availableKeys, createMultiSelectDropdown, fetchAIHistory, isPendingResponse, jwt, mentionDropdownFn, selectedNames, sendPrompt, sourceList, theme } from "../taskpane";
import { Confirmationpopup, DataModalPopup, toaster } from "./bodyelements";

let preview = '';


export function loadHomepage(availableKeys) {
    const searchBoxClass = theme === 'Dark' ? 'bg-secondary text-light' : 'bg-white text-dark';

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

    const searchBox = document.getElementById('search-box');
    const suggestionList = document.getElementById('suggestion-list');

    function updateSuggestions() {
        const searchTerm = searchBox.value.trim().toLowerCase();
        suggestionList.replaceChildren(); // Clear previous results
        if (searchTerm === '') {
            suggestionList.innerHTML = ''
            return;
        }

        const filteredMentions = availableKeys.filter(mention =>
            mention.DisplayName.toLowerCase().includes(searchTerm)
        );

        const nonAITags = filteredMentions.filter(m => m.AIFlag === 0);
        const aiTags = filteredMentions.filter(m => m.AIFlag === 1);

        const createSection = (labelText, mentions, isAISection = false) => {
            if (mentions.length === 0) return;

            // Define the theme classes based on the current theme
            const themeClasses = theme === 'Dark'
                ? { itemClass: 'bg-dark text-light list-hover-dark', labelClass: 'bg-dark text-light' }
                : { itemClass: 'bg-light text-dark list-hover-light', labelClass: 'bg-light text-dark' };


            // Create the section label
            const label = document.createElement('li');
            label.className = `list-group-item fw-bold text-secondary ${themeClasses.labelClass}`;
            label.textContent = labelText;
            suggestionList.appendChild(label);

            // Loop through mentions and create the list items
            mentions.forEach(mention => {
                const listItem = document.createElement('li');
                listItem.className = `list-group-item list-group-item-action ${themeClasses.itemClass}`; // Apply the theme classes

                // Create the icon for AI or non-AI tags
                const icon = isAISection
                    ? `<i class="fa-solid fa-microchip-ai text-muted me-2"></i>`
                    : `<i class="fa-solid fa-layer-group text-muted me-2"></i>`;

                listItem.innerHTML = `${icon} ${mention.DisplayName}`;

                listItem.onclick = () => {
                    if (isAISection) {
                        const appBody = document.getElementById('app-body');
                        appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';
                        generateCheckboxHistory(mention).catch(error => {
                            appBody.innerHTML = '<div class="text-danger p-2">Error loading data</div>';
                        }).then(html => {
                            appBody.innerHTML = html;
                        });
                    } else {
                        replaceMention(mention, mention.ComponentKeyDataType);
                        suggestionList.replaceChildren();
                    }
                };

                suggestionList.appendChild(listItem);
            });
        };

        // Call the function for each section
        createSection('Properties', nonAITags, false);
        createSection('AI Tags', aiTags, true);

    }
    if (selectedNames.length > 0) {
        const badgeWrapper = document.getElementById('tags-in-selected-text');
        badgeWrapper.classList.remove('d-none');
        badgeWrapper.classList.add('d-block');
        renderSelectedTags(selectedNames, availableKeys);
    }

    // Add input event listener to the search box
    let debounceTimeout;
    searchBox.addEventListener('input', () => {
        clearTimeout(debounceTimeout);
        debounceTimeout = setTimeout(updateSuggestions, 300); // Delay input handling by 300ms
    });

    document.getElementById('add-btn-tag').addEventListener('click', () => {
        if (!isPendingResponse) {
            addGenAITags();
        }
    });

    document.getElementById('apply-btn-tag').addEventListener('click', () => {
        if (!isPendingResponse) {
            applyAITagFn();
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

                            newSelection = table.getCell(0, 0); // Set the cursor to the start of the table
                        } else {
                            const elementText = element.innerText.trim();
                            if (elementText) {
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

export async function generateCheckboxHistory(tag) {
    if (!tag.FilteredReportHeadAIHistoryList || tag.FilteredReportHeadAIHistoryList.length === 0) {
        await fetchAIHistory(tag);
    }

    const history = tag.FilteredReportHeadAIHistoryList;

    if (history.length === 0) {
        return '<div>No AI history available.</div>';
    }

    // Check current theme
    const isDark = theme === 'Dark';
    const closeBtnClass = isDark
        ? 'fa-solid fa-circle-xmark bg-dark text-light'
        : 'fa-solid fa-circle-xmark bg-light text-dark';

    const closeBar = `
        <div class="d-flex justify-content-between align-items-center px-2 pt-3">
            <div class="d-flex align-items-center ms-3">
                <i class="fa fa-microchip-ai text-muted me-2"></i>
                <span class="fw-bold">${tag.DisplayName}</span>
            </div>
            <div class="d-flex justify-content-center align-items-center me-3 c-pointer" id="close-btn-tag">
                <i class="${closeBtnClass}" id="close-ai-window"></i>
            </div>
        </div>
        <hr class="mt-2 mb-1 mx-3">
    `;

    const chatBody = `
        <div class="chat-body">
            ${generateChatHistoryHtml(history)}
        </div>
    `;

    const chatFooterHtml = `
        <div class="d-flex align-items-end justify-content-end chatbox p-2" id="chatFooter">
            ${chatfooter(tag)}
        </div>
    `;

    initializeAIHistoryEvents(tag, jwt, availableKeys);

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


async function insertTagPrompt(tag: any) {
    return Word.run(async (context) => {
        try {
            const selection = context.document.getSelection();
            await context.sync();

            if (!selection) {
                throw new Error('Selection is invalid or not found.');
            }


            const cleanDisplayName = tag.ID;
            const uniqueStr = new Date().getTime();
            const bookmarkName = `ID${cleanDisplayName}_Split_${uniqueStr}`;

            const startMarker = selection.insertParagraph("[[BOOKMARK_START]]", Word.InsertLocation.before);
            await context.sync();

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
                                            insertLineWithHeadingStyle(selection, line);
                                        }
                                    });
                                }
                            }
                        }
                    }
                } else {
                    let content = removeQuotes(tag.EditorValue);
                    let lines = content.split(/\r?\n/);
                    lines.forEach(line => {
                        insertLineWithHeadingStyle(selection, line)
                        // selection.insertParagraph(line, Word.InsertLocation.before);
                    });
                }
            }

            const endMarker = selection.insertParagraph("[[BOOKMARK_END]]", Word.InsertLocation.after);
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

                afterBookmark.select();
                start.delete()
                end.delete();
                afterBookmark.delete();
                await context.sync();

            }
            toaster('Inserted successfully', 'success');

        } catch (error) {
            toaster('Something went wrong', 'error');

            console.error('Detailed error:', error);
        }
    });
}



export function initializeAIHistoryEvents(tag: any, jwt: string, availableKeys: any) {
    setTimeout(() => {
        tag.FilteredReportHeadAIHistoryList.forEach((chat: any, index: number) => {
            // Copy buttons
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
                        const chatSources = chat.SourceValue.map((item: any) => {
                            return sourceList.find(
                                (source: any) => Number(item) === source.VectorID
                            );
                        });

                        const sources = chatSources.filter((src: any) => !!src);
                        const popupData = {
                            Data: chat.Evidences,
                            Name: tag.DisplayName,
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

                            document.getElementById('datamodel-popup-ok')?.addEventListener('click',async() =>{
                                container.innerHTML=''
                            })
                        }, 0);
                    }
                });
            }


            document.getElementById(`copyResponse-${index}`)?.addEventListener('click', () => copyText(chat.Response));

            // Close button
            document.getElementById(`close-btn-tag`)?.addEventListener('click', () => loadHomepage(availableKeys));

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
                        const data = await updateAiHistory(chat, jwt);
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
                        }
                    } catch (err) {
                        console.error('Failed to update AI history:', err);
                    }
                });
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
            sendPrompt(tag, textareaValue);
        });

        // Button: Change Source
        document.getElementById(`changeSourceButton`)?.addEventListener('click', () => {
            createMultiSelectDropdown(tag);
        });

        // Mention dropdown
        mentionDropdownFn(`chatInput`, `mention-dropdown`, 'edit');
    }, 0);
}
