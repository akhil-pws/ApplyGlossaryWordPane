import { getPromptTemplateById, updateGroupKey, updateAiHistory, updatePromptTemplate } from "./draft.api";
import { chatfooter, copyText, generateChatHistoryHtml, insertLineWithHeadingStyle, removeQuotes, switchToAddTag, updateEditorFinalTable, colorTable, svgBase64ToPngBase64, resolveWordTableStyle, renderSelectedTags, parseHtmlTableToGrid, transposeGrid, detectTableCase, confirmSwitchChatHistory, applyCustomTextStyleToCell } from "./draft-functions";
import { addGenAITags, applyTagFn, createMultiSelectDropdown, mentionDropdownFn } from "../taskpane";
import { StoreService } from "../services/store.service";
import { AIService } from "../services/ai.service";
import { Confirmationpopup, DataModalPopup, toaster, PromptBuilderModalPopup, NewChatSourceModalPopup } from "../components/bodyelements";
import { loadSummarypage } from "../summary/summary";
import { summaryService } from "../services/summary.service";
import { updateSummaryHistory, updateSummaryTagPrompt } from "../summary/summary.api";
import { DocStorage } from "../utils/doc-storage";
import { faMicrochipAi, getIconSvg } from "../utils/fontawesome-icons";

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
                if (isAISection) icon = getIconSvg(faMicrochipAi, 'text-muted me-2');
                if (isImageSection) icon = `<i class="fa-solid fa-image text-muted me-2"></i>`;

                listItem.innerHTML = `${icon} ${mention.DisplayName}`;

                listItem.onclick = () => {
                    if (isAISection) {
                        const tagId = mention.ID || mention.ReportHeadSummaryTagID;
                        if (store.currentChatTagId !== -1 && store.currentChatTagId !== undefined && store.currentChatTagId !== null && String(store.currentChatTagId) === String(tagId)) {
                            return;
                        }
                        confirmSwitchChatHistory(() => {
                            const appBody = document.getElementById('app-body');
                            appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';
                            generateCheckboxHistory(mention, "AITag")
                                .catch(() => appBody.innerHTML = '<div class="text-danger p-2">Error loading data</div>')
                                .then(html => { appBody.innerHTML = html; });
                        });
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

                            let grid = parseHtmlTableToGrid(rows);
                            const tableCase = detectTableCase(grid);

                            const store = StoreService.getInstance();
                            const base = store.tableStyle.split(" - ")[0].trim();

                            if (base === 'Table Grid 2') {
                                store.isReversed = true;
                            } else {
                                store.isReversed = false;
                            }

                            if (store.isReversed && tableCase !== "CASE_1") {
                                grid = transposeGrid(grid);
                            }

                            const numRows = grid.length;
                            const numCols = grid[0]?.length || 0;

                            const paragraph = selection.insertParagraph("", Word.InsertLocation.before);
                            await context.sync();

                            const table = paragraph.insertTable(numRows, numCols, Word.InsertLocation.after);
                            const resolvedTableStyle = resolveWordTableStyle(store.tableStyle);
                            if (resolvedTableStyle !== 'none') {
                                table.style = resolvedTableStyle;
                            }

                            await context.sync();

                            // Population/Merging logic
                            if (!store.isReversed) {
                                if (!store.colorPallete.Customize) {
                                    grid.forEach((row, rowIndex) => {
                                        row.forEach((cellValue, cellIndex) => {
                                            const tableCell = table.getCell(rowIndex, cellIndex);
                                            tableCell.value = cellValue;
                                            applyCustomTextStyleToCell(tableCell, store);
                                        });
                                    });

                                    // Vertical merging logic for 1st column (standard view)
                                    let lastParamRowIndex = -1;
                                    grid.forEach((row, rowIndex) => {
                                        if (rowIndex === 0) return; // skip header
                                        const firstColText = row[0];
                                        if (firstColText) {
                                            lastParamRowIndex = rowIndex;
                                        } else if (lastParamRowIndex !== -1) {
                                            const topCell = table.getCell(lastParamRowIndex, 0);
                                            const bottomCell = table.getCell(rowIndex, 0);
                                            topCell.merge(bottomCell);
                                            try {
                                                topCell.verticalAlignment = Word.VerticalAlignment.center;
                                                topCell.body.paragraphs.getFirst().alignment = Word.Alignment.center;
                                            } catch (e) { }
                                        }
                                    });
                                }
                            } else {
                                // Manual population for transposed table
                                grid.forEach((row, rowIndex) => {
                                    row.forEach((cellValue, cellIndex) => {
                                        const tableCell = table.getCell(rowIndex, cellIndex);
                                        tableCell.value = cellValue;
                                        applyCustomTextStyleToCell(tableCell, store);
                                    });
                                });
                            }

                            // Styling logic (always call if Customize, now passing isReversed)
                            if (store.colorPallete.Customize) {
                                await colorTable(table, rows, context, store.isReversed);
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

export async function generateCheckboxHistory(tag, type: "Summary" | "AITag", targetChatId?: string | number) {
    const store = StoreService.getInstance();
    store.currentChatTagId = tag.ID || tag.ReportHeadSummaryTagID;
    DocStorage.setItem("currentChatTagId", String(store.currentChatTagId));

    if (!tag.ChatSessions || tag.ChatSessions.length === 0) {
        if (type !== 'Summary') {
            await AIService.fetchAIHistory(tag);
        } else {
            await summaryService.fetchSummaryAIHistory(tag);
        }
    }

    const sessions = tag.ChatSessions || [];

    // If targetChatId is provided from bookmark, locate its conversation thread and message
    if (targetChatId !== undefined && targetChatId !== null && String(targetChatId).trim() !== '') {
        const strChatId = String(targetChatId).trim();
        let foundSessionIndex = -1;
        let foundMessageIndex = -1;

        for (let sIdx = 0; sIdx < sessions.length; sIdx++) {
            const sess = sessions[sIdx];
            if (String(sess.chatHistoryId) === strChatId || strChatId.endsWith(`_${sess.chatHistoryId}`)) {
                foundSessionIndex = sIdx;
            }
            if (sess.history && sess.history.length > 0) {
                const mIdx = sess.history.findIndex((m: any) =>
                    String(m.ID) === strChatId ||
                    String(m.ReportHeadAIHistoryID) === strChatId ||
                    String(m.ChatHistoryID) === strChatId ||
                    strChatId.endsWith(`_${m.ID}`) ||
                    strChatId.endsWith(`_${m.ChatHistoryID}`)
                );
                if (mIdx !== -1) {
                    foundSessionIndex = sIdx;
                    foundMessageIndex = mIdx;
                    break;
                }
            }
        }

        if (foundSessionIndex !== -1) {
            tag.ActiveSessionIndex = foundSessionIndex;
        } else {
            // Chat ID not found in any session, open the most recent chat (index 0)
            tag.ActiveSessionIndex = 0;
        }
    } else {
        // No chat ID in bookmark: keep active session or default to 0
        if (tag.ActiveSessionIndex === undefined || tag.ActiveSessionIndex < 0 || tag.ActiveSessionIndex >= sessions.length) {
            tag.ActiveSessionIndex = 0;
        }
    }

    const activeIdx = (tag.ActiveSessionIndex !== undefined && tag.ActiveSessionIndex >= 0 && tag.ActiveSessionIndex < sessions.length)
        ? tag.ActiveSessionIndex
        : 0;

    const activeSession = sessions[activeIdx] || {
        id: 'default',
        chatHistoryId: 0,
        title: 'Conversation',
        createdAt: new Date().toISOString(),
        sources: tag.Sources || [],
        sourceValues: tag.TempSourceValue || [],
        history: tag.FilteredReportHeadAIHistoryList || []
    };

    const history = activeSession.history || [];
    tag.FilteredReportHeadAIHistoryList = history;
    tag.ReportHeadAIHistoryList = history;

    const chat = history.find((item: any) => item.Selected === 1);
    if (chat) {
        const finalResponse = chat.FormattedResponse
            ? '\n' + updateEditorFinalTable(chat.FormattedResponse)
            : chat.Response;

        tag.ComponentKeyDataType = chat.FormattedResponse ? 'TABLE' : 'TEXT';
        tag.UserValue = finalResponse;
        tag.EditorValue = finalResponse;
        tag.text = finalResponse;
        tag.IsApplied = false;
        tag.ChatHistoryID = chat.ChatHistoryID || chat.ID || chat.ReportHeadAIHistoryID || activeSession?.chatHistoryId || '';
    } else {
        // Active session does not have a selected prompt, check if another session has one
        const anySelectedSession = sessions.find((s: any) => s.history?.some((m: any) => m.Selected === 1));
        const anySelectedChat = anySelectedSession?.history?.find((m: any) => m.Selected === 1);
        if (anySelectedChat) {
            const finalResponse = anySelectedChat.FormattedResponse
                ? '\n' + updateEditorFinalTable(anySelectedChat.FormattedResponse)
                : anySelectedChat.Response;

            tag.ComponentKeyDataType = anySelectedChat.FormattedResponse ? 'TABLE' : 'TEXT';
            tag.UserValue = finalResponse;
            tag.EditorValue = finalResponse;
            tag.text = finalResponse;
            tag.IsApplied = false;
            tag.ChatHistoryID = anySelectedChat.ChatHistoryID || anySelectedChat.ID || anySelectedChat.ReportHeadAIHistoryID || anySelectedSession.chatHistoryId || '';
        } else {
            tag.UserValue = '';
            tag.EditorValue = '';
            tag.text = '';
            tag.IsApplied = true;
            tag.ChatHistoryID = '';
        }
    }

    // Check current theme
    const isDark = store.theme === 'Dark';
    const closeBtnClass = isDark
        ? 'fa-solid fa-circle-xmark bg-dark text-light'
        : 'fa-solid fa-circle-xmark bg-light text-dark';
    const jumpBtnColorClass = isDark ? 'text-light' : 'text-dark';
    const headerBgClass = isDark ? 'bg-dark text-light' : 'bg-white text-dark';
    const DisplayName = type === 'Summary' ? tag.Name : tag.DisplayName;

    // Session dropdown menu items
    const sessionMenuItems = sessions.map((s: any, idx: number) => {
        const isActive = idx === activeIdx;
        const msgCount = s.history ? s.history.length : 0;
        const dateStr = s.createdAt ? new Date(s.createdAt).toLocaleDateString(undefined, { month: 'short', day: 'numeric' }) : '';
        const rawTitle = s.title || `Chat ${idx + 1}`;
        const titleSafe = String(rawTitle).replace(/"/g, '&quot;');
        return `
            <li>
                <a class="dropdown-item chat-session-item d-flex align-items-center justify-content-between py-2 px-3 c-pointer ${isActive ? (isDark ? 'bg-secondary text-light' : 'bg-light text-primary fw-bold') : (isDark ? 'text-light' : 'text-dark')}" href="#" data-session-index="${idx}" title="${titleSafe}">
                    <div class="d-flex align-items-center text-truncate me-2" style="max-width: 200px;" title="${titleSafe}">
                        <i class="fa-regular ${isActive ? 'fa-comment-dots text-primary' : 'fa-message text-muted'} me-2"></i>
                        <span class="text-truncate ${isActive ? 'fw-bold' : 'fw-normal'}" title="${titleSafe}">${rawTitle}</span>
                    </div>
                    <div class="d-flex align-items-center gap-1 flex-shrink-0">
                        <span class="badge rounded-pill ${isDark ? 'bg-dark text-light' : 'bg-white text-muted border'} small" style="font-size: 9.5px;">${msgCount} msg${msgCount === 1 ? '' : 's'}</span>
                        ${dateStr ? `<span class="text-muted small" style="font-size: 9.5px;">${dateStr}</span>` : ''}
                    </div>
                </a>
            </li>
        `;
    }).join('');

    // Source pool & grouping for in-chat source dropdown
    const sourcePool = (type === 'Summary')
        ? (store.sourceSummaryList?.length ? store.sourceSummaryList : store.sourceList || [])
        : (store.sourceList?.length ? store.sourceList : store.sourceSummaryList || []);

    const sourceGroups: { [key: string]: any[] } = {};
    sourcePool.forEach((item: any) => {
        const group = item.GroupName || item.SourceType || 'Document Sources';
        if (!sourceGroups[group]) sourceGroups[group] = [];
        sourceGroups[group].push(item);
    });

    const activeSourcesList: string[] = activeSession.sources || tag.Sources || [];
    const activeSourceValuesList: string[] = activeSession.sourceValues || tag.TempSourceValue || [];

    const groupKeys = Object.keys(sourceGroups);
    let totalSourcesCount = 0;
    let selectedSourcesCount = 0;
    const selectedSourceNames: string[] = [];

    const groupedSourcesHtml = groupKeys.map((groupName, gIdx) => {
        const items = sourceGroups[groupName];
        let groupSelectedCount = 0;

        const itemsHtml = items.map((src: any, sIdx: number) => {
            totalSourcesCount++;
            const name = src.SourceName || src.FileName || src.SourceValue || `Source ${sIdx + 1}`;
            const val = src.VectorID ? String(src.VectorID) : (src.SourceValue || name);
            const safeName = String(name).replace(/"/g, '&quot;');
            const safeVal = String(val).replace(/"/g, '&quot;');

            const isChecked = activeSourcesList.some(s => String(s).toLowerCase() === String(name).toLowerCase()) ||
                activeSourceValuesList.some(v => String(v) === String(val) || String(v) === String(src.VectorID));

            if (isChecked) {
                selectedSourcesCount++;
                groupSelectedCount++;
                selectedSourceNames.push(name);
            }

            return `
                <li class="chat-source-item-row" data-search-text="${safeName.toLowerCase()}">
                    <div class="d-flex align-items-center gap-2 py-2 ps-4 pe-3 border-bottom ${isDark ? 'border-secondary' : 'border-light'}">
                        <input class="form-check-input chat-source-single-checkbox m-0 flex-shrink-0" type="checkbox" value="${safeVal}" data-source-name="${safeName}" data-group-index="${gIdx}" id="chat-src-${gIdx}-${sIdx}" ${isChecked ? 'checked' : ''}>
                        <label class="form-check-label small text-truncate flex-grow-1 m-0 c-pointer d-flex align-items-center gap-2" for="chat-src-${gIdx}-${sIdx}" title="${safeName}">
                            <i class="fa-regular fa-file-lines text-muted flex-shrink-0" style="font-size: 11.5px;"></i>
                            <span class="text-truncate" style="font-size: 12px;">${name}</span>
                        </label>
                    </div>
                </li>
            `;
        }).join('');

        const isGroupAllChecked = items.length > 0 && groupSelectedCount === items.length;

        return `
            <div class="chat-source-group-section mb-2.5">
                <div class="d-flex align-items-center gap-2 px-3 py-2 bg-light text-dark fw-bold small rounded-top border">
                    <input class="form-check-input chat-source-group-checkbox m-0 flex-shrink-0" type="checkbox" data-group-index="${gIdx}" id="chat-group-chk-${gIdx}" ${isGroupAllChecked ? 'checked' : ''}>
                    <label class="form-check-label text-truncate flex-grow-1 m-0 c-pointer" style="font-size: 12px;" for="chat-group-chk-${gIdx}">${groupName}</label>
                    <span class="badge bg-secondary text-white rounded-pill ms-auto chat-group-count-badge" style="font-size: 9px;">${groupSelectedCount}/${items.length}</span>
                </div>
                <ul class="list-unstyled p-0 m-0 border border-top-0 rounded-bottom">
                    ${itemsHtml}
                </ul>
            </div>
        `;
    }).join('');

    const isAllChecked = totalSourcesCount > 0 && selectedSourcesCount === totalSourcesCount;
    const sourcesNamesText = selectedSourcesCount === 0
        ? 'No sources selected'
        : selectedSourceNames.join(', ');
    const fullSourcesTitle = `Sources: ${sourcesNamesText}`;

    const activeTitleSafe = String(activeSession.title || `Chat ${activeIdx + 1}`).replace(/"/g, '&quot;');
    const displayNameSafe = String(DisplayName).replace(/"/g, '&quot;');

    const closeBar = `
    <div class="chat-header sticky-top ${headerBgClass} z-3">
        <!-- Main Tag Header (Selected Text Tag Bar) -->
        <div class="d-flex justify-content-between align-items-center px-3 pt-2 pb-1">
            <div class="d-flex align-items-center flex-grow-1 text-truncate" style="max-width: calc(100% - 85px);" title="${displayNameSafe}">
                ${getIconSvg(faMicrochipAi, 'text-muted me-2', 'font-size: 13px;')}
                <span class="fw-bold text-truncate" style="font-size: 13px; line-height: 1.4; letter-spacing: 0.3px;" title="${displayNameSafe}">${DisplayName}</span>
            </div>
            <div class="d-flex align-items-center ms-2" style="margin-top: 2px;">
                <!-- Toggle Hide/Show Icon for Chat Selector & New Chat -->
                <button id="toggleChatControlsBtn" class="btn btn-sm p-0 me-2 border-0 bg-transparent ${jumpBtnColorClass} c-pointer" title="Hide chat controls" style="display: inline-flex; align-items: center; justify-content: center; width: 20px; height: 20px;">
                    <i class="fa-solid fa-chevron-up" id="toggleControlsIcon" style="font-size: 11px; transition: transform 0.2s ease;"></i>
                </button>
                <button id="jump-to-next-tag" class="btn btn-sm p-0 me-2 border-0 bg-transparent ${jumpBtnColorClass} c-pointer" title="Jump to next replaced instance" style="display: inline-flex; align-items: center; justify-content: center; transition: transform 0.2s ease;">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                        <circle cx="12" cy="12" r="10" />
                        <path d="M 15 14 L 15 12.5 A 2.5 2.5 0 0 0 12.5 10 L 9 10" />
                        <polyline points="11 8 9 10 11 12" />
                    </svg>
                </button>
                <div class="c-pointer d-inline-flex align-items-center justify-content-center" id="close-btn-tag" title="Close AI chat">
                    <i class="${closeBtnClass}" id="close-ai-window" style="font-size: 13px;"></i>
                </div>
            </div>
        </div>
        <hr class="mt-1 mb-0 mx-3">

        <!-- Collapsible Controls Panel (Session Selector, New Chat, and Sources Dropdown) -->
        <div id="chatControlsPanel" class="chat-controls-panel">

            <!-- Multi-Session Toolbar -->
            <div class="chat-session-toolbar px-3 py-2 border-bottom ${headerBgClass} d-flex align-items-center justify-content-between gap-2 shadow-xs">
                <!-- Session Selector Dropdown -->
                <div class="dropdown flex-grow-1" style="min-width: 0;">
                    <button class="btn btn-sm ${isDark ? 'btn-outline-secondary text-light' : 'btn-outline-secondary text-dark'} dropdown-toggle w-100 text-start d-flex align-items-center justify-content-between text-truncate chat-thread-selector-btn"
                            type="button" id="chatSessionDropdown" data-bs-toggle="dropdown" aria-expanded="false" title="${activeTitleSafe}" style="background: ${isDark ? '#2b3035' : '#f8f9fa'}; border-color: ${isDark ? '#495057' : '#dee2e6'};">
                        <span class="text-truncate d-flex align-items-center me-2" title="${activeTitleSafe}">
                            <i class="fa-regular fa-message me-2 text-primary"></i>
                            <span class="fw-semibold active-chat-title text-truncate" title="${activeTitleSafe}">${activeSession.title || `Chat ${activeIdx + 1}`}</span>
                        </span>
                    </button>
                    <ul class="dropdown-menu shadow w-100 p-1 chat-session-menu ${isDark ? 'dropdown-menu-dark' : ''}" aria-labelledby="chatSessionDropdown" style="max-height: 240px; overflow-y: auto; z-index: 1060;">
                        <li class="dropdown-header small text-muted px-3 py-1 fw-bold text-uppercase" style="font-size: 10px; letter-spacing: 0.5px;">Conversations (${sessions.length})</li>
                        ${sessionMenuItems}
                    </ul>
                </div>

                <!-- + New Chat Button -->
                <button class="btn btn-sm btn-primary text-white d-flex align-items-center text-nowrap new-chat-btn shadow-sm px-2 py-1" id="addNewChatBtn" title="Start a new chat with selected sources">
                    <i class="fa fa-plus me-1" style="font-size: 11px;"></i>
                    <span class="fw-semibold" style="font-size: 11.5px;">New Chat</span>
                </button>
            </div>

            <!-- Active Sources Bar (Matching Image 2) -->
            <div class="chat-sources-toolbar border-bottom ${isDark ? 'border-secondary' : 'border-light-subtle'}" style="background-color: ${isDark ? '#212529' : '#f8f9fa'};">
                <div class="dropdown w-100" id="chatSourceDropdownContainer">
                    <button class="btn btn-sm w-100 text-start d-flex align-items-center justify-content-between p-0 border-0 bg-transparent chat-sources-trigger-btn"
                            type="button" id="chatSourcesDropdown" data-bs-toggle="dropdown" data-bs-auto-close="outside" aria-expanded="false"
                            title="${fullSourcesTitle}" style="padding: 7px 14px !important;">
                        <div class="d-flex align-items-center text-truncate me-2" style="min-width: 0;">
                            <i class="fa-solid fa-file-lines me-2 flex-shrink-0" style="font-size: 13.5px; color: #337ab7;"></i>
                            <span class="fw-bold me-1.5 flex-shrink-0 ${isDark ? 'text-light' : 'text-dark'}" style="font-size: 12px;">Sources:</span>
                            <span class="text-truncate ${isDark ? 'text-light-50' : 'text-muted'}" id="chatSourcesDropdownText" style="font-size: 12px;">${sourcesNamesText}</span>
                        </div>
                        <span class="badge rounded-circle flex-shrink-0 d-inline-flex align-items-center justify-content-center" id="chatSourcesCountBadge" style="width: 19px; height: 19px; min-width: 19px; font-size: 10.5px; font-weight: 600; background-color: #6c757d; color: #ffffff; padding: 0;">${selectedSourcesCount}</span>
                    </button>
                    <ul class="dropdown-menu shadow w-100 p-2.5 chat-sources-menu ${isDark ? 'dropdown-menu-dark' : ''}" aria-labelledby="chatSourcesDropdown" style="max-height: 290px; overflow-y: auto; z-index: 1060; min-width: 270px; margin-top: 1px;">
                        <!-- Select All item -->
                        <li class="pb-2 border-bottom ${isDark ? 'border-secondary' : 'border-light'} mb-2">
                            <div class="d-flex align-items-center gap-2 px-3 py-1.5">
                                <input class="form-check-input m-0 flex-shrink-0" type="checkbox" id="chat-sources-select-all" ${isAllChecked ? 'checked' : ''}>
                                <label class="form-check-label small fw-bold m-0 c-pointer" style="font-size: 12px;" for="chat-sources-select-all">Select All Sources</label>
                            </div>
                        </li>
                        <!-- Grouped Sources items -->
                        ${groupedSourcesHtml || '<li class="p-2 text-center text-muted small">No sources available</li>'}
                    </ul>
                </div>
            </div>
        </div>
    </div>
    `;

    const chatSourcesForDisplay = selectedSourceNames.length > 0 ? selectedSourceNames : activeSourcesList;

    const chatBody = `
        <div class="chat-body flex-grow-1 d-flex flex-column overflow-auto">
            ${generateChatHistoryHtml(history, chatSourcesForDisplay)}
        </div>
    `;

    const chatFooterHtml = `
        <div class="d-flex align-items-end justify-content-end chatbox p-2" id="chatFooter">
            ${chatfooter(tag)}
        </div>
    `;

    const horizontalLoader = (String(tag.Status) === "0") ? '<div class="horizontal-loader"></div>' : '';

    initializeAIHistoryEvents(tag, store.jwt, store.availableKeys, type);

    return `${closeBar}${chatBody}${horizontalLoader}${chatFooterHtml}`;
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
    <button id="applyBtn" class="btn btn-primary text-white" disabled>Insert</button>
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
        const jwt = DocStorage.getItem('token') || '';

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


export async function insertTagPrompt(tag, type: "Summary" | "AITag" = "AITag") {
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
                        for (const line of txt.split(/\r?\n/)) {
                            if (!line.trim()) {
                                const p = cursor.insertParagraph("", Word.InsertLocation.after);
                                insertLineWithHeadingStyle(p, "");
                                include(p.getRange());
                                cursor = p.getRange();
                                continue;
                            }

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

                            let grid = parseHtmlTableToGrid(rows);
                            const tableCase = detectTableCase(grid);
                            const store = StoreService.getInstance();
                            const base = store.tableStyle.split(" - ")[0].trim();

                            if (base === 'Table Grid 2') {
                                store.isReversed = true;
                            } else {
                                store.isReversed = false;
                            }

                            if (store.isReversed && tableCase !== 'CASE_1') {
                                grid = transposeGrid(grid);
                            }

                            const numRows = grid.length;
                            const numCols = grid[0]?.length || 0;

                            const p = cursor.insertParagraph("", Word.InsertLocation.after);
                            const table = p.insertTable(numRows, numCols, Word.InsertLocation.after);

                            const resolvedTableStyle = resolveWordTableStyle(store.tableStyle);
                            if (resolvedTableStyle !== 'none') {
                                table.style = resolvedTableStyle;
                            }

                            // Population/Merging logic
                            if (!store.isReversed) {
                                if (!store.colorPallete.Customize) {
                                    grid.forEach((rowGrid, rowIndex) => {
                                        rowGrid.forEach((cellValue, cellIndex) => {
                                            const tableCell = table.getCell(rowIndex, cellIndex);
                                            tableCell.value = cellValue;
                                            applyCustomTextStyleToCell(tableCell, store);
                                        });
                                    });

                                    // Vertical merging logic for 1st column (standard view)
                                    let lastParamRowIndex = -1;
                                    grid.forEach((rowGrid, rowIndex) => {
                                        if (rowIndex === 0) return; // skip header
                                        const firstColText = rowGrid[0];
                                        if (firstColText) {
                                            lastParamRowIndex = rowIndex;
                                        } else if (lastParamRowIndex !== -1) {
                                            const topCell = table.getCell(lastParamRowIndex, 0);
                                            const bottomCell = table.getCell(rowIndex, 0);
                                            topCell.merge(bottomCell);
                                            try {
                                                topCell.verticalAlignment = Word.VerticalAlignment.center;
                                                topCell.body.paragraphs.getFirst().alignment = Word.Alignment.center;
                                            } catch (e) { }
                                        }
                                    });
                                }
                            } else {
                                // Manual population for transposed table
                                grid.forEach((rowGrid, rowIndex) => {
                                    rowGrid.forEach((cellValue, cellIndex) => {
                                        const tableCell = table.getCell(rowIndex, cellIndex);
                                        tableCell.value = cellValue;
                                        applyCustomTextStyleToCell(tableCell, store);
                                    });
                                });
                            }

                            // Styling logic
                            if (store.colorPallete.Customize) {
                                await colorTable(table, rows, context, store.isReversed);
                            }


                            include(table.getRange());
                            cursor = table.getRange();
                        }

                        // OTHER HTML ELEMENTS
                        else {
                            let txt = el.innerText?.trim();
                            if (!txt) continue;

                            txt = txt.replace(/\n- /g, "\n• ");
                            for (const line of txt.split(/\r?\n/)) {
                                if (!line.trim()) {
                                    const p = cursor.insertParagraph("", Word.InsertLocation.after);
                                    insertLineWithHeadingStyle(p, "");
                                    include(p.getRange());
                                    cursor = p.getRange();
                                    continue;
                                }

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

                for (const line of txt.split(/\r?\n/)) {
                    if (!line.trim()) {
                        const p = cursor.insertParagraph("", Word.InsertLocation.after);
                        insertLineWithHeadingStyle(p, "");
                        include(p.getRange());
                        cursor = p.getRange();
                        continue;
                    }

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
                const prefix = type === "Summary" ? "SM" : "ID";
                const tagId = tag.ID || tag.ReportHeadSummaryTagID || tag.ReportHeadGroupKeyID;

                const allSessions = tag.ChatSessions || [];
                let chatId = '';

                if (allSessions && allSessions.length > 0) {
                    // Find the selected/checked chat message across all sessions or history
                    let checkedChat: any = null;
                    for (const s of allSessions) {
                        const found = s.history?.find((item: any) => item.Selected === 1);
                        if (found) {
                            checkedChat = found;
                            break;
                        }
                    }
                    if (!checkedChat) {
                        const fallbackHistory = tag.FilteredReportHeadAIHistoryList || tag.ReportHeadAIHistoryList || [];
                        checkedChat = fallbackHistory.find((item: any) => item.Selected === 1);
                    }

                    chatId = checkedChat?.ChatHistoryID || checkedChat?.ID || checkedChat?.ReportHeadAIHistoryID || tag.ChatHistoryID || '';
                } else {
                    chatId = tag.ChatHistoryID || '';
                }
                const chatSuffix = chatId ? `_${chatId}` : '';
                const bookmarkName =
                    `${prefix}${tagId}_Split_${getDateTimeStamp()}${chatSuffix}`;

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




export async function openPromptBuilderModal(tag: any, type: "Summary" | "AITag") {
    const container = document.getElementById('confirmation-popup');
    if (!container) return;

    // Show the modal
    container.innerHTML = PromptBuilderModalPopup();

    const store = StoreService.getInstance();
    const promptBuilderList = store.promptBuilderList || [];

    // References to modal elements
    const templateSelect = document.getElementById('promptBuilderTemplatePopup') as HTMLSelectElement;
    const insertBtn = document.getElementById('prompt-builder-popup-insert') as HTMLButtonElement;
    const cancelBtn = document.getElementById('prompt-builder-popup-cancel') as HTMLButtonElement;
    const previewDiv = document.getElementById('previewPopup') as HTMLDivElement;
    const fieldsContainer = document.getElementById('fieldsContainerPopup') as HTMLDivElement;
    const previewContainer = document.getElementById('previewContainerPopup') as HTMLDivElement;
    const templateError = document.getElementById('templateErrorPopup') as HTMLDivElement;

    let fieldsList: any[] = [];
    let templateText = '';
    let currentPreview = '';

    // Populate template dropdown
    promptBuilderList.forEach((item) => {
        const option = document.createElement('option');
        option.value = item.ID.toString();
        option.textContent = item.Name;
        templateSelect.appendChild(option);
    });

    // Close/Cancel functions
    const closeModal = () => {
        container.innerHTML = '';
    };

    cancelBtn.addEventListener('click', closeModal);

    // Template change logic
    templateSelect.addEventListener('change', async () => {
        const templateId = templateSelect.value;
        const jwt = DocStorage.getItem('token') || '';

        try {
            const data = await getPromptTemplateById(templateId, jwt);
            if (data.Status && data.Data) {
                fieldsList = data.Data;
                const selectedTemplateObj = promptBuilderList.find((item) => item.ID.toString() === templateId);
                templateText = selectedTemplateObj ? selectedTemplateObj.Template : '';
                currentPreview = templateText;
            }
            if (!templateId) {
                templateError.classList.remove('d-none');
                return;
            }

            templateError.classList.add('d-none');

            renderFields();
            updatePreview();
        } catch (err) {
            console.error("Failed to load prompt template fields:", err);
            toaster("Failed to load template fields", "error");
        }
    });

    function renderFields() {
        fieldsContainer.innerHTML = '';

        fieldsList.forEach((field) => {
            const div = document.createElement('div');
            div.className = 'form-group mb-3';

            const label = document.createElement('label');
            label.className = 'form-label fw-semibold small';
            label.textContent = field.Label;
            div.appendChild(label);

            if (field.Type === 1) {
                const input = document.createElement('input');
                input.type = 'text';
                input.className = 'form-control';
                input.id = `modal-field-${field.Label}`;
                input.addEventListener('input', replaceKeywordsManually);
                div.appendChild(input);
            } else if (field.Type === 2) {
                const select = document.createElement('select');
                select.className = 'form-select';
                select.id = `modal-field-${field.Label}`;
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
            const id = `modal-field-${field.Label}`;
            const keyword = `#${field.Label}#`;

            let value = '';
            const element = document.getElementById(id) as HTMLInputElement | HTMLSelectElement;

            if (element) {
                value = element.value || '';
            }

            keywordMap[keyword] = value ? value : keyword;
        });

        let insertValue = templateText;
        for (const [keyword, value] of Object.entries(keywordMap)) {
            insertValue = insertValue.replace(new RegExp(keyword, 'g'), value);
        }

        currentPreview = insertValue;
        previewDiv.textContent = currentPreview;
        previewContainer.style.display = currentPreview ? 'block' : 'none';
        insertBtn.disabled = currentPreview === '';
    }

    function updatePreview() {
        replaceKeywordsManually();
    }

    // Insert logic: copy generated template text and insert into #chatInput
    insertBtn.addEventListener('click', () => {
        if (!currentPreview) return;
        const chatInput = document.getElementById('chatInput') as HTMLTextAreaElement;
        if (chatInput) {
            chatInput.value = currentPreview;
            chatInput.dispatchEvent(new Event('input', { bubbles: true }));
        }
        closeModal();
    });
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
                // Enabled by default
                changeSourceButton.disabled = false;
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
                                    let data: any;
                                    if (type === 'Summary') {
                                        const payload = {
                                            Name: tag.Name,
                                            Prompt: chat.Prompt
                                        };
                                        data = await updateSummaryTagPrompt(payload, jwt);
                                    } else {
                                        let updatedTag = JSON.parse(JSON.stringify(tag));
                                        updatedTag.Prompt = chat.Prompt;
                                        data = await updatePromptTemplate(updatedTag, jwt);
                                    }
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

                    // Reset all DOM checkboxes and visual styles for current active view
                    tag.FilteredReportHeadAIHistoryList.forEach((_: any, otherIndex: number) => {
                        const otherCheckbox = document.getElementById(`checkbox-${otherIndex}`) as HTMLInputElement;
                        const responseContainer = document.getElementById(`responseContainer-${otherIndex}`);
                        if (otherCheckbox) otherCheckbox.checked = false;
                        if (responseContainer) {
                            responseContainer.classList.remove('ai-selected-response');
                            responseContainer.classList.add('bg-light');
                        }
                    });

                    // Set Selected = 0 across ALL messages in ALL sessions of this tag
                    if (tag.ChatSessions && Array.isArray(tag.ChatSessions)) {
                        tag.ChatSessions.forEach((s: any) => {
                            s.history?.forEach((m: any) => {
                                m.Selected = 0;
                            });
                        });
                    }
                    if (tag.ReportHeadAIHistoryList && Array.isArray(tag.ReportHeadAIHistoryList)) {
                        tag.ReportHeadAIHistoryList.forEach((m: any) => {
                            m.Selected = 0;
                        });
                    }
                    if (tag.FilteredReportHeadAIHistoryList && Array.isArray(tag.FilteredReportHeadAIHistoryList)) {
                        tag.FilteredReportHeadAIHistoryList.forEach((m: any) => {
                            m.Selected = 0;
                        });
                    }

                    // Set selected on target chat
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

                    const isTable = chat.FormattedResponse && chat.FormattedResponse !== '';
                    const finalResponse = isTable
                        ? '\n' + updateEditorFinalTable(chat.FormattedResponse)
                        : (chat.Response || '');

                    tag.ComponentKeyDataType = isTable ? 'TABLE' : 'TEXT';
                    tag.UserValue = isChecked ? finalResponse : '';
                    tag.EditorValue = isChecked ? finalResponse : '';
                    tag.text = isChecked ? finalResponse : '';
                    tag.IsApplied = !isChecked;
                    tag.ChatHistoryID = isChecked ? (chat.ChatHistoryID || chat.ID || chat.ReportHeadAIHistoryID || tag.ChatSessions?.[tag.ActiveSessionIndex || 0]?.chatHistoryId || '') : '';

                    // Sync across store lists immediately
                    const store = StoreService.getInstance();
                    const tagId = tag.ID || tag.ReportHeadGroupKeyID || tag.ReportHeadSummaryTagID;

                    if (type === 'Summary') {
                        store.summaryTagList?.forEach(currentTag => {
                            const currentId = currentTag.ID || currentTag.ReportHeadSummaryTagID;
                            if (currentTag === tag || currentId === tagId || currentTag.Name === tag.Name || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName)) {
                                currentTag.ComponentKeyDataType = tag.ComponentKeyDataType;
                                currentTag.UserValue = tag.UserValue;
                                currentTag.EditorValue = tag.EditorValue;
                                currentTag.text = tag.text;
                                currentTag.IsApplied = tag.IsApplied;
                                currentTag.ChatSessions = tag.ChatSessions;
                                currentTag.ActiveSessionIndex = tag.ActiveSessionIndex;
                                currentTag.FilteredReportHeadAIHistoryList = tag.FilteredReportHeadAIHistoryList;
                                currentTag.ReportHeadAIHistoryList = tag.ReportHeadAIHistoryList;
                                currentTag.ChatHistoryID = tag.ChatHistoryID;
                            }
                        });
                    } else {
                        availableKeys?.forEach(currentTag => {
                            const currentId = currentTag.ID || currentTag.ReportHeadGroupKeyID;
                            if (currentTag === tag || currentId === tagId || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName) || (currentTag.GroupKey && currentTag.GroupKey === tag.GroupKey)) {
                                currentTag.ComponentKeyDataType = tag.ComponentKeyDataType;
                                currentTag.UserValue = tag.UserValue;
                                currentTag.EditorValue = tag.EditorValue;
                                currentTag.text = tag.text;
                                currentTag.IsApplied = tag.IsApplied;
                                currentTag.ChatSessions = tag.ChatSessions;
                                currentTag.ActiveSessionIndex = tag.ActiveSessionIndex;
                                currentTag.FilteredReportHeadAIHistoryList = tag.FilteredReportHeadAIHistoryList;
                                currentTag.ReportHeadAIHistoryList = tag.ReportHeadAIHistoryList;
                                currentTag.ChatHistoryID = tag.ChatHistoryID;
                            }
                        });

                        store.aiTagList?.forEach(currentTag => {
                            const currentId = currentTag.ID || currentTag.ReportHeadGroupKeyID;
                            if (currentTag === tag || currentId === tagId || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName) || (currentTag.GroupKey && currentTag.GroupKey === tag.GroupKey)) {
                                currentTag.ComponentKeyDataType = tag.ComponentKeyDataType;
                                currentTag.UserValue = tag.UserValue;
                                currentTag.EditorValue = tag.EditorValue;
                                currentTag.text = tag.text;
                                currentTag.IsApplied = tag.IsApplied;
                                currentTag.ChatSessions = tag.ChatSessions;
                                currentTag.ActiveSessionIndex = tag.ActiveSessionIndex;
                                currentTag.FilteredReportHeadAIHistoryList = tag.FilteredReportHeadAIHistoryList;
                                currentTag.ReportHeadAIHistoryList = tag.ReportHeadAIHistoryList;
                                currentTag.ChatHistoryID = tag.ChatHistoryID;
                            }
                        });
                    }

                    try {
                        const data = type === 'Summary' ? await updateSummaryHistory(chat, jwt) : await updateAiHistory(chat, jwt);
                        if (data['Data']) {
                            const rawData = Array.isArray(data['Data']) ? data['Data'] : [data['Data']];
                            const sessions = AIService.groupHistoryIntoSessions(rawData, tag, store, type);

                            const currentActiveChatHistoryId = tag.ChatSessions?.[tag.ActiveSessionIndex || 0]?.chatHistoryId;
                            tag.ChatSessions = sessions;
                            if (currentActiveChatHistoryId !== undefined) {
                                const foundIdx = sessions.findIndex(s => Number(s.chatHistoryId) === Number(currentActiveChatHistoryId));
                                tag.ActiveSessionIndex = foundIdx !== -1 ? foundIdx : 0;
                            } else {
                                tag.ActiveSessionIndex = 0;
                            }

                            // Ensure ONLY this chat is selected across all sessions
                            tag.ChatSessions.forEach((s: any) => {
                                s.history?.forEach((m: any) => {
                                    const isMatch = (chat.ID && m.ID && String(m.ID) === String(chat.ID)) ||
                                        (chat.ReportHeadAIHistoryID && m.ReportHeadAIHistoryID && String(m.ReportHeadAIHistoryID) === String(chat.ReportHeadAIHistoryID)) ||
                                        (m.Response === chat.Response && m.Prompt === chat.Prompt);
                                    m.Selected = (isMatch && isChecked) ? 1 : 0;
                                });
                            });

                            const activeSession = tag.ChatSessions[tag.ActiveSessionIndex] || tag.ChatSessions[0];
                            if (activeSession) {
                                tag.FilteredReportHeadAIHistoryList = activeSession.history;
                                tag.ReportHeadAIHistoryList = activeSession.history;
                            }
                        }
                    } catch (err) {
                        console.error('Failed to update AI history:', err);
                    }
                });
            }
        });

        // Chat Session Switcher items click
        document.querySelectorAll('.chat-session-item').forEach((elem: Element) => {
            elem.addEventListener('click', (e) => {
                e.preventDefault();
                const sessionIdx = parseInt(elem.getAttribute('data-session-index') || '0', 10);
                if (sessionIdx === (tag.ActiveSessionIndex || 0)) return;

                confirmSwitchChatHistory(async () => {
                    AIService.switchChatSession(tag, sessionIdx);
                    const appBody = document.getElementById('app-body');
                    if (appBody) {
                        appBody.innerHTML = await generateCheckboxHistory(tag, type);
                    }
                });
            });
        });

        // + New Chat button click
        document.getElementById('addNewChatBtn')?.addEventListener('click', () => {
            openNewChatModal(tag, type);
        });

        // In-chat Source Dropdown checkboxes logic
        const chatSourcesSelectAll = document.getElementById('chat-sources-select-all') as HTMLInputElement;
        const chatSourceGroupBoxes = document.querySelectorAll('.chat-source-group-checkbox');
        const chatSourceSingleBoxes = document.querySelectorAll('.chat-source-single-checkbox');
        const chatSourcesText = document.getElementById('chatSourcesDropdownText');
        const chatSourcesBadge = document.getElementById('chatSourcesCountBadge');
        const chatSourcesDropdown = document.getElementById('chatSourcesDropdown');

        const updateInChatSources = () => {
            const selectedNames: string[] = [];
            const selectedValues: string[] = [];

            chatSourceSingleBoxes.forEach((cb: Element) => {
                const input = cb as HTMLInputElement;
                if (input.checked) {
                    selectedValues.push(input.value);
                    const name = input.dataset.sourceName || input.value;
                    selectedNames.push(name);
                }
            });

            // Update active session & tag
            const activeSession = tag.ChatSessions && tag.ActiveSessionIndex !== undefined
                ? tag.ChatSessions[tag.ActiveSessionIndex]
                : null;
            if (activeSession) {
                activeSession.sources = [...selectedNames];
                activeSession.sourceValues = [...selectedValues];
            }
            tag.Sources = [...selectedNames];
            tag.SourceName = [...selectedNames];
            tag.TempSourceValue = [...selectedValues];
            tag.SourceValueID = selectedValues;

            const count = selectedNames.length;
            const newNamesText = count === 0
                ? 'No sources selected'
                : selectedNames.join(', ');
            const fullTitle = `Sources: ${newNamesText}`;

            if (chatSourcesText) chatSourcesText.innerText = newNamesText;
            if (chatSourcesBadge) chatSourcesBadge.innerText = String(count);
            if (chatSourcesDropdown) chatSourcesDropdown.setAttribute('title', fullTitle);

            // Update empty state description if present
            const emptyStateDesc = document.querySelector('.chat-empty-state-desc');
            if (emptyStateDesc) {
                let descHtml = 'Ask questions, draft summaries, or generate structured tables.';
                if (count === 1) {
                    descHtml = `Configured with <strong>1 source</strong> (${selectedNames[0]}). Ask questions, draft summaries, or generate structured tables.`;
                } else if (count === 2) {
                    descHtml = `Configured with <strong>2 sources</strong> (${selectedNames[0]}, ${selectedNames[1]}). Ask questions, draft summaries, or generate structured tables.`;
                } else if (count > 2) {
                    const remaining = count - 2;
                    descHtml = `Configured with <strong>${count} sources</strong> (${selectedNames[0]}, ${selectedNames[1]} +${remaining} more). Ask questions, draft summaries, or generate structured tables.`;
                }
                emptyStateDesc.innerHTML = descHtml;
            }

            // Update group counts & group checkbox states
            chatSourceGroupBoxes.forEach((gcb: Element) => {
                const groupInput = gcb as HTMLInputElement;
                const groupIdx = groupInput.dataset.groupIndex;
                const groupItems = document.querySelectorAll(`.chat-source-single-checkbox[data-group-index="${groupIdx}"]`);
                const checkedGroupItems = document.querySelectorAll(`.chat-source-single-checkbox[data-group-index="${groupIdx}"]:checked`);
                groupInput.checked = groupItems.length > 0 && checkedGroupItems.length === groupItems.length;
                const groupBadge = groupInput.parentElement?.querySelector('.chat-group-count-badge');
                if (groupBadge) {
                    groupBadge.textContent = `${checkedGroupItems.length}/${groupItems.length}`;
                }
            });

            if (chatSourcesSelectAll) {
                chatSourcesSelectAll.checked = chatSourceSingleBoxes.length > 0 && Array.from(chatSourceSingleBoxes).every((item: Element) => (item as HTMLInputElement).checked);
            }

            // Sync with tag in store
            const store = StoreService.getInstance();
            const tagId = tag.ID || tag.ReportHeadGroupKeyID || tag.ReportHeadSummaryTagID;
            store.aiTagList?.forEach((t: any) => {
                const curId = t.ID || t.ReportHeadGroupKeyID;
                if (t === tag || curId === tagId || t.DisplayName === tag.DisplayName) {
                    t.Sources = [...selectedNames];
                    t.SourceName = [...selectedNames];
                    t.TempSourceValue = [...selectedValues];
                    t.SourceValueID = selectedValues;
                    if (t.ChatSessions && t.ActiveSessionIndex !== undefined && t.ChatSessions[t.ActiveSessionIndex]) {
                        t.ChatSessions[t.ActiveSessionIndex].sources = [...selectedNames];
                        t.ChatSessions[t.ActiveSessionIndex].sourceValues = [...selectedValues];
                    }
                }
            });
            store.availableKeys?.forEach((t: any) => {
                const curId = t.ID || t.ReportHeadGroupKeyID;
                if (t === tag || curId === tagId || t.DisplayName === tag.DisplayName) {
                    t.Sources = [...selectedNames];
                    t.SourceName = [...selectedNames];
                    t.TempSourceValue = [...selectedValues];
                    t.SourceValueID = selectedValues;
                    if (t.ChatSessions && t.ActiveSessionIndex !== undefined && t.ChatSessions[t.ActiveSessionIndex]) {
                        t.ChatSessions[t.ActiveSessionIndex].sources = [...selectedNames];
                        t.ChatSessions[t.ActiveSessionIndex].sourceValues = [...selectedValues];
                    }
                }
            });
            store.summaryTagList?.forEach((t: any) => {
                const curId = t.ID || t.ReportHeadSummaryTagID;
                if (t === tag || curId === tagId || t.Name === tag.Name) {
                    t.Sources = [...selectedNames];
                    t.TempSourceValue = [...selectedValues];
                    t.SourceValueID = selectedValues;
                    if (t.ChatSessions && t.ActiveSessionIndex !== undefined && t.ChatSessions[t.ActiveSessionIndex]) {
                        t.ChatSessions[t.ActiveSessionIndex].sources = [...selectedNames];
                        t.ChatSessions[t.ActiveSessionIndex].sourceValues = [...selectedValues];
                    }
                }
            });
        };

        // Select All listener
        chatSourcesSelectAll?.addEventListener('change', () => {
            const isChecked = chatSourcesSelectAll.checked;
            chatSourceSingleBoxes.forEach((cb: Element) => ((cb as HTMLInputElement).checked = isChecked));
            chatSourceGroupBoxes.forEach((gcb: Element) => ((gcb as HTMLInputElement).checked = isChecked));
            updateInChatSources();
        });

        // Group checkboxes listener
        chatSourceGroupBoxes.forEach((gcb: Element) => {
            gcb.addEventListener('change', () => {
                const groupInput = gcb as HTMLInputElement;
                const groupIdx = groupInput.dataset.groupIndex;
                const childBoxes = document.querySelectorAll(`.chat-source-single-checkbox[data-group-index="${groupIdx}"]`);
                childBoxes.forEach((cb: Element) => ((cb as HTMLInputElement).checked = groupInput.checked));
                updateInChatSources();
            });
        });

        // Single checkbox listener
        chatSourceSingleBoxes.forEach((cb: Element) => {
            cb.addEventListener('change', () => {
                updateInChatSources();
            });
        });

        // Toggle Chat Controls (Hide/Show selector dropdown, new chat, and sources banner)
        const toggleChatControlsBtn = document.getElementById('toggleChatControlsBtn');
        const chatControlsPanel = document.getElementById('chatControlsPanel');
        const toggleControlsIcon = document.getElementById('toggleControlsIcon');

        toggleChatControlsBtn?.addEventListener('click', (e) => {
            e.preventDefault();
            if (!chatControlsPanel) return;
            const isHidden = chatControlsPanel.classList.contains('d-none');
            if (isHidden) {
                chatControlsPanel.classList.remove('d-none');
                if (toggleControlsIcon) {
                    toggleControlsIcon.classList.remove('fa-chevron-down');
                    toggleControlsIcon.classList.add('fa-chevron-up');
                }
                toggleChatControlsBtn.setAttribute('title', 'Hide chat & source controls');
            } else {
                chatControlsPanel.classList.add('d-none');
                if (toggleControlsIcon) {
                    toggleControlsIcon.classList.remove('fa-chevron-up');
                    toggleControlsIcon.classList.add('fa-chevron-down');
                }
                toggleChatControlsBtn.setAttribute('title', 'Show chat & source controls');
            }
        });

        // Jump to next bookmark button
        document.getElementById(`jump-to-next-tag`)?.addEventListener('click', async () => {
            await jumpToNextBookmarkOfTag(tag, type);
        });

        // Close button
        document.getElementById(`close-btn-tag`)?.addEventListener('click', () => {
            confirmSwitchChatHistory(() => {
                const store = StoreService.getInstance();
                store.currentChatTagId = -1;
                DocStorage.setItem("currentChatTagId", "-1");
                if (store.mode === "Home") {
                    loadHomepage(availableKeys)
                } else if (store.mode === "Summary") {
                    loadSummarypage(availableKeys);
                }
            });
        });

        // Button: Prompt Builder
        document.getElementById(`promptBuilderButton`)?.addEventListener('click', () => {
            openPromptBuilderModal(tag, type);
        });

        // Button: Insert Tag
        document.getElementById(`insertTagButton`)?.addEventListener('click', () => {
            if (!tag.IsApplied) {
                insertTagPrompt(tag, type);
            }
        });

        // Button: Send Prompt
        document.getElementById(`sendPromptButton`)?.addEventListener('click', () => {
            const textareaValue = (document.getElementById(`chatInput`) as HTMLTextAreaElement).value;
            AIService.sendPrompt(tag, textareaValue, type);
        });

        // Mention dropdown
        mentionDropdownFn(`chatInput`, `mention-dropdown`, 'edit');
    }, 0);
}

export async function jumpToNextBookmarkOfTag(tag: any, type: "Summary" | "AITag") {
    return Word.run(async (context) => {
        const selection = context.document.getSelection();
        const bodyRange = context.document.body.getRange();
        const bookmarks = bodyRange.getBookmarks();
        await context.sync();

        const bookmarkNames = bookmarks.value || [];
        const prefix = type === "Summary" ? "SM" : "ID";
        const tagId = tag.ID || tag.ReportHeadSummaryTagID;
        const targetPrefix = `${prefix}${tagId}_Split_`;

        // Filter relevant bookmarks (case-insensitive start match)
        const relevantNames = bookmarkNames.filter(name =>
            name.toUpperCase().startsWith(targetPrefix.toUpperCase())
        );

        if (relevantNames.length === 0) {
            toaster("No replaced instances found for this tag in the document.", "info");
            return;
        }

        // Get range and compare relation for each
        const items = relevantNames.map(name => {
            const r = context.document.getBookmarkRangeOrNullObject(name);
            r.load("isNullObject");
            const rel = r.compareLocationWith(selection);
            return { name, range: r, rel };
        });

        await context.sync();

        const validItems = items.filter(item => !item.range.isNullObject);
        if (validItems.length === 0) {
            toaster("No active instances found for this tag.", "info");
            return;
        }

        // Find the first bookmark that is positioned after the current selection
        let nextIndex = validItems.findIndex(item => {
            const val = String(item.rel.value).toLowerCase();
            return val === "after" || val === "adjacentafter";
        });

        // If none is after, wrap around to the first one
        if (nextIndex === -1) {
            nextIndex = 0;
        }

        const targetItem = validItems[nextIndex];
        targetItem.range.select();
        await context.sync();

        toaster(`Jumped to instance ${nextIndex + 1} of ${validItems.length}`, "success");
    });
}

/**
 * Opens modal allowing user to select knowledge sources and start a New Chat conversation
 */
export function openNewChatModal(tag: any, type: "Summary" | "AITag" = "AITag") {
    confirmSwitchChatHistory(() => {
        const store = StoreService.getInstance();
        const container = document.getElementById('confirmation-popup');
        if (!container) return;

        const sourceList = type === 'Summary'
            ? (store.sourceSummaryList && store.sourceSummaryList.length > 0 ? store.sourceSummaryList : store.sourceList)
            : (store.sourceList && store.sourceList.length > 0 ? store.sourceList : store.sourceSummaryList);

        const currentSessionsCount = tag.ChatSessions ? tag.ChatSessions.length : 0;
        const defaultTitle = `Chat ${currentSessionsCount + 1}`;

        container.innerHTML = NewChatSourceModalPopup(sourceList || [], defaultTitle, type);

        setTimeout(() => {
            const searchInput = document.getElementById('new-chat-source-search') as HTMLInputElement;
            const selectAllChk = document.getElementById('new-chat-select-all') as HTMLInputElement;
            const groupCheckboxes = document.querySelectorAll('.group-source-checkbox');
            const singleCheckboxes = document.querySelectorAll('.single-source-checkbox');
            const badgeElem = document.getElementById('new-chat-selected-badge');
            const confirmBtn = document.getElementById('new-chat-confirm-btn') as HTMLButtonElement;
            const cancelBtn = document.getElementById('new-chat-cancel-btn');
            const closeXBtn = document.getElementById('new-chat-close-x');
            const errorElem = document.getElementById('new-chat-source-error');

            const closeModal = () => {
                if (container) container.innerHTML = '';
            };

            cancelBtn?.addEventListener('click', closeModal);
            closeXBtn?.addEventListener('click', closeModal);

            let selectedValues: string[] = [];
            let selectedVectorIds: string[] = [];

            // Pre-select all available sources by default
            singleCheckboxes.forEach((cb: Element) => {
                const input = cb as HTMLInputElement;
                input.checked = true;
                selectedValues.push(input.value);
                if (input.dataset.vectorId) selectedVectorIds.push(input.dataset.vectorId);
            });
            if (selectAllChk) selectAllChk.checked = true;
            groupCheckboxes.forEach((gcb: Element) => (gcb as HTMLInputElement).checked = true);

            const updateSelectedState = () => {
                selectedValues = [];
                selectedVectorIds = [];
                singleCheckboxes.forEach((cb: Element) => {
                    const input = cb as HTMLInputElement;
                    if (input.checked) {
                        selectedValues.push(input.value);
                        if (input.dataset.vectorId) selectedVectorIds.push(input.dataset.vectorId);
                    }
                });

                if (badgeElem) {
                    badgeElem.innerText = `${selectedValues.length} Selected`;
                }

                if (selectedValues.length === 0) {
                    if (errorElem) errorElem.classList.remove('d-none');
                    if (confirmBtn) {
                        confirmBtn.disabled = true;
                        confirmBtn.classList.add('opacity-50');
                    }
                } else {
                    if (errorElem) errorElem.classList.add('d-none');
                    if (confirmBtn) {
                        confirmBtn.disabled = false;
                        confirmBtn.classList.remove('opacity-50');
                    }
                }
            };

            updateSelectedState();

            // Select All listener
            selectAllChk?.addEventListener('change', () => {
                const isChecked = selectAllChk.checked;
                singleCheckboxes.forEach((cb: Element) => {
                    (cb as HTMLInputElement).checked = isChecked;
                });
                groupCheckboxes.forEach((gcb: Element) => {
                    (gcb as HTMLInputElement).checked = isChecked;
                });
                updateSelectedState();
            });

            // Group checkbox listener
            groupCheckboxes.forEach((gcb: Element) => {
                gcb.addEventListener('change', () => {
                    const groupInput = gcb as HTMLInputElement;
                    const groupIdx = groupInput.dataset.groupIndex;
                    const childBoxes = document.querySelectorAll(`.single-source-checkbox[data-group-index="${groupIdx}"]`);
                    childBoxes.forEach((cb: Element) => ((cb as HTMLInputElement).checked = groupInput.checked));

                    if (selectAllChk) {
                        selectAllChk.checked = Array.from(singleCheckboxes).every((cb: Element) => (cb as HTMLInputElement).checked);
                    }
                    updateSelectedState();
                });
            });

            // Single checkbox listener
            singleCheckboxes.forEach((cb: Element) => {
                cb.addEventListener('change', () => {
                    const singleInput = cb as HTMLInputElement;
                    const groupIdx = singleInput.dataset.groupIndex;
                    const groupItems = document.querySelectorAll(`.single-source-checkbox[data-group-index="${groupIdx}"]`);
                    const groupChk = document.getElementById(`group-chk-${groupIdx}`) as HTMLInputElement;
                    if (groupChk) {
                        groupChk.checked = Array.from(groupItems).every((item: Element) => (item as HTMLInputElement).checked);
                    }
                    if (selectAllChk) {
                        selectAllChk.checked = Array.from(singleCheckboxes).every((item: Element) => (item as HTMLInputElement).checked);
                    }
                    updateSelectedState();
                });
            });

            // Search live filter
            searchInput?.addEventListener('input', () => {
                const term = searchInput.value.trim().toLowerCase();
                const sourceRows = document.querySelectorAll('.source-item-row');
                sourceRows.forEach((row: Element) => {
                    const rowElem = row as HTMLElement;
                    const text = rowElem.dataset.searchText || '';
                    if (term === '' || text.includes(term)) {
                        rowElem.style.display = 'block';
                    } else {
                        rowElem.style.display = 'none';
                    }
                });
            });

            // Confirm Start Chat
            confirmBtn?.addEventListener('click', async () => {
                if (selectedValues.length === 0) {
                    if (errorElem) errorElem.classList.remove('d-none');
                    return;
                }

                AIService.createNewChatSession(tag, defaultTitle, selectedValues, selectedVectorIds);
                closeModal();

                toaster(`Started new conversation`, 'success');

                const appBody = document.getElementById('app-body');
                if (appBody) {
                    appBody.innerHTML = await generateCheckboxHistory(tag, type);
                }
            });
        }, 0);
    });
}


