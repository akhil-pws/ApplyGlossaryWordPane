import { getPromptTemplateById, updateGroupKey, updateAiHistory, updatePromptTemplate } from "./draft.api";
import { chatfooter, copyText, generateChatHistoryHtml, insertLineWithHeadingStyle, removeQuotes, switchToAddTag, updateEditorFinalTable, colorTable, svgBase64ToPngBase64, resolveWordTableStyle, renderSelectedTags, parseHtmlTableToGrid, transposeGrid } from "./draft-functions";
import { addGenAITags, applyTagFn, createMultiSelectDropdown, customizeTable, getReport, mentionDropdownFn, saveAppState } from "../taskpane";
import { StoreService } from "../services/store.service";
import { AIService } from "../services/ai.service";
import { Confirmationpopup, DataModalPopup, toaster } from "../components/bodyelements";
import { summaryService } from "../services/summary.service";
import { updateSummaryHistory, updateSummaryTagPrompt } from "../summary/summary.api";

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
                        <a class="dropdown-item" href="#" id="sync-btn-tag">
                            <i class="fa fa-refresh me-2" aria-hidden="true"></i> Refresh Properties
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
            m.Name.toLowerCase().includes(searchTerm)
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

                listItem.innerHTML = `${icon} ${mention.Name}`;

                listItem.onclick = () => {
                    if (isAISection) {
                        const store = StoreService.getInstance();
                        store.currentChatTagId = mention.ID || mention.GroupKeyID;
                        saveAppState();

                        const appBody = document.getElementById('app-body');
                        appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';
                        generateCheckboxHistory(mention, "AITag")
                            .catch(() => appBody.innerHTML = '<div class="text-danger p-2">Error loading data</div>')
                            .then(html => {
                                appBody.innerHTML = html;
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

    document.getElementById('sync-btn-tag').addEventListener('click', async () => {
        if (!store.isPendingResponse) {
            loadSyncScreen();
        }
    });

    // Update sync button state based on store
    updateSyncButtonState();

    // Reopen last active AI Tag if applicable
    if (store.currentChatTagId !== -1) {
        const activeTag = availableKeys.find(k => (k.ID || k.GroupKeyID) === store.currentChatTagId);
        if (activeTag && activeTag.AIFlag === 1) {
            const appBody = document.getElementById('app-body')!;
            appBody.innerHTML = '<div class="text-muted p-2">Loading...</div>';
            generateCheckboxHistory(activeTag, "AITag")
                .catch(() => appBody.innerHTML = '<div class="text-danger p-2">Error loading data</div>')
                .then(html => {
                    if (html) {
                        appBody.innerHTML = html;
                    }
                });
        }
    }
}




export async function insertContentAtRange(context: Word.RequestContext, range: Word.Range, word: any, type: any, updateSelection: boolean = true) {

    // 1. Insert Anchor
    const anchorChar = range.insertText("\u200B", Word.InsertLocation.replace);
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

    let newSelection = range;

    if (type === 'TABLE') {
        const parser = new DOMParser();
        const doc = parser.parseFromString(word.Response, 'text/html');
        const bodyNodes = Array.from(doc.body.childNodes);

        await context.sync();

        for (const node of bodyNodes) {
            if (node.nodeType === Node.TEXT_NODE) {
                let textContent = node.textContent?.trim();
                if (textContent) {
                    textContent = textContent.replace(/\n- /g, "\n• ");

                    textContent.split('\n').forEach(line => {
                        if (line.trim()) {
                            const p = cursor.insertParagraph(line, Word.InsertLocation.after);
                            insertLineWithHeadingStyle(p, line);
                            include(p.getRange());
                            cursor = p.getRange();
                        }
                    });
                }
            } else if (node.nodeType === Node.ELEMENT_NODE) {
                const element = node as HTMLElement;

                if (element.tagName.toLowerCase() === 'table') {
                    const rows = Array.from(element.querySelectorAll('tr'));

                    if (rows.length === 0) {
                        const p = cursor.insertParagraph("[Empty Table]", Word.InsertLocation.after);
                        include(p.getRange());
                        cursor = p.getRange();
                        continue;
                    }

                    let grid = parseHtmlTableToGrid(rows);
                    const store = StoreService.getInstance();
                    const base = store.tableStyle.split(" - ")[0].trim();

                    if (base === 'Table Grid 2') {
                        store.isReversed = true;
                    } else {
                        store.isReversed = false;
                    }

                    if (store.isReversed) {
                        grid = transposeGrid(grid);
                    }

                    const numRows = grid.length;
                    const numCols = grid[0]?.length || 0;

                    const paragraph = cursor.insertParagraph("", Word.InsertLocation.after);
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
                                    table.getCell(rowIndex, cellIndex).value = cellValue;
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
                                        topCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
                                    } catch (e) { }
                                }
                            });
                        }
                    } else {
                        // Manual population for transposed table
                        grid.forEach((row, rowIndex) => {
                            row.forEach((cellValue, cellIndex) => {
                                table.getCell(rowIndex, cellIndex).value = cellValue;
                            });
                        });
                    }

                    // Styling logic (always call if Customize, now passing isReversed)
                    if (store.colorPallete.Customize) {
                        await colorTable(table, rows, context, store.isReversed);
                    }

                    include(table.getRange());
                    cursor = table.getRange();
                    newSelection = table.getCell(0, 0).body.getRange(); // Set the cursor to the start of the table
                } else {
                    let elementText = element.innerText.trim();
                    if (elementText) {
                        elementText = elementText.replace(/\n- /g, "\n• ");

                        elementText.split('\n').forEach(line => {
                            if (line.trim()) {
                                const p = cursor.insertParagraph(line, Word.InsertLocation.after);
                                insertLineWithHeadingStyle(p, line);
                                include(p.getRange());
                                cursor = p.getRange();
                            }
                        });
                    }
                    // newSelection = selection; // If it's not a table, just use the existing selection.
                }
            }
        }
    }
    else if (type === "IMAGE") {
        let base64Image: string = word.Response;

        if (base64Image.startsWith("data:image/svg+xml")) {
            // Convert SVG → PNG
            base64Image = await svgBase64ToPngBase64(base64Image);
        } else if (base64Image.startsWith("data:image")) {
            base64Image = base64Image.split(",")[1]; // strip prefix
        }

        // Insert at cursor
        const pic = cursor.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.replace);
        newSelection = pic.getRange();
    } else {
        // TEXT / Other
        if (word.Response === '' || word.IsApplied) {
            const p = cursor.insertParagraph(`#${word.Name}#`, Word.InsertLocation.after);
            include(p.getRange());
            cursor = p.getRange();
        } else {
            let content = removeQuotes(word.Response);
            let lines = content.split(/\r?\n/); // Handle both \r\n and \n
            lines.forEach(line => {
                const p = cursor.insertParagraph(line, Word.InsertLocation.after);
                include(p.getRange());
                cursor = p.getRange();
            });
        }
        newSelection = cursor; // After inserting the text, set selection to it.
    }

    await context.sync();

    // 3. Create Bookmark (if NOT Image)
    if (type !== 'IMAGE' && bookmarkStart && bookmarkEnd) {
        const bookmarkName = word.AIFlag === 1
            ? `ID${word.GroupKeyID}_Split_${getDateTimeStamp()}`
            : `PT${word.GroupKeyID}_Split_${getDateTimeStamp()}`;
        bookmarkStart.expandTo(bookmarkEnd).insertBookmark(bookmarkName);
    }

    // 4. Cleanup Anchor
    anchorChar.delete();
    await context.sync();

    if (updateSelection) {
        // Move the cursor to the next line after content insertion
        const nextLineParagraph = cursor.insertParagraph("", Word.InsertLocation.after);
        nextLineParagraph.select();
        await context.sync();
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

            await insertContentAtRange(context, selection, word, type, true);

        } catch (error) {
            console.error('Detailed error:', error);
        }
    });
}

export async function syncBookmarks() {
    return Word.run(async (context) => {
        try {
            const bookmarks = context.document.bookmarks;
            bookmarks.load("items/name");
            await context.sync();

            const store = StoreService.getInstance();
            const availableKeys = store.availableKeys;

            for (let i = 0; i < bookmarks.items.length; i++) {
                const bookmark = bookmarks.items[i];
                const name = bookmark.name;

                // Only process PT bookmarks
                if (!name.startsWith("PT")) continue;

                const parts = name.split("_");
                if (parts.length < 3) continue;

                const idPart = parts[0].substring(2);
                const groupKeyId = parseInt(idPart);

                if (isNaN(groupKeyId)) continue;

                const word = availableKeys.find(
                    (k: any) => k.GroupKeyID === groupKeyId
                );
                if (!word) continue;

                // Get bookmark range
                const range = context.document.getBookmarkRange(name);

                // 🔥 Replace content
                range.insertText(
                    word.Response ?? "",
                    Word.InsertLocation.replace
                );

                // 🔥 Recreate bookmark using range method
                const newBookmarkName = `PT${groupKeyId}_Split_${Date.now()}`;
                range.insertBookmark(newBookmarkName);
            }

            await context.sync();

            toaster("Properties refreshed!", "success");

            // Disable sync button after successful sync
            store.isSyncEnabled = false;
            updateSyncButtonState();

        } catch (err) {
            console.error("Sync error:", err);
            toaster("Refresh failed", "error");
        }
    });
}

export async function checkBookmarksForSync(): Promise<boolean> {
    return Word.run(async (context) => {
        try {
            const bookmarks = context.document.bookmarks;
            bookmarks.load("items/name");
            await context.sync();
            const store = StoreService.getInstance();
            const availableKeys = store.availableKeys;

            let hasChanges = false;

            for (let i = 0; i < bookmarks.items.length; i++) {
                const bookmark = bookmarks.items[i];
                const name = bookmark.name;

                // Only process PT bookmarks
                if (!name.startsWith("PT")) continue;

                const parts = name.split("_");
                if (parts.length < 3) continue;

                const idPart = parts[0].substring(2);
                const groupKeyId = parseInt(idPart);

                if (isNaN(groupKeyId)) continue;

                // Find matching tag by GroupKeyID
                const word = availableKeys.find(
                    (k: any) => k.GroupKeyID === groupKeyId
                );
                if (!word) continue;

                // Get bookmark content
                const range = context.document.getBookmarkRange(name);
                range.load("text");
                await context.sync();

                const bookmarkContent = range.text.trim();
                const tagResponse = (word.Response ?? "").trim();

                // If content is different, we have changes
                if (bookmarkContent !== tagResponse) {
                    hasChanges = true;
                    break; // No need to check further
                }
            }

            return hasChanges;

        } catch (err) {
            console.error("Error checking bookmarks for sync:", err);
            return false;
        }
    });
}

interface SyncItem {
    propertyName: string;
    updatedValue: string;
    existingValue: string;
    location: string;
    bookmarkName: string;
    tagId: number;
}

export async function getDesyncedBookmarks(): Promise<SyncItem[]> {
    return Word.run(async (context) => {
        const syncItems: SyncItem[] = [];
        try {
            const bookmarks = context.document.bookmarks;
            bookmarks.load("items/name");
            await context.sync();

            const store = StoreService.getInstance();
            const availableKeys = store.availableKeys;

            for (let i = 0; i < bookmarks.items.length; i++) {
                const bookmark = bookmarks.items[i];
                const name = bookmark.name;

                if (!name.startsWith("PT")) continue;

                const parts = name.split("_");
                if (parts.length < 3) continue;

                const idPart = parts[0].substring(2);
                const groupKeyId = parseInt(idPart);

                if (isNaN(groupKeyId)) continue;

                const word = availableKeys.find((k: any) => k.GroupKeyID === groupKeyId);
                if (!word) continue;

                const range = context.document.getBookmarkRange(name);
                range.load("text");
                const paragraph = range.paragraphs.getFirstOrNullObject();
                paragraph.load("text");

                await context.sync();

                const bookmarkContent = range.text.trim();
                const tagResponse = (word.Response ?? "").trim();

                if (bookmarkContent !== tagResponse) {
                    // Extract a snippet from the paragraph text
                    let pText = paragraph.isNullObject ? "" : paragraph.text;
                    if (pText.length > 150) pText = pText.substring(0, 150) + "...";

                    syncItems.push({
                        propertyName: word.Name,
                        updatedValue: tagResponse,
                        existingValue: bookmarkContent,
                        location: pText || "Unknown",
                        bookmarkName: name,
                        tagId: groupKeyId
                    });
                }
            }
        } catch (err) {
            console.error("Error fetching desynced bookmarks:", err);
            toaster("Error fetching sync details.", "error");
        }
        return syncItems;
    });
}

export async function syncSingleBookmark(bookmarkName: string, groupKeyId: number): Promise<boolean> {
    return Word.run(async (context) => {
        try {
            const store = StoreService.getInstance();
            const word = store.availableKeys.find((k: any) => k.GroupKeyID === groupKeyId);
            if (!word) return false;

            const range = context.document.getBookmarkRange(bookmarkName);
            range.insertText(word.Response ?? "", Word.InsertLocation.replace);

            const newBookmarkName = `PT${groupKeyId}_Split_${Date.now()}`;
            range.insertBookmark(newBookmarkName);

            await context.sync();
            toaster(`${word.Name} updated!`, "success");
            return true;
        } catch (err) {
            console.error("Error syncing single bookmark:", err);
            toaster("Failed to update property.", "error");
            return false;
        }
    });
}

export async function selectBookmarkByName(bookmarkName: string) {
    return Word.run(async (context) => {
        try {
            const range = context.document.getBookmarkRangeOrNullObject(bookmarkName);
            range.load('isNullObject');
            await context.sync();

            if (!range.isNullObject) {
                range.select();
                await context.sync();
            } else {
                toaster("Bookmark not found in document.", "error");
            }
        } catch (err) {
            console.error("Error selecting bookmark:", err);
            toaster("Failed to navigate to property.", "error");
        }
    });
}

export async function loadSyncScreen() {
    const store = StoreService.getInstance();
    const appBody = document.getElementById('app-body');
    const isDark = store.theme === 'Dark';
    const bgClass = isDark ? 'bg-dark text-light' : 'bg-white text-dark';
    const cardBgClass = isDark ? 'bg-secondary text-light' : 'bg-light text-dark';
    const btnClass = isDark ? 'btn-outline-light' : 'btn-outline-primary';

    appBody.innerHTML = `
        <div class="chat-header sticky-top ${bgClass} z-3">
            <div class="d-flex justify-content-between align-items-center px-2 pt-3">
                <div class="d-flex align-items-center ms-3 c-pointer" id="back-from-sync">
                    <i class="fa fa-arrow-left text-muted me-2"></i>
                    <span class="fw-bold">Refresh Properties</span>
                </div>
                <div class="d-flex justify-content-center align-items-center me-3">
                    <button class="btn btn-sm btn-primary" id="apply-all-sync">Apply All</button>
                </div>
            </div>
            <hr class="mt-2 mb-1 mx-3">
        </div>
        <div class="container pt-3 pb-5">
            <div class="card ${cardBgClass} mx-2" id="sync-card-container">
                <div class="card-header fw-semibold">Properties to Update</div>
                <div class="card-body p-2" id="sync-items-container">
                    <div class="text-center text-muted"><i class="fa fa-spinner fa-spin"></i> Scanning document...</div>
                </div>
                <div class="card-footer d-flex justify-content-between align-items-center flex-wrap gap-2 d-none" id="sync-pagination-footer">
                </div>
            </div>
        </div>
    `;

    document.getElementById('back-from-sync')?.addEventListener('click', () => {
        loadHomepage(store.availableKeys);
    });

    document.getElementById('apply-all-sync')?.addEventListener('click', async () => {
        const btn = document.getElementById('apply-all-sync') as HTMLButtonElement;
        btn.disabled = true;
        btn.innerHTML = '<i class="fa fa-spinner fa-spin"></i> Applying...';
        await syncBookmarks();
        loadHomepage(store.availableKeys);
    });

    // Fetch desynced items
    const desyncedItems = await getDesyncedBookmarks();
    const container = document.getElementById('sync-items-container');
    const footer = document.getElementById('sync-pagination-footer');

    if (desyncedItems.length === 0) {
        if (container) {
            container.innerHTML = `
                <div class="text-center mt-5 mb-5">
                    <i class="fa-solid fa-circle-check text-success fa-3x mb-3"></i>
                    <p class="text-muted">No desynchronized properties detected.</p>
                </div>
            `;
        }
        const applyAllBtn = document.getElementById('apply-all-sync') as HTMLButtonElement;
        if (applyAllBtn) applyAllBtn.classList.add('d-none');
        if (footer) footer.classList.add('d-none');

        return;
    }

    // Group items by property name
    const groupedItems: { [key: string]: { items: SyncItem[], currentIndex: number } } = {};
    desyncedItems.forEach(item => {
        if (!groupedItems[item.propertyName]) {
            groupedItems[item.propertyName] = { items: [], currentIndex: 0 };
        }
        groupedItems[item.propertyName].items.push(item);
    });

    let currentPage = 1;
    const pageSize = 2;

    function getTotalPages() {
        return Math.max(1, Math.ceil(Object.keys(groupedItems).length / pageSize));
    }

    const renderPage = () => {
        if (!container || !footer) return;
        container.innerHTML = '';
        footer.innerHTML = '';
        footer.classList.remove('d-none');

        const propNames = Object.keys(groupedItems);
        if (propNames.length === 0) {
            store.isSyncEnabled = false;
            updateSyncButtonState();
            loadHomepage(store.availableKeys);
            return;
        }

        const totalPages = getTotalPages();
        if (currentPage > totalPages) currentPage = Math.max(1, totalPages);

        const start = (currentPage - 1) * pageSize;
        const end = Math.min(start + pageSize, propNames.length);
        const currentProps = propNames.slice(start, end);

        currentProps.forEach(propName => {
            const group = groupedItems[propName];
            const card = document.createElement('div');
            card.className = `card mb-3 ${cardBgClass}`;
            card.id = `sync-card-${propName.replace(/\s+/g, '-')}`;

            const updateCardUI = () => {
                const currentItem = group.items[group.currentIndex];
                card.innerHTML = `
                    <div class="card-body p-2">
                        <div class="d-flex justify-content-between align-items-center mb-1">
                            <span class="fw-bold text-truncate">${propName}</span>
                            <div class="d-flex align-items-center gap-2">
                                <span class="badge bg-primary rounded-pill small" style="font-size: 0.7rem;">
                                    ${group.currentIndex + 1} / ${group.items.length}
                                </span>
                                <div class="btn-group btn-group-sm">
                                    <button class="btn btn-sm ${btnClass} sync-prev-btn">
                                        <i class="fa-solid fa-chevron-up"></i>
                                    </button>
                                    <button class="btn btn-sm ${btnClass} sync-next-btn">
                                        <i class="fa-solid fa-chevron-down"></i>
                                    </button>
                                </div>
                            </div>
                        </div>
                        <div class="text-muted small mb-1 border-bottom pb-1">
                            <div class="text-success"><strong>Updated Value:</strong> <span class="text-break">${currentItem.updatedValue}</span></div>
                            <div class="text-danger mt-1"><strong>Existing Value:</strong> <span class="text-break">${currentItem.existingValue}</span></div>
                        </div>
                        <div class="d-flex justify-content-end mt-2">
                            <button class="btn btn-sm ${btnClass} sync-single-btn">
                                <i class="fa fa-refresh"></i> Update
                            </button>
                        </div>
                    </div>
                `;

                card.querySelector('.sync-prev-btn')?.addEventListener('click', (e) => {
                    e.stopPropagation();
                    group.currentIndex = (group.currentIndex - 1 + group.items.length) % group.items.length;
                    updateCardUI();
                    selectBookmarkByName(group.items[group.currentIndex].bookmarkName);
                });

                card.querySelector('.sync-next-btn')?.addEventListener('click', (e) => {
                    e.stopPropagation();
                    group.currentIndex = (group.currentIndex + 1) % group.items.length;
                    updateCardUI();
                    selectBookmarkByName(group.items[group.currentIndex].bookmarkName);
                });

                card.querySelector('.sync-single-btn')?.addEventListener('click', async (e) => {
                    const button = e.currentTarget as HTMLButtonElement;
                    button.disabled = true;
                    button.innerHTML = '<i class="fa fa-spinner fa-spin"></i>';

                    const itemToSync = group.items[group.currentIndex];
                    const success = await syncSingleBookmark(itemToSync.bookmarkName, itemToSync.tagId);

                    if (success) {
                        group.items.splice(group.currentIndex, 1);
                        if (group.items.length === 0) {
                            delete groupedItems[propName];
                            renderPage();
                        } else {
                            if (group.currentIndex >= group.items.length) {
                                group.currentIndex = group.items.length - 1;
                            }
                            updateCardUI();
                        }
                    } else {
                        button.disabled = false;
                        button.innerHTML = '<i class="fa fa-refresh"></i> Update';
                    }
                });
            };

            updateCardUI();
            container.appendChild(card);
        });

        // Aligned Pagination Controls (Summary Tag Style)
        if (totalPages > 1 || propNames.length > 0) {
            const startItem = propNames.length === 0 ? 0 : (currentPage - 1) * pageSize + 1;
            const endItem = Math.min(currentPage * pageSize, propNames.length);

            footer.innerHTML = `
                <div class="btn-group btn-group-sm" role="group" aria-label="pagination">
                    <button class="btn btn-outline-secondary" id="sync-page-first" title="First" ${currentPage === 1 ? 'disabled' : ''}>
                        <i class="fa-solid fa-backward-fast"></i>
                    </button>
                    <button class="btn btn-outline-secondary" id="sync-page-prev" title="Previous" ${currentPage === 1 ? 'disabled' : ''}>
                        <i class="fa-solid fa-backward-step"></i>
                    </button>

                    <div class="btn-group btn-group-sm" id="sync-page-buttons"></div>

                    <button class="btn btn-outline-secondary" id="sync-page-next" title="Next" ${currentPage === totalPages ? 'disabled' : ''}>
                        <i class="fa-solid fa-forward-step"></i>
                    </button>
                    <button class="btn btn-outline-secondary" id="sync-page-last" title="Last" ${currentPage === totalPages ? 'disabled' : ''}>
                        <i class="fa-solid fa-forward-fast"></i>
                    </button>
                </div>
                <div class="text-muted small">${startItem} - ${endItem} of ${propNames.length} items</div>
            `;

            const pageButtons = footer.querySelector('#sync-page-buttons') as HTMLElement;
            const maxButtons = 5;
            let pStart = Math.max(1, currentPage - Math.floor(maxButtons / 2));
            let pEnd = Math.min(totalPages, pStart + maxButtons - 1);
            pStart = Math.max(1, pEnd - maxButtons + 1);

            for (let p = pStart; p <= pEnd; p++) {
                const b = document.createElement('button');
                b.className = `btn ${p === currentPage ? 'btn-primary text-white' : 'btn-outline-secondary'}`;
                b.textContent = String(p);
                b.onclick = () => {
                    currentPage = p;
                    renderPage();
                };
                pageButtons.appendChild(b);
            }

            footer.querySelector('#sync-page-first')?.addEventListener('click', () => { currentPage = 1; renderPage(); });
            footer.querySelector('#sync-page-prev')?.addEventListener('click', () => { currentPage = Math.max(1, currentPage - 1); renderPage(); });
            footer.querySelector('#sync-page-next')?.addEventListener('click', () => { currentPage = Math.min(totalPages, currentPage + 1); renderPage(); });
            footer.querySelector('#sync-page-last')?.addEventListener('click', () => { currentPage = totalPages; renderPage(); });
        } else {
            footer.classList.add('d-none');
        }
    };

    renderPage();
}

export function updateSyncButtonState() {
    const syncBtn = document.getElementById('sync-btn-tag');
    if (syncBtn) {
        syncBtn.classList.remove('disabled');
    }
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
    const history = tag.FilteredReportHeadAIHistoryList || [];
    if (history.length === 0) {
        return '<div>No AI history available.</div>';
    }

    // Default to the first (latest) item if none is explicitly selected
    const chat = history.find((item: any) => item.Selected) || history[0];

    const finalResponse = chat.FormattedResponse
        ? '\n' + updateEditorFinalTable(chat.FormattedResponse)
        : chat.Response;

    tag.ComponentKeyDataType = chat.FormattedResponse ? 'TABLE' : 'TEXT';
    // tag.UserValue = finalResponse;
    tag.Response = finalResponse;
    tag.text = finalResponse;

    // Check current theme
    const store = StoreService.getInstance();
    const isDark = store.theme === 'Dark';
    const closeBtnClass = isDark
        ? 'fa-solid fa-circle-xmark bg-dark text-light'
        : 'fa-solid fa-circle-xmark bg-light text-dark';

    const headerBgClass = isDark ? 'bg-dark text-light' : 'bg-white text-dark';
    const Name = type === 'Summary' ? tag.Name : tag.Name;
    const closeBar = `
        <div class="chat-header sticky-top ${headerBgClass} z-3">
            <div class="d-flex justify-content-between align-items-center px-2 pt-3">
                <div class="d-flex align-items-center ms-3">
                    <i class="fa fa-microchip-ai text-muted me-2"></i>
                    <span class="fw-bold">${Name}</span>
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
            <label class='form-label'><span class="text-danger">* </span> Prompt Builder Template</label>
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
        const id = item.ID || item.id;
        option.value = id ? id.toString() : '';
        option.textContent = item.Name;
        templateSelect.appendChild(option);
    });

    templateSelect.addEventListener('change', async () => {
        const templateId = templateSelect.value;
        const jwt = sessionStorage.getItem('token') || '';

        const data = await getPromptTemplateById(templateId, jwt);
        if (data.Status && data.Data) {
            fieldsList = data.Data;
            preview = promptBuilderList.find((item) => (item.ID || item.id || '').toString() === templateId).Template;

            templateText = promptBuilderList.find((item) => (item.ID || item.id || '').toString() === templateId).Template;
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
                const doc = parser.parseFromString(tag.Response, "text/html");
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

                            let grid = parseHtmlTableToGrid(rows);
                            const store = StoreService.getInstance();
                            const base = store.tableStyle.split(" - ")[0].trim();

                            if (base === 'Table Grid 2') {
                                store.isReversed = true;
                            } else {
                                store.isReversed = false;
                            }

                            if (store.isReversed) {
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
                                            table.getCell(rowIndex, cellIndex).value = cellValue;
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
                                                topCell.body.paragraphs.getFirst().alignment = Word.Alignment.centered;
                                            } catch (e) { }
                                        }
                                    });
                                }
                            } else {
                                // Manual population for transposed table
                                grid.forEach((rowGrid, rowIndex) => {
                                    rowGrid.forEach((cellValue, cellIndex) => {
                                        table.getCell(rowIndex, cellIndex).value = cellValue;
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
                const txt = tag.Response.replace(/\n- /g, "\n• ").trim();

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
                const prefix = type === "Summary" ? "SM" : "ID";
                const bookmarkName =
                    `${prefix}${tag.GroupKeyID || tag.ReportHeadSummaryTagID}_Split_${getDateTimeStamp()}`;

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
                            Name: type === 'Summary' ? tag.Name : tag.Name,
                            // UserValue: chat.Response,
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
                            // tag.UserValue = finalResponse;
                            tag.Response = finalResponse;
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
                                    // currentTag.UserValue = finalResponse;
                                    currentTag.Response = finalResponse;
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
                                    // currentTag.UserValue = finalResponse;
                                    currentTag.Response = finalResponse;
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
            store.currentChatTagId = -1;
            saveAppState();
            if (store.mode === "Home") {
                loadHomepage(availableKeys)
            } else if (store.mode === "Summary") {
                const { loadSummarypage } = require("../summary/summary");
                loadSummarypage(availableKeys);
            }
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

