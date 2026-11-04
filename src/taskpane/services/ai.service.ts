import { AITag, ChatMessage, ChatSession } from "../models/tag.model";
import { StoreService } from "./store.service";
import { getAiHistory, addAiHistory, getAITagReprocessStatus } from "../draft/draft.api";
import { removeQuotes, updateEditorFinalTable, REPROCESS_MESSAGE, getSelectedResponseMode, refreshResponseModeBar } from "../draft/draft-functions";
import { generateCheckboxHistory } from "../draft/home";
import { addSummaryHistory } from "../summary/summary.api";

export class AIService {
    private static reprocessPollTimer: ReturnType<typeof setTimeout> | null = null;

    /**
     * Syncs IsReprocess on AI tags from the API and keeps polling every 10 sec while any tag is reprocessing.
     * Pass refreshOpenChat = false when the caller is about to render the chat itself.
     */
    static async syncTagReprocessStatus(refreshOpenChat: boolean = true): Promise<void> {
        const store = StoreService.getInstance();

        if (AIService.reprocessPollTimer) {
            clearTimeout(AIService.reprocessPollTimer);
            AIService.reprocessPollTimer = null;
        }
        if (!store.jwt || !store.documentID) return;

        const allTags = () => [...(store.aiTagList || []), ...(store.availableKeys || [])].filter((tag: any) => tag.AIFlag !== 0);
        const tagId = (tag: any) => String(tag.ID || tag.ReportHeadGroupKeyID);

        // A chat response re-renders the chat when it returns, so completion waits for the next poll
        if (!store.isPendingResponse) {
            try {
                const res = await getAITagReprocessStatus(store.documentID, store.jwt);
                const activeIds = new Set(
                    (res?.Data || [])
                        .filter((item: any) => item.IsReprocess)
                        .map((item: any) => String(item.ReportHeadGroupKeyID))
                );

                // Tags can exist in both lists, so completed tags are keyed by ID
                const completed = new Map<string, any>();
                allTags().forEach((tag: any) => {
                    const isReprocess = activeIds.has(tagId(tag));
                    if (tag.IsReprocess && !isReprocess && !completed.has(tagId(tag))) {
                        completed.set(tagId(tag), tag);
                    }
                    tag.IsReprocess = isReprocess;
                });

                await Promise.all(Array.from(completed.values()).map(async (tag: any) => {
                    const activeChatHistoryId = tag.ChatSessions?.[tag.ActiveSessionIndex || 0]?.chatHistoryId;
                    await AIService.fetchAIHistory(tag);
                    const activeIdx = (tag.ChatSessions || []).findIndex((s: any) => String(s.chatHistoryId) === String(activeChatHistoryId));
                    if (activeIdx !== -1) tag.ActiveSessionIndex = activeIdx;
                }));

                if (refreshOpenChat) {
                    await AIService.updateOpenChatReprocessState(completed);
                }
            } catch (error) {
                console.error('Error fetching AI tag reprocess status:', error);
            }
        }

        if (allTags().some((tag: any) => tag.IsReprocess)) {
            AIService.reprocessPollTimer = setTimeout(() => AIService.syncTagReprocessStatus(), 10000);
        }
    }

    private static async updateOpenChatReprocessState(completed: Map<string, any>): Promise<void> {
        const store = StoreService.getInstance();
        const isAIChatOpen = () => document.getElementById('chatFooter')?.dataset.chatType === 'AITag';
        if (!isAIChatOpen()) return;

        const openTagId = String(store.currentChatTagId);
        const completedTag = completed.get(openTagId);

        if (completedTag) {
            const draft = (document.getElementById('chatInput') as HTMLTextAreaElement)?.value || '';
            const html = await generateCheckboxHistory(completedTag, 'AITag');
            const appBody = document.getElementById('app-body');
            if (appBody && isAIChatOpen() && String(store.currentChatTagId) === openTagId) {
                appBody.innerHTML = html;
                const chatInput = document.getElementById('chatInput') as HTMLTextAreaElement;
                if (chatInput) chatInput.value = draft;
            }
            return;
        }

        const openTag = (store.aiTagList || []).find((tag: any) => String(tag.ID || tag.ReportHeadGroupKeyID) === openTagId);
        if (openTag?.IsReprocess) {
            AIService.showChatReprocessState();
        }
    }

    // Applies the same locked state chatfooter renders when a refresh starts with the chat already open
    private static showChatReprocessState(): void {
        ['chatInput', 'promptBuilderButton', 'sendPromptButton', 'responseModeStandard', 'responseModeAgent'].forEach((id) => {
            const control = document.getElementById(id);
            control?.setAttribute('disabled', 'true');
            // The approval lock restores this on unapprove, so it must not re-enable the control mid-refresh
            if (control?.hasAttribute('data-approval-locked')) {
                control.setAttribute('data-approval-was-disabled', '');
            }
        });
        // While approved the lock message stays visible, and the approval lock restores these on unapprove
        const chatInput = document.getElementById('chatInput');
        const isApprovalLocked = chatInput?.hasAttribute('data-approval-locked');
        chatInput?.setAttribute(isApprovalLocked ? 'data-approval-prev-placeholder' : 'placeholder', REPROCESS_MESSAGE);
        chatInput?.setAttribute(isApprovalLocked ? 'data-approval-prev-title' : 'title', REPROCESS_MESSAGE);
        ['promptBuilderButton', 'sendPromptButton'].forEach((id) => {
            const tooltip = document.querySelector(`#${id} .tooltiptext`);
            if (tooltip) tooltip.textContent = REPROCESS_MESSAGE;
        });
        if (!document.querySelector('.horizontal-loader')) {
            document.getElementById('chatFooter')?.insertAdjacentHTML(
                'beforebegin',
                `<div class="horizontal-loader" title="${REPROCESS_MESSAGE}"></div>`
            );
        }
    }

    /**
     * Groups flat history list from API (with ChatHistoryID) into structured ChatSession objects.
     */
    static groupHistoryIntoSessions(rawHistory: any[], tag: any, store: any, type: "Summary" | "AITag" = "AITag"): ChatSession[] {
        if (!rawHistory || !Array.isArray(rawHistory) || rawHistory.length === 0) {
            return [];
        }

        const sourcePool = (type === 'Summary')
            ? (store.sourceSummaryList?.length ? store.sourceSummaryList : store.sourceList || [])
            : (store.sourceList?.length ? store.sourceList : store.sourceSummaryList || []);

        // Group messages by ChatHistoryID (fallback to 1 if not present)
        const sessionMap = new Map<number | string, ChatMessage[]>();

        rawHistory.forEach((item: any) => {
            const hId = item.ChatHistoryID !== undefined && item.ChatHistoryID !== null && item.ChatHistoryID !== ''
                ? item.ChatHistoryID
                : 1;

            if (!sessionMap.has(hId)) {
                sessionMap.set(hId, []);
            }
            const cleanItem: ChatMessage = {
                ...item,
                Response: removeQuotes(item.Response || '')
            };
            sessionMap.get(hId)!.push(cleanItem);
        });

        const sessions: ChatSession[] = [];

        sessionMap.forEach((messages, hId) => {
            // rawHistory from API has newest messages first.
            // Reverse messages so older messages appear at the top and the latest messages appear at the bottom of the chat window.
            const chronologicalMessages = [...messages].reverse();

            // Determine session title from the initial (oldest) prompt in this thread
            const firstMsg = chronologicalMessages.find(m => m.Prompt && m.Prompt.trim() !== '') || chronologicalMessages[0];
            let title = `Chat ${hId}`;
            if (firstMsg?.Prompt) {
                title = firstMsg.Prompt.trim().replace(/\s+/g, ' ');
            }

            // Resolve sources from the latest message's SourceValue / SourceVector (messages is newest first)
            let sessionSources: string[] = [];
            let sessionSourceValues: string[] = [];

            const latestMsg = messages[0];
            const rawSourceVal = latestMsg?.SourceValue || latestMsg?.SourceVector || tag?.SourceValue || tag?.SourceVector;
            if (rawSourceVal) {
                const sourceIds: string[] = Array.isArray(rawSourceVal)
                    ? rawSourceVal.map(String)
                    : String(rawSourceVal).split(',').map(s => s.trim()).filter(Boolean);

                const matched = sourcePool.filter((s: any) =>
                    sourceIds.includes(String(s.VectorID)) || sourceIds.includes(String(s.SourceValue))
                );

                if (matched.length > 0) {
                    sessionSources = matched.map((s: any) => s.SourceName || s.FileName || s.SourceValue);
                    sessionSourceValues = matched.map((s: any) => String(s.VectorID || s.SourceValue));
                } else if (sourceIds.length > 0) {
                    sessionSourceValues = sourceIds;
                    sessionSources = sourceIds;
                }
            }

            if (sessionSources.length === 0 && tag?.Sources?.length) {
                sessionSources = [...tag.Sources];
                sessionSourceValues = tag?.TempSourceValue ? [...tag.TempSourceValue] : [];
            }

            const createdAt = firstMsg?.CreatedDate || new Date().toISOString();

            sessions.push({
                id: `session-${hId}`,
                chatHistoryId: hId,
                title,
                createdAt,
                sources: sessionSources,
                sourceValues: sessionSourceValues,
                history: chronologicalMessages
            });
        });

        // Sort sessions in descending order of ChatHistoryID (most recent thread at top)
        sessions.sort((a, b) => {
            const numA = Number(a.chatHistoryId) || 0;
            const numB = Number(b.chatHistoryId) || 0;
            if (numA !== numB) {
                return numB - numA;
            }
            return new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime();
        });

        return sessions;
    }

    /**
     * Initializes or fetches AI history for a tag from the real API and groups into multi-sessions.
     */
    static async fetchAIHistory(tag: any): Promise<any[]> {
        const store = StoreService.getInstance();
        try {
            const tagId = tag.ID || tag.ReportHeadGroupKeyID;
            const data = await getAiHistory(tagId, store.jwt);

            if (data.Status && data.Data && Array.isArray(data.Data) && data.Data.length > 0) {
                const sessions = AIService.groupHistoryIntoSessions(data.Data, tag, store, "AITag");
                tag.ChatSessions = sessions;

                if (tag.ActiveSessionIndex === undefined || tag.ActiveSessionIndex < 0 || tag.ActiveSessionIndex >= sessions.length) {
                    tag.ActiveSessionIndex = 0;
                }

                const activeSession = sessions[tag.ActiveSessionIndex];
                if (activeSession) {
                    tag.FilteredReportHeadAIHistoryList = activeSession.history;
                    tag.ReportHeadAIHistoryList = activeSession.history;
                    tag.Sources = [...activeSession.sources];
                    tag.SourceName = [...activeSession.sources];
                    tag.TempSourceValue = [...activeSession.sourceValues];
                    tag.SourceValueID = activeSession.sourceValues;

                    const selectedChat = activeSession.history.find((item: any) => item.Selected === 1);
                    if (selectedChat) {
                        // Enforce single selection across all sessions
                        sessions.forEach(s => {
                            s.history.forEach(m => {
                                m.Selected = (m === selectedChat) ? 1 : 0;
                            });
                        });

                        const finalResponse = selectedChat.FormattedResponse
                            ? '\n' + updateEditorFinalTable(selectedChat.FormattedResponse)
                            : selectedChat.Response;

                        tag.ComponentKeyDataType = selectedChat.FormattedResponse ? 'TABLE' : 'TEXT';
                        tag.UserValue = finalResponse;
                        tag.EditorValue = finalResponse;
                        tag.text = finalResponse;
                        tag.IsApplied = false;
                    } else {
                        tag.UserValue = '';
                        tag.EditorValue = '';
                        tag.text = '';
                        tag.IsApplied = true;
                    }

                    const chatId = selectedChat?.ChatHistoryID || selectedChat?.ID || selectedChat?.ReportHeadAIHistoryID || activeSession?.chatHistoryId || '';
                    tag.ChatHistoryID = chatId;

                    const tagId = tag.ID || tag.ReportHeadGroupKeyID;
                    store.aiTagList?.forEach((currentTag: any) => {
                        const currentId = currentTag.ID || currentTag.ReportHeadGroupKeyID;
                        if (currentTag === tag || currentId === tagId || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName) || (currentTag.GroupKey && currentTag.GroupKey === tag.GroupKey)) {
                            currentTag.ChatSessions = tag.ChatSessions;
                            currentTag.ActiveSessionIndex = tag.ActiveSessionIndex;
                            currentTag.FilteredReportHeadAIHistoryList = tag.FilteredReportHeadAIHistoryList;
                            currentTag.ReportHeadAIHistoryList = tag.ReportHeadAIHistoryList;
                            currentTag.ChatHistoryID = chatId;
                            currentTag.UserValue = tag.UserValue;
                            currentTag.EditorValue = tag.EditorValue;
                            currentTag.text = tag.text;
                            currentTag.IsApplied = tag.IsApplied;
                            currentTag.ComponentKeyDataType = tag.ComponentKeyDataType;
                        }
                    });

                    store.availableKeys?.forEach((currentTag: any) => {
                        const currentId = currentTag.ID || currentTag.ReportHeadGroupKeyID;
                        if (currentTag === tag || currentId === tagId || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName) || (currentTag.GroupKey && currentTag.GroupKey === tag.GroupKey)) {
                            currentTag.ChatSessions = tag.ChatSessions;
                            currentTag.ActiveSessionIndex = tag.ActiveSessionIndex;
                            currentTag.FilteredReportHeadAIHistoryList = tag.FilteredReportHeadAIHistoryList;
                            currentTag.ReportHeadAIHistoryList = tag.ReportHeadAIHistoryList;
                            currentTag.ChatHistoryID = chatId;
                            currentTag.UserValue = tag.UserValue;
                            currentTag.EditorValue = tag.EditorValue;
                            currentTag.text = tag.text;
                            currentTag.IsApplied = tag.IsApplied;
                            currentTag.ComponentKeyDataType = tag.ComponentKeyDataType;
                        }
                    });

                    return tag.FilteredReportHeadAIHistoryList;
                }
            } else {
                console.warn("No AI history available.");
                tag.ChatSessions = [];
                tag.ActiveSessionIndex = 0;
                tag.FilteredReportHeadAIHistoryList = [];
                tag.ReportHeadAIHistoryList = [];
                return [];
            }
            return [];
        } catch (error) {
            console.error('Error fetching AI history:', error);
            tag.ChatSessions = [];
            tag.ActiveSessionIndex = 0;
            tag.FilteredReportHeadAIHistoryList = [];
            tag.ReportHeadAIHistoryList = [];
            return [];
        }
    }

    /**
     * Creates a new chat session with user-selected sources and sets it as active (ChatHistoryID = 0)
     */
    static createNewChatSession(tag: any, title: string, selectedSources: string[], selectedSourceValues: string[]): ChatSession {
        if (!tag.ChatSessions) {
            tag.ChatSessions = [];
        }

        const sessionIndex = tag.ChatSessions.length + 1;
        const newSession: ChatSession = {
            id: `session-new-${Date.now()}`,
            chatHistoryId: 0,
            title: title && title.trim() ? title.trim() : `Chat ${sessionIndex}`,
            createdAt: new Date().toISOString(),
            sources: selectedSources && selectedSources.length > 0 ? [...selectedSources] : [],
            sourceValues: selectedSourceValues && selectedSourceValues.length > 0 ? [...selectedSourceValues] : [],
            history: []
        };

        // Add new session to top of list and make it active
        tag.ChatSessions.unshift(newSession);
        tag.ActiveSessionIndex = 0;
        tag.FilteredReportHeadAIHistoryList = [];
        tag.ReportHeadAIHistoryList = [];
        tag.Sources = [...newSession.sources];
        tag.SourceName = [...newSession.sources];
        tag.TempSourceValue = [...newSession.sourceValues];
        tag.SourceValueID = newSession.sourceValues;

        // Sync across store lists
        const store = StoreService.getInstance();
        const tagId = tag.ID || tag.ReportHeadGroupKeyID || tag.ReportHeadSummaryTagID;

        store.aiTagList?.forEach((currentTag: any) => {
            const currentId = currentTag.ID || currentTag.ReportHeadGroupKeyID;
            if (currentTag === tag || currentId === tagId || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName) || (currentTag.GroupKey && currentTag.GroupKey === tag.GroupKey)) {
                currentTag.ChatSessions = tag.ChatSessions;
                currentTag.ActiveSessionIndex = 0;
                currentTag.FilteredReportHeadAIHistoryList = [];
                currentTag.ReportHeadAIHistoryList = [];
                currentTag.Sources = [...newSession.sources];
                currentTag.SourceName = [...newSession.sources];
                currentTag.TempSourceValue = [...newSession.sourceValues];
                currentTag.SourceValueID = newSession.sourceValues;
            }
        });
        store.availableKeys?.forEach((currentTag: any) => {
            const currentId = currentTag.ID || currentTag.ReportHeadGroupKeyID;
            if (currentTag === tag || currentId === tagId || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName) || (currentTag.GroupKey && currentTag.GroupKey === tag.GroupKey)) {
                currentTag.ChatSessions = tag.ChatSessions;
                currentTag.ActiveSessionIndex = 0;
                currentTag.FilteredReportHeadAIHistoryList = [];
                currentTag.ReportHeadAIHistoryList = [];
                currentTag.Sources = [...newSession.sources];
                currentTag.SourceName = [...newSession.sources];
                currentTag.TempSourceValue = [...newSession.sourceValues];
                currentTag.SourceValueID = newSession.sourceValues;
            }
        });
        store.summaryTagList?.forEach((currentTag: any) => {
            const currentId = currentTag.ID || currentTag.ReportHeadSummaryTagID;
            if (currentTag === tag || currentId === tagId || currentTag.Name === tag.Name || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName)) {
                currentTag.ChatSessions = tag.ChatSessions;
                currentTag.ActiveSessionIndex = 0;
                currentTag.FilteredReportHeadAIHistoryList = [];
                currentTag.ReportHeadAIHistoryList = [];
                currentTag.Sources = [...newSession.sources];
                currentTag.SourceName = [...newSession.sources];
                currentTag.TempSourceValue = [...newSession.sourceValues];
                currentTag.SourceValueID = newSession.sourceValues;
            }
        });

        return newSession;
    }

    /**
     * Switches the active session to the chosen session index
     */
    static switchChatSession(tag: any, sessionIndex: number): void {
        if (!tag.ChatSessions || sessionIndex < 0 || sessionIndex >= tag.ChatSessions.length) {
            return;
        }

        tag.ActiveSessionIndex = sessionIndex;
        const activeSession = tag.ChatSessions[sessionIndex];

        tag.FilteredReportHeadAIHistoryList = activeSession.history || [];
        tag.ReportHeadAIHistoryList = activeSession.history || [];
        tag.Sources = [...activeSession.sources];
        tag.SourceName = [...activeSession.sources];
        tag.TempSourceValue = [...activeSession.sourceValues];
        tag.SourceValueID = activeSession.sourceValues;

        // Find if active session has a selected message, or if any other session has a selected message
        let selectedChat = activeSession.history?.find((item: any) => item.Selected === 1);
        let selectedSession = activeSession;
        if (!selectedChat && tag.ChatSessions && Array.isArray(tag.ChatSessions)) {
            const foundSession = tag.ChatSessions.find((s: any) => s.history?.some((m: any) => m.Selected === 1));
            if (foundSession) {
                selectedSession = foundSession;
                selectedChat = foundSession.history?.find((m: any) => m.Selected === 1);
            }
        }

        if (selectedChat) {
            const finalResponse = selectedChat.FormattedResponse
                ? '\n' + updateEditorFinalTable(selectedChat.FormattedResponse)
                : selectedChat.Response;

            tag.ComponentKeyDataType = selectedChat.FormattedResponse ? 'TABLE' : 'TEXT';
            tag.UserValue = finalResponse;
            tag.EditorValue = finalResponse;
            tag.text = finalResponse;
            tag.IsApplied = false;
            tag.ChatHistoryID = selectedChat.ChatHistoryID || selectedChat.ID || selectedChat.ReportHeadAIHistoryID || selectedSession?.chatHistoryId || '';
        } else {
            tag.UserValue = '';
            tag.EditorValue = '';
            tag.text = '';
            tag.IsApplied = true;
            tag.ChatHistoryID = '';
        }

        const store = StoreService.getInstance();
        const tagId = tag.ID || tag.ReportHeadGroupKeyID;
        store.aiTagList?.forEach((currentTag: any) => {
            const currentId = currentTag.ID || currentTag.ReportHeadGroupKeyID;
            if (currentTag === tag || currentId === tagId || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName) || (currentTag.GroupKey && currentTag.GroupKey === tag.GroupKey)) {
                currentTag.ChatSessions = tag.ChatSessions;
                currentTag.ActiveSessionIndex = tag.ActiveSessionIndex;
                currentTag.FilteredReportHeadAIHistoryList = tag.FilteredReportHeadAIHistoryList;
                currentTag.ReportHeadAIHistoryList = tag.ReportHeadAIHistoryList;
                currentTag.ChatHistoryID = tag.ChatHistoryID;
                currentTag.UserValue = tag.UserValue;
                currentTag.EditorValue = tag.EditorValue;
                currentTag.text = tag.text;
                currentTag.IsApplied = tag.IsApplied;
                currentTag.ComponentKeyDataType = tag.ComponentKeyDataType;
            }
        });
        store.availableKeys?.forEach((currentTag: any) => {
            const currentId = currentTag.ID || currentTag.ReportHeadGroupKeyID;
            if (currentTag === tag || currentId === tagId || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName) || (currentTag.GroupKey && currentTag.GroupKey === tag.GroupKey)) {
                currentTag.ChatSessions = tag.ChatSessions;
                currentTag.ActiveSessionIndex = tag.ActiveSessionIndex;
                currentTag.FilteredReportHeadAIHistoryList = tag.FilteredReportHeadAIHistoryList;
                currentTag.ReportHeadAIHistoryList = tag.ReportHeadAIHistoryList;
                currentTag.ChatHistoryID = tag.ChatHistoryID;
                currentTag.UserValue = tag.UserValue;
                currentTag.EditorValue = tag.EditorValue;
                currentTag.text = tag.text;
                currentTag.IsApplied = tag.IsApplied;
                currentTag.ComponentKeyDataType = tag.ComponentKeyDataType;
            }
        });
        store.summaryTagList?.forEach((currentTag: any) => {
            const currentId = currentTag.ID || currentTag.ReportHeadSummaryTagID;
            if (currentTag === tag || currentId === tagId || currentTag.Name === tag.Name || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName)) {
                currentTag.ChatSessions = tag.ChatSessions;
                currentTag.ActiveSessionIndex = tag.ActiveSessionIndex;
                currentTag.FilteredReportHeadAIHistoryList = tag.FilteredReportHeadAIHistoryList;
                currentTag.ReportHeadAIHistoryList = tag.ReportHeadAIHistoryList;
                currentTag.ChatHistoryID = tag.ChatHistoryID;
                currentTag.UserValue = tag.UserValue;
                currentTag.EditorValue = tag.EditorValue;
                currentTag.text = tag.text;
                currentTag.IsApplied = tag.IsApplied;
                currentTag.ComponentKeyDataType = tag.ComponentKeyDataType;
            }
        });
    }

    /**
     * Locks the response mode toggle and chat input while a response is generating.
     */
    static setChatInputLocked(tag: any, locked: boolean): void {
        const disable = locked || !!tag.IsReprocess;
        refreshResponseModeBar(tag, disable);

        const chatInput = document.getElementById('chatInput') as HTMLTextAreaElement | null;
        if (chatInput) {
            chatInput.disabled = disable;
            chatInput.placeholder = locked ? 'Generating response…' : (tag.IsReprocess ? REPROCESS_MESSAGE : '');
        }
        const insertButton = document.getElementById('insertTagButton') as HTMLButtonElement | null;
        if (insertButton) insertButton.disabled = locked;
        const promptBuilderButton = document.getElementById('promptBuilderButton') as HTMLButtonElement | null;
        if (promptBuilderButton) promptBuilderButton.disabled = disable;
    }

    /**
     * Sends prompt to the backend API. Passes chatHistoryID = 0 for new chat, or existing chatHistoryID for continuation.
     */
    static async sendPrompt(tag: any, prompt: string, type: "Summary" | "AITag" = "AITag"): Promise<void> {
        const store = StoreService.getInstance();

        if (prompt && prompt.trim() !== '' && !store.isTagUpdating) {
            store.isTagUpdating = true;
            store.isPendingResponse = true;

            const iconelement = document.getElementById(`sendPromptButton`);
            if (iconelement) iconelement.innerHTML = `<i class="fa fa-spinner fa-spin text-white"></i>`;
            AIService.setChatInputLocked(tag, true);

            const activeSession: ChatSession | null = tag.ChatSessions && tag.ActiveSessionIndex !== undefined
                ? tag.ChatSessions[tag.ActiveSessionIndex]
                : null;

            // chatHistoryId = 0 for brand new chat, or the existing integer ID (e.g. 1, 2)
            const targetChatHistoryId = activeSession ? (Number(activeSession.chatHistoryId) || 0) : 0;

            let payload: any;
            const documentInstruction = store.documentInstruction || store.dataList?.DocumentInstruction || store.dataList?.DocumentInstructions || '';

            if (type === 'Summary') {
                payload = {
                    ReportHeadID: store.dataList?.ID,
                    ReportHeadSummaryTagID: tag.ID || tag.ReportHeadSummaryTagID,
                    Prompt: prompt.trim(),
                    Response: "",
                    Selected: 1,
                    SourceVector: tag.TempSourceValue ? tag.TempSourceValue.join(",") : "",
                    Name: tag.Name,
                    DocumentInstruction: documentInstruction,
                    ChatHistoryID: targetChatHistoryId,
                    ResponseMode: getSelectedResponseMode(tag)
                };
            } else {
                const firstHistoryItem = tag.FilteredReportHeadAIHistoryList?.[0] || tag.ReportHeadAIHistoryList?.[0];
                payload = {
                    ReportHeadID: store.dataList?.ID || firstHistoryItem?.ReportHeadID || tag.ReportHeadID,
                    ReportHeadGroupKeyID: tag.ID || tag.ReportHeadGroupKeyID || firstHistoryItem?.ReportHeadGroupKeyID,
                    DocumentID: store.dataList?.NCTID,
                    DocumentType: store.dataList?.DocumentType,
                    TextSetting: store.dataList?.TextSetting,
                    DocumentTemplate: store.dataList?.ReportTemplate,
                    ThreadID: tag.ThreadID,
                    AssistantID: store.dataList?.AssistantID,
                    Container: store.dataList?.Container,
                    GroupName: tag.GroupName || store.GroupName || '',
                    GroupKey: tag.GroupKey || tag.DisplayName || tag.Name || tag.GroupName || store.GroupName || '',
                    Prompt: prompt.trim(),
                    PromptType: 1,
                    Response: '',
                    VectorID: store.dataList?.VectorID,
                    Selected: 0,
                    ID: 0,
                    ChatHistoryID: targetChatHistoryId,
                    SourceValue: tag.TempSourceValue ? tag.TempSourceValue : [],
                    DocumentInstruction: documentInstruction,
                    ResponseMode: getSelectedResponseMode(tag)
                };
            }

            try {
                const data = type === 'Summary'
                    ? await addSummaryHistory(payload, store.jwt)
                    : await addAiHistory(payload, store.jwt);

                if (data['Data'] && data['Data'] !== 'false') {
                    const rawData = Array.isArray(data['Data']) ? data['Data'] : [data['Data']];
                    const sessions = AIService.groupHistoryIntoSessions(rawData, tag, store, type);
                    tag.ChatSessions = sessions;

                    // Locate active session after refresh
                    if (targetChatHistoryId !== 0) {
                        const foundIdx = sessions.findIndex(s => Number(s.chatHistoryId) === targetChatHistoryId);
                        tag.ActiveSessionIndex = foundIdx !== -1 ? foundIdx : 0;
                    } else {
                        // New chat was assigned a new ChatHistoryID by backend; it will be the newest/first session
                        tag.ActiveSessionIndex = 0;
                    }

                    const currentActive = sessions[tag.ActiveSessionIndex] || sessions[0];
                    if (currentActive) {
                        tag.FilteredReportHeadAIHistoryList = currentActive.history;
                        tag.ReportHeadAIHistoryList = currentActive.history;
                        tag.Sources = [...currentActive.sources];
                        tag.SourceName = [...currentActive.sources];
                        tag.TempSourceValue = [...currentActive.sourceValues];
                        tag.SourceValueID = currentActive.sourceValues;

                        const chat = currentActive.history.find((item: any) => item.Selected === 1) || currentActive.history[currentActive.history.length - 1];
                        if (chat) {
                            chat.Selected = 1;
                            // Enforce single selection across all sessions
                            sessions.forEach(s => {
                                s.history?.forEach(m => {
                                    m.Selected = (m === chat) ? 1 : 0;
                                });
                            });

                            const isTable = chat.FormattedResponse && chat.FormattedResponse !== '';
                            const finalResponse = isTable
                                ? '\n' + updateEditorFinalTable(chat.FormattedResponse)
                                : chat.Response;

                            tag.ComponentKeyDataType = isTable ? 'TABLE' : 'TEXT';
                            tag.UserValue = finalResponse;
                            tag.EditorValue = finalResponse;
                            tag.text = finalResponse;
                            tag.IsApplied = false;

                            // Update lists in Store
                            if (type === 'Summary') {
                                store.summaryTagList?.forEach((currentTag: any) => {
                                    const currentId = currentTag.ID || currentTag.ReportHeadSummaryTagID;
                                    const tagId = tag.ID || tag.ReportHeadSummaryTagID;
                                    if (currentTag === tag || currentId === tagId || currentTag.Name === tag.Name || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName)) {
                                        AIService.updateTagWithChat(currentTag, chat, tag.IsApplied, tag);
                                    }
                                });
                            } else {
                                store.aiTagList?.forEach((currentTag: any) => {
                                    const currentId = currentTag.ID || currentTag.ReportHeadGroupKeyID;
                                    const tagId = tag.ID || tag.ReportHeadGroupKeyID;
                                    if (currentTag === tag || currentId === tagId || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName) || (currentTag.GroupKey && currentTag.GroupKey === tag.GroupKey)) {
                                        AIService.updateTagWithChat(currentTag, chat, tag.IsApplied, tag);
                                    }
                                });

                                store.availableKeys?.forEach((currentTag: any) => {
                                    const currentId = currentTag.ID || currentTag.ReportHeadGroupKeyID;
                                    const tagId = tag.ID || tag.ReportHeadGroupKeyID;
                                    if (currentTag === tag || currentId === tagId || (currentTag.DisplayName && currentTag.DisplayName === tag.DisplayName) || (currentTag.GroupKey && currentTag.GroupKey === tag.GroupKey)) {
                                        AIService.updateTagWithChat(currentTag, chat, tag.IsApplied, tag);
                                    }
                                });
                            }
                        }
                    }

                    const appbody = document.getElementById('app-body');
                    if (appbody) appbody.innerHTML = await generateCheckboxHistory(tag, type);
                }

                if (iconelement) iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`;
                const chatInput = document.getElementById(`chatInput`) as HTMLInputElement;
                if (chatInput) chatInput.value = '';

                store.isTagUpdating = false;
                store.isPendingResponse = false;
                AIService.setChatInputLocked(tag, false);

            } catch (error) {
                if (iconelement) iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`;
                store.isTagUpdating = false;
                store.isPendingResponse = false;
                AIService.setChatInputLocked(tag, false);
                console.error('Error sending AI prompt:', error);
            }
        } else {
            console.error('No empty prompt allowed or tag updating');
        }
    }

    private static updateTagWithChat(currentTag: any, chat: any, isApplied: any, tag?: any) {
        const isTable = chat.FormattedResponse && chat.FormattedResponse !== '';
        const finalResponse = isTable
            ? '\n' + updateEditorFinalTable(chat.FormattedResponse)
            : chat.Response;

        currentTag.ComponentKeyDataType = isTable ? 'TABLE' : 'TEXT';
        currentTag.UserValue = finalResponse;
        currentTag.EditorValue = finalResponse;
        currentTag.text = finalResponse;
        currentTag.IsApplied = isApplied;
        if (tag) {
            currentTag.ChatSessions = tag.ChatSessions;
            currentTag.ActiveSessionIndex = tag.ActiveSessionIndex;
            currentTag.FilteredReportHeadAIHistoryList = tag.FilteredReportHeadAIHistoryList;
            currentTag.ReportHeadAIHistoryList = tag.ReportHeadAIHistoryList;
        }
        currentTag.ChatHistoryID = chat.ChatHistoryID || chat.ID || chat.ReportHeadAIHistoryID || (tag && tag.ChatHistoryID);
    }
}


