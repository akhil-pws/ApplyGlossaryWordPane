import { AITag, ChatMessage, ChatSession } from "../models/tag.model";
import { StoreService } from "./store.service";
import { getAiHistory, addAiHistory } from "../draft/draft.api";
import { removeQuotes, updateEditorFinalTable } from "../draft/draft-functions";
import { generateCheckboxHistory } from "../draft/home";
import { addSummaryHistory } from "../summary/summary.api";

export class AIService {

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
            // Determine session title from the first prompt in this thread
            const firstMsg = messages.find(m => m.Prompt && m.Prompt.trim() !== '') || messages[0];
            let title = `Chat ${hId}`;
            if (firstMsg?.Prompt) {
                title = firstMsg.Prompt.trim().replace(/\s+/g, ' ');
            }

            // Resolve sources from message's SourceValue / SourceVector
            let sessionSources: string[] = [];
            let sessionSourceValues: string[] = [];

            const rawSourceVal = firstMsg?.SourceValue || firstMsg?.SourceVector || tag?.SourceValue || tag?.SourceVector;
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
                history: messages
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
            const data = await getAiHistory(tag.ID, store.jwt);

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

                    const selectedChat = activeSession.history.find((item: any) => item.Selected === 1) || activeSession.history[0];
                    if (selectedChat) {
                        const finalResponse = selectedChat.FormattedResponse
                            ? '\n' + updateEditorFinalTable(selectedChat.FormattedResponse)
                            : selectedChat.Response;

                        tag.ComponentKeyDataType = selectedChat.FormattedResponse ? 'TABLE' : 'TEXT';
                        tag.UserValue = finalResponse;
                        tag.EditorValue = finalResponse;
                        tag.text = finalResponse;
                    }
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

        tag.FilteredReportHeadAIHistoryList = activeSession.history;
        tag.ReportHeadAIHistoryList = activeSession.history;
        tag.Sources = [...activeSession.sources];
        tag.SourceName = [...activeSession.sources];
        tag.TempSourceValue = [...activeSession.sourceValues];
        tag.SourceValueID = activeSession.sourceValues;

        // Sync selected response
        const selectedChat = activeSession.history.find((item: any) => item.Selected === 1) || activeSession.history[0];
        if (selectedChat) {
            const finalResponse = selectedChat.FormattedResponse
                ? '\n' + updateEditorFinalTable(selectedChat.FormattedResponse)
                : selectedChat.Response;

            tag.ComponentKeyDataType = selectedChat.FormattedResponse ? 'TABLE' : 'TEXT';
            tag.UserValue = finalResponse;
            tag.EditorValue = finalResponse;
            tag.text = finalResponse;
        }
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
                    ChatHistoryID: targetChatHistoryId
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
                    DocumentInstruction: documentInstruction
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

                        const chat = currentActive.history.find((item: any) => item.Selected === 1) || currentActive.history[0];
                        if (chat) {
                            const isTable = chat.FormattedResponse && chat.FormattedResponse !== '';
                            const finalResponse = isTable
                                ? '\n' + updateEditorFinalTable(chat.FormattedResponse)
                                : chat.Response;

                            tag.ComponentKeyDataType = isTable ? 'TABLE' : 'TEXT';
                            tag.UserValue = finalResponse;
                            tag.EditorValue = finalResponse;
                            tag.text = finalResponse;

                            // Update lists in Store
                            if (type === 'Summary') {
                                store.summaryTagList?.forEach((currentTag: any) => {
                                    const currentId = currentTag.ID || currentTag.ReportHeadSummaryTagID;
                                    const tagId = tag.ID || tag.ReportHeadSummaryTagID;
                                    if (currentId === tagId) {
                                        AIService.updateTagWithChat(currentTag, chat, tag.IsApplied);
                                    }
                                });
                            } else {
                                store.aiTagList?.forEach((currentTag: any) => {
                                    if (currentTag.ID === tag.ID) {
                                        AIService.updateTagWithChat(currentTag, chat, tag.IsApplied);
                                    }
                                });

                                store.availableKeys?.forEach((currentTag: any) => {
                                    if (currentTag.ID === tag.ID) {
                                        AIService.updateTagWithChat(currentTag, chat, tag.IsApplied);
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

            } catch (error) {
                if (iconelement) iconelement.innerHTML = `<i class="fa fa-paper-plane text-white"></i>`;
                store.isTagUpdating = false;
                store.isPendingResponse = false;
                console.error('Error sending AI prompt:', error);
            }
        } else {
            console.error('No empty prompt allowed or tag updating');
        }
    }

    private static updateTagWithChat(currentTag: any, chat: any, isApplied: any) {
        const isTable = chat.FormattedResponse && chat.FormattedResponse !== '';
        const finalResponse = isTable
            ? '\n' + updateEditorFinalTable(chat.FormattedResponse)
            : chat.Response;

        currentTag.ComponentKeyDataType = isTable ? 'TABLE' : 'TEXT';
        currentTag.UserValue = finalResponse;
        currentTag.EditorValue = finalResponse;
        currentTag.text = finalResponse;
        currentTag.IsApplied = isApplied;
    }
}


