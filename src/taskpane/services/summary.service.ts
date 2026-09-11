import { removeQuotes, updateEditorFinalTable } from "../draft/draft-functions";
import { getSummaryTagHistory } from "../summary/summary.api";
import { AIService } from "./ai.service";
import { StoreService } from "./store.service";

export class summaryService {

    static async fetchSummaryAIHistory(tag: any): Promise<any[]> {
        const store = StoreService.getInstance();
        try {
            const data = await getSummaryTagHistory(tag.ID || tag.ReportHeadSummaryTagID, store.jwt);

            if (data.Status && data.Data && Array.isArray(data.Data) && data.Data.length > 0) {
                const sessions = AIService.groupHistoryIntoSessions(data.Data, tag, store, "Summary");
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
                console.warn("No Summary AI history available.");
                tag.ChatSessions = [];
                tag.ActiveSessionIndex = 0;
                tag.FilteredReportHeadAIHistoryList = [];
                tag.ReportHeadAIHistoryList = [];
                return [];
            }
            return [];
        } catch (error) {
            console.error('Error fetching Summary AI history:', error);
            tag.ChatSessions = [];
            tag.ActiveSessionIndex = 0;
            tag.FilteredReportHeadAIHistoryList = [];
            tag.ReportHeadAIHistoryList = [];
            return [];
        }
    }
}

