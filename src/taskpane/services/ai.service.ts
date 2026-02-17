import { AITag } from "../models/tag.model";
import { CONFIG } from "../utils/config";
import { StoreService } from "./store.service";
import { getAiHistory, addAiHistory } from "../draft/draft.api";
import { removeQuotes, updateEditorFinalTable } from "../draft/draft-functions";
import { generateCheckboxHistory } from "../draft/home";
import { addSummaryHistory } from "../summary/summary.api";

export class AIService {

    static async fetchAIHistory(tag: any): Promise<any[]> {
        const store = StoreService.getInstance();
        try {
            const data = await getAiHistory(tag.GroupKeyID, store.jwt);

            if (data.Status && data.Data) {
                tag.ReportHeadAIHistoryList = data['Data'] || [];
                tag.FilteredReportHeadAIHistoryList = [];
                tag.SourceValueID = tag.ReportHeadAIHistoryList[0].SourceVector;

                const selectedSources = store.sourceList.filter((list: any) =>
                    tag.SourceValueID.includes(String(list.VectorID))
                );

                tag.SourceName = selectedSources.map((item: any) => {
                    return item.SourceName;
                });
                tag.Sources = tag.SourceName.join(',');
                tag.TempSourceValue = selectedSources.map((item: any) => {
                    return item.VectorID ? String(item.VectorID) : item.SourceValue;
                });

                tag.ReportHeadAIHistoryList.forEach((historyList: any) => {
                    historyList.Response = removeQuotes(historyList.Response);
                    tag.FilteredReportHeadAIHistoryList.unshift(historyList);
                });
                return tag.FilteredReportHeadAIHistoryList;
            } else {
                console.warn("No AI history available.");
                return [];
            }
        } catch (error) {
            console.error('Error fetching AI history:', error);
            return [];
        }
    }

    static async sendPrompt(tag: any, prompt: string, type: "Summary" | "AITag" = "AITag"): Promise<void> {
        const store = StoreService.getInstance();

        if (prompt !== '' && !store.isTagUpdating) {
            store.isTagUpdating = true;

            // UI Updates via direct DOM or another UI service? 
            // For now, keeping direct DOM manipulation as in original code, but cleaner would be callbacks.
            const iconelement = document.getElementById(`sendPromptButton`);
            if (iconelement) iconelement.innerHTML = `<i class="fa fa-spinner fa-spin text-white"></i>`;

            let payload: any;
            if (type === 'Summary') {
                payload = {
                    Prompt: prompt,
                    WorkbenchSummaryTagID: tag.WorkbenchSummaryTagID,
                    WorkbenchID: Number(store.WorkbenchID),
                    Response: "",
                    SourceVector: tag.SourceValueID || "",
                    SummarySourceID: tag.WorkbenchSourceID || 1
                };
            } else {

                payload = {
                    WorkbenchSourceID: tag.WorkbenchSourceID || 1,
                    Prompt: prompt,
                    GroupKey: tag.Name,
                    WorkbenchGroupKeyID: tag.GroupKeyID,
                    WorkbenchID: Number(store.WorkbenchID),
                    WorkbenchName: store.dataList?.WorkbenchName || "",
                    Response: "",
                    SourceVector: tag.SourceValueID || "",
                    FormattedResponse: ""
                };

            }

            try {
                store.isPendingResponse = true;
                const data = type === 'Summary'
                    ? await addSummaryHistory(payload, store.jwt)
                    : await addAiHistory(payload, store.jwt);

                if (data['Data'] && data['Data'] !== 'false') {
                    tag.ReportHeadAIHistoryList = JSON.parse(JSON.stringify(data['Data']));
                    tag.FilteredReportHeadAIHistoryList = [];

                    tag.ReportHeadAIHistoryList.forEach((historyList: any) => {
                        historyList.Response = removeQuotes(historyList.Response);
                        tag.FilteredReportHeadAIHistoryList.unshift(historyList);
                    });
                    const chat = tag.ReportHeadAIHistoryList[0];

                    const finalResponse = chat.FormattedResponse
                        ? '\n' + updateEditorFinalTable(chat.FormattedResponse)
                        : chat.Response;

                    tag.ComponentKeyDataType = chat.FormattedResponse ? 'TABLE' : 'TEXT';
                    // tag.UserValue = finalResponse;
                    tag.Response = finalResponse;
                    tag.text = finalResponse;


                    // Update lists in Store
                    store.aiTagList.forEach((currentTag: any) => {
                        if (currentTag.ID === tag.ID) {
                            AIService.updateTagWithChat(currentTag, chat, tag.IsApplied);
                        }
                    });

                    store.availableKeys.forEach((currentTag: any) => {
                        if (currentTag.ID === tag.ID) {
                            AIService.updateTagWithChat(currentTag, chat, tag.IsApplied);
                        }
                    });

                    const appbody = document.getElementById('app-body');
                    if (appbody) appbody.innerHTML = await generateCheckboxHistory(tag, type);
                    store.isPendingResponse = false;
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
        const isTable = chat.FormattedResponse !== '';
        const finalResponse = chat.FormattedResponse
            ? '\n' + updateEditorFinalTable(chat.FormattedResponse)
            : chat.Response;

        currentTag.ComponentKeyDataType = isTable ? 'TABLE' : 'TEXT';
        // currentTag.UserValue = finalResponse;
        currentTag.Response = finalResponse;
        currentTag.text = finalResponse;
        currentTag.IsApplied = isApplied;
    }
}

