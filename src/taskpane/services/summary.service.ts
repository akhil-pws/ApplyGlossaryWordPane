import { removeQuotes, updateEditorFinalTable } from "../draft/draft-functions";
import { getSummaryTagHistory } from "../summary/summary.api";
import { StoreService } from "./store.service";

export class summaryService {

    static async fetchSummaryAIHistory(tag: any): Promise<any[]> {
        const store = StoreService.getInstance();
        try {
            const data = await getSummaryTagHistory(tag.ID || tag.ReportHeadSummaryTagID, store.jwt);

            if (data.Status && data.Data && data.Data.length > 0) {
                tag.ReportHeadAIHistoryList = data['Data'] || [];
                tag.FilteredReportHeadAIHistoryList = [];

                const latestHistory = tag.ReportHeadAIHistoryList[0];
                const rawSources = latestHistory.SourceVector || latestHistory.SourceValue || '';
                const sourceIds = Array.isArray(rawSources)
                    ? rawSources.map(String)
                    : String(rawSources).split(',').map(s => s.trim()).filter(Boolean);

                const selectedSources = store.sourceSummaryList.filter((list: any) =>
                    sourceIds.includes(String(list.VectorID))
                );

                tag.SourceName = selectedSources.map((item: any) => item.FileName || item.SourceName);
                tag.Sources = [...tag.SourceName];
                tag.TempSourceValue = selectedSources.map((item: any) =>
                    item.VectorID ? String(item.VectorID) : item.SourceValue
                );

                tag.ReportHeadAIHistoryList.forEach((historyList: any) => {
                    historyList.Response = removeQuotes(historyList.Response);
                    tag.FilteredReportHeadAIHistoryList.unshift(historyList);
                });
                return tag.FilteredReportHeadAIHistoryList;
            } else {
                console.warn("No Summary AI history available.");
                return [];
            }
        } catch (error) {
            console.error('Error fetching Summary AI history:', error);
            return [];
        }
    }
}
