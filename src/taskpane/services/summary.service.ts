import { removeQuotes } from "../draft/draft-functions";
import { getSummaryTagHistory } from "../summary/summary.api";
import { StoreService } from "./store.service";

export class summaryService {

    static async fetchSummaryAIHistory(tag: any): Promise<any[]> {
        const store = StoreService.getInstance();
        try {
            const data = await getSummaryTagHistory(tag.ID, store.jwt);

            if (data.Status && data.Data) {
                tag.ReportHeadAIHistoryList = data['Data'] || [];
                tag.FilteredReportHeadAIHistoryList = [];
                tag.SourceValueID = tag.ReportHeadAIHistoryList[0].SourceValue;
                const selectedSources = store.sourceSummaryList.filter((list: any) =>
                    tag.SourceVector.includes(String(list.VectorID))
                );
                tag.SourceName = selectedSources.map((item: any) => {
                    return item.SourceName;
                }
                );
                tag.Sources = tag.SourceName.join(',');
                tag.TempSourceValue = selectedSources.map((item: any) => {
                    return item.VectorID ? String(item.VectorID) : item.SourceValue;
                }

                );

                tag.ReportHeadAIHistoryList.forEach((historyList: any) => {
                    historyList.Response = removeQuotes(historyList.Response);
                    tag.FilteredReportHeadAIHistoryList.unshift(historyList);
                }

                );
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
