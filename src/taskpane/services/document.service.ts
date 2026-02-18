/* global Word */
import { DocumentProperty } from "../models/app.model";
import { getReportById, getAllClients, getAllPromptTemplates, getGeneralImages, getReportHeadImageById } from "../draft/draft.api";
import { updateEditorFinalTable, mapImagesToComponentObjects } from "../draft/draft-functions";

export class DocumentService {

    private static transformDocumentName(value: string): string {
        if (!value || value.trim() === '') return value;
        const parts = value.split('_');
        if (parts.length <= 1) return value;
        return parts.slice(1).join('_').replace(/%20/g, ' ').replace(/%25/g, '%');
    }

    static async loadReportData(documentId: string, jwt: string, userId: string): Promise<any> {
        try {
            console.log(`Fetching report data for ID: ${documentId}`);
            const data = await getReportById(documentId, jwt);

            if (!data.Status || !data.Data) {
                throw new Error("Failed to fetch report data");
            }

            const dataList = data.Data;

            // Basic processing
            if (!dataList.SourceTypeList) dataList.SourceTypeList = [];

            const sourceList = dataList.SourceTypeList
                .filter((item: any) => item.SourceValue !== '' && item.AIFlag === 1)
                .map((item: any) => ({
                    ...item,
                    SourceName: decodeURIComponent(DocumentService.transformDocumentName(item.SourceValue))
                }));

            const clientId = dataList.ClientID;
            const aiGroup = dataList.Group.find((element: any) => element.DisplayName === 'AIGroup');
            const groupName = aiGroup ? aiGroup.Name : '';
            const aiTagList = aiGroup ? aiGroup.GroupKey : [];

            // Image processing - Deferred to background (getImages)

            // Available Keys filtering
            let availableKeys = dataList.GroupKeyAll.filter((element: any) =>
                element.ComponentKeyDataType === 'TABLE' ||
                element.ComponentKeyDataType === 'TEXT'
            );

            const imageList: any[] = [];

            // Apply transformations to keys (updateEditorFinalTable)
            const processKey = (key: any) => {
                if (key.AIFlag === 1) {
                    const regex = /<TableStart>([\s\S]*?)<TableEnd>/gi;
                    if (regex.exec(key.EditorValue) !== null) {
                        key.EditorValue = updateEditorFinalTable(key.EditorValue);
                        key.UserValue = key.EditorValue;
                        key.InitialTable = true;
                        key.ComponentKeyDataType = 'TABLE';
                    }
                }
            };

            availableKeys.forEach(processKey);
            aiTagList.forEach(processKey);

            return {
                dataList,
                availableKeys,
                sourceList,
                clientId,
                groupName,
                aiTagList,
                imageList,
                clientList: [], // Will be filled if needed or fetched separately
                promptBuilderList: [] // Fetched via loadPromptTemplates
            };

        } catch (error) {
            console.error("Error in loadReportData:", error);
            throw error;
        }
    }

    static async retrieveDocumentProperties(): Promise<{ documentID: string, organizationName: string } | null> {
        try {
            return await Word.run(async (context) => {
                const properties = context.document.properties.customProperties;
                properties.load("items");

                await context.sync();

                const property = properties.items.find(prop => prop.key === 'DocumentID');
                const orgName = properties.items.find(prop => prop.key === 'Organization');

                if (property && orgName) {
                    return {
                        documentID: property.value,
                        organizationName: orgName.value
                    };
                } else {
                    return null;
                }
            });
        } catch (error) {
            console.error("Error retrieving document properties", error);
            throw error;
        }
    }

    static async insertText(text: string): Promise<void> {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.insertParagraph(text, Word.InsertLocation.end);
            await context.sync();
        });
    }
}
