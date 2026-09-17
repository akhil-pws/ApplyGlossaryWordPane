export interface ChatEvidence {
    ID?: number | string;
    FileName: string;
    PageNumber: number | string;
    Sentence: string;
    ReportHeadID?: number | string;
    ReportHeadAIHistoryID?: number | string;
    VectorChunkID?: number | string;
    CreatedDate?: string;
}

export interface ChatMessage {
    ID: number | string;
    ChatHistoryID?: number | string;
    ReportHeadID?: number | string;
    ReportHeadGroupKeyID?: number | string;
    Prompt: string;
    Response: string;
    FormattedResponse?: string;
    Selected: number;
    Sources?: string[];
    SourceValue?: string[] | string;
    SourceVector?: string[] | string;
    CreatedByName?: string;
    CreatedDate?: string;
    DocumentInstruction?: string;
    Evidences?: ChatEvidence[];
    [key: string]: any;
}

export interface ChatSession {
    id: string | number;
    chatHistoryId: number | string;
    title: string;
    createdAt: string;
    sources: string[];
    sourceValues: string[];
    history: ChatMessage[];
}

export interface AITag {
    ID: number;
    DisplayName: string;
    AIFlag: number;
    ChatSessions?: ChatSession[];
    ActiveSessionIndex?: number;
    ReportHeadAIHistoryList?: ChatMessage[];
    FilteredReportHeadAIHistoryList?: ChatMessage[];
    Sources?: string[];
    SourceName?: string[];
    SourceValue?: any;
    SourceValueID?: any;
    TempSourceValue?: string[];
    [key: string]: any; // Allow flexibility for now
}

export interface TagUpdate {
    tagName: string;
    status: 'added' | 'removed';
}

