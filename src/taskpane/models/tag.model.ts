export interface AITag {
    ID: number;
    DisplayName: string;
    AIFlag: number;
    [key: string]: any; // Allow flexibility for now
}

export interface TagUpdate {
    tagName: string;
    status: 'added' | 'removed';
}
