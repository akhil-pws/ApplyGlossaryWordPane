export interface GlossaryTerm {
    clinicalTerm: string;
    layTerm: string;
    defination?: string; // Keeping as per potential API typo or standard
    [key: string]: any;
}
