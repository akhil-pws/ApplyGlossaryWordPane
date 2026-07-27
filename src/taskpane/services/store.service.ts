
import { DocStorage, setDocStorageId } from '../utils/doc-storage';

export class StoreService {
    private static instance: StoreService;

    private static readonly STORAGE_KEY = 'link_ai_store_v1';

    // State Variables
    public jwt: string = '';
    public UserRole: any = {};
    public documentID: string = '';
    public organizationName: string = '';
    public documentInstruction: string = '';
    public aiTagList: any[] = [];
    public summaryTagList: any[] = [];
    public imageList: any[] = [];
    public initialised: boolean = true;
    public availableKeys: any[] = [];
    public promptBuilderList: any[] = [];
    public glossaryName: string = '';
    public isGlossaryActive: boolean = false;
    public GroupName: string = '';
    public layTerms: any[] = [];
    public dataList: any = [];
    public isTagUpdating: boolean = false;
    public capturedFormatting: any = {};
    public emptyFormat: boolean = false;
    public isNoFormatTextAvailable: boolean = false;
    public clientId: string = '0';
    public userId: number = 0;
    public clientList: any[] = [];
    public currentYear: number = new Date().getFullYear();
    public sourceList: any;
    public sourceSummaryList: any;
    public filteredGlossaryTerm: any;
    public selectedNames: any[] = [];
    public isPendingResponse: boolean = false;
    public theme: string = 'Light';
    public mode: string = 'Home';
    public tableStyle: string = 'Plain Table 5';
    public defaultTextStyle: string = 'Normal';
    public customizedTextStyle: any = null;
    public colorPallete: any = {
        "Header": '#FFFFFF',
        "Primary": '#FFFFFF',
        "Secondary": '#FFFFFF',
        "Customize": true,
        "IsHeaderBold": true,
        "IsSideHeaderBold": false
    };
    public environment: string = '';

    public customTableStyle: any[] = [];
    public customizedStyles: any[] = [];
    public customTextStylesLoaded: boolean = false;
    public currentChatTagId: number = -1;
    public isReversed: boolean = false;
    public reprocessingTagIds: { [key: string]: boolean } = {};

    private constructor() {
        // Do NOT load from storage here — documentID is not yet known.
        // Call initForDocument(docId) after retrieving the document properties.
    }

    public static getInstance(): StoreService {
        if (!StoreService.instance) {
            StoreService.instance = new StoreService();
        }
        return StoreService.instance;
    }

    /**
     * Must be called once, after documentID is known (from document properties).
     * Scopes all localStorage I/O to this document and rehydrates saved state.
     */
    public initForDocument(docId: string, environment?: string): void {
        this.documentID = docId;
        this.environment = environment || '';
        const storageId = environment ? `${docId}_${environment}` : docId;
        setDocStorageId(storageId);
        this.loadFromStorage();
    }

    /**
     * Persist current critical state to localStorage (scoped to this document).
     */
    public saveToStorage(): void {
        const dataToSave = {
            jwt: this.jwt,
            UserRole: this.UserRole,
            documentID: this.documentID,
            organizationName: this.organizationName,
            documentInstruction: this.documentInstruction,
            theme: this.theme,
            mode: this.mode,
            userId: this.userId,
            tableStyle: this.tableStyle,
            colorPallete: this.colorPallete,
            clientId: this.clientId,
            defaultTextStyle: this.defaultTextStyle,
            customizedTextStyle: this.customizedTextStyle
        };
        DocStorage.setItem(StoreService.STORAGE_KEY, JSON.stringify(dataToSave));
    }

    /**
     * Rehydrate state from localStorage (scoped to this document).
     */
    private loadFromStorage(): void {
        try {
            const stored = DocStorage.getItem(StoreService.STORAGE_KEY);
            if (stored) {
                const data = JSON.parse(stored);
                if (data.jwt) this.jwt = data.jwt;
                if (data.UserRole) this.UserRole = data.UserRole;
                if (data.organizationName) this.organizationName = data.organizationName;
                if (data.documentInstruction) this.documentInstruction = data.documentInstruction;
                if (data.theme) this.theme = data.theme;
                if (data.mode) this.mode = data.mode;
                if (data.userId) this.userId = data.userId;
                if (data.tableStyle) this.tableStyle = data.tableStyle;
                if (data.colorPallete) this.colorPallete = data.colorPallete;
                if (data.clientId) this.clientId = data.clientId;
                if (data.defaultTextStyle) this.defaultTextStyle = data.defaultTextStyle;
                if (data.customizedTextStyle) this.customizedTextStyle = data.customizedTextStyle;
            }
        } catch (error) {
            console.error("Failed to load state from storage", error);
        }
    }

    /**
     * Clear stored state for this document.
     */
    public clearStorage(): void {
        DocStorage.removeItem(StoreService.STORAGE_KEY);
        // Reset local variables
        this.jwt = '';
        this.UserRole = {};
        this.userId = 0;
        this.documentInstruction = '';
        this.isReversed = false;
        this.aiTagList = [];
        this.summaryTagList = [];
        this.imageList = [];
        this.dataList = [];
        this.selectedNames = [];
        this.currentChatTagId = -1;
        this.reprocessingTagIds = {};
        this.isPendingResponse = false;
        this.isTagUpdating = false;
        this.tableStyle = 'Plain Table 5';
        this.defaultTextStyle = 'Normal';
        this.customizedTextStyle = null;
        this.colorPallete = {
            "Header": '#FFFFFF',
            "Primary": '#FFFFFF',
            "Secondary": '#FFFFFF',
            "Customize": true,
            "IsHeaderBold": true,
            "IsSideHeaderBold": false
        };
        this.mode = 'Home';
        this.customizedStyles = [];
        this.customTextStylesLoaded = false;
    }
}
