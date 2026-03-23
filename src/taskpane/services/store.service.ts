
export class StoreService {
    private static instance: StoreService;

    private static readonly STORAGE_KEY = 'link_ai_store_v1';

    // State Variables
    public jwt: string = '';
    public UserRole: any = {};
    public documentID: string = '';
    public organizationName: string = '';
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
    public colorPallete: any = {
        "Header": '#FFFFFF',
        "Primary": '#FFFFFF',
        "Secondary": '#FFFFFF',
        "Customize": true,
        "IsHeaderBold": true,
        "IsSideHeaderBold": false
    };
    public customTableStyle: any[] = [];
    public currentChatTagId: number = -1;
    public isReversed: boolean = false;
    public reprocessingTagIds: { [key: string]: boolean } = {};

    private constructor() {
        this.loadFromStorage();
    }

    public static getInstance(): StoreService {
        if (!StoreService.instance) {
            StoreService.instance = new StoreService();
        }
        return StoreService.instance;
    }

    /**
     * Persist current critical state to localStorage.
     */
    public saveToStorage(): void {
        const dataToSave = {
            jwt: this.jwt,
            UserRole: this.UserRole,
            documentID: this.documentID,
            organizationName: this.organizationName,
            theme: this.theme,
            mode: this.mode,
            userId: this.userId,
            tableStyle: this.tableStyle,
            colorPallete: this.colorPallete,
            clientId: this.clientId
        };
        localStorage.setItem(StoreService.STORAGE_KEY, JSON.stringify(dataToSave));
    }

    /**
     * Rehydrate state from localStorage.
     */
    private loadFromStorage(): void {
        try {
            const stored = localStorage.getItem(StoreService.STORAGE_KEY);
            if (stored) {
                const data = JSON.parse(stored);
                if (data.jwt) this.jwt = data.jwt;
                if (data.UserRole) this.UserRole = data.UserRole;
                if (data.documentID) this.documentID = data.documentID;
                if (data.organizationName) this.organizationName = data.organizationName;
                if (data.theme) this.theme = data.theme;
                if (data.mode) this.mode = data.mode;
                if (data.userId) this.userId = data.userId;
                if (data.tableStyle) this.tableStyle = data.tableStyle;
                if (data.colorPallete) this.colorPallete = data.colorPallete;
                if (data.clientId) this.clientId = data.clientId;
            }
        } catch (error) {
            console.error("Failed to load state from storage", error);
        }
    }

    /**
     * Clear stored state.
     */
    public clearStorage(): void {
        localStorage.removeItem(StoreService.STORAGE_KEY);
        // Reset local variables
        this.jwt = '';
        this.UserRole = {};
        this.userId = 0;
    }
}
