
export class StoreService {
    private static instance: StoreService;

    // State Variables
    public jwt: string = '';
    public UserRole: any = {};
    public WorkbenchID: string = '';
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
    public sponsorID: string = '0';
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
    public isSyncEnabled: boolean = false;
    public isDraftGenerated: boolean = false;

    // Navigation State
    public view: string = 'Home'; // Home, AddTag, Sync, Glossary, Formatting, AIHistory
    public viewParams: any = {};

    // Draft States for forms
    public tagDraft: any = {
        name: '',
        description: '',
        prompt: '',
        saveGlobally: true,
        isAvailableForAll: true,
        selectedSponsors: [],
        selectedSources: []
    };
    public summaryTagDraft: any = {
        name: '',
        description: '',
        prompt: '',
        saveGlobally: true,
        isAvailableForAll: true,
        selectedSponsors: [],
        selectedSources: []
    };

    private constructor() { }

    public static getInstance(): StoreService {
        if (!StoreService.instance) {
            console.log("StoreService: Initializing new instance");
            StoreService.instance = new StoreService();
        }

        if (!StoreService.instance) {
            console.error("StoreService: Critical Error - instance is still undefined after initialization attempt!");
            // Fallback: create it again or throw an error to fail fast
            StoreService.instance = new StoreService();
        }

        return StoreService.instance;
    }

    private saveTimeout: any = null;

    /**
     * Debounced save of the current state.
     */
    public async save(): Promise<void> {
        if (this.saveTimeout) {
            clearTimeout(this.saveTimeout);
        }

        this.saveTimeout = setTimeout(async () => {
            const { PersistenceService } = await import("./persistence.service");

            // 1. Session-wide state (Encrypted LocalStorage)
            const encryptedState = {
                jwt: this.jwt,
                userRole: this.UserRole,
                userId: this.userId,
                organizationName: this.organizationName,
                WorkbenchID: this.WorkbenchID,
                view: this.view,
                viewParams: this.viewParams
            };
            await PersistenceService.saveState(encryptedState);

            // 2. Document-specific state (Office Settings)
            const documentSettings = {
                mode: this.mode,
                tableStyle: this.tableStyle,
                colorPallete: this.colorPallete,
                currentChatTagId: this.currentChatTagId,
                theme: this.theme,
                tagDraft: this.tagDraft,
                summaryTagDraft: this.summaryTagDraft
            };
            PersistenceService.saveSettings(documentSettings);

            // 3. Sync to SessionStorage (immediate refresh survival)
            this.syncToSessionStorage();

            console.log("StoreService: State, settings, and sessionStorage auto-saved.");
        }, 500); // 500ms debounce
    }

    private syncToSessionStorage(): void {
        sessionStorage.setItem('token', this.jwt);
        if (this.UserRole) sessionStorage.setItem('userRole', JSON.stringify(this.UserRole));
        if (this.userId) sessionStorage.setItem('userId', this.userId.toString());
        if (this.mode) sessionStorage.setItem('mode', this.mode);
        if (this.tableStyle) sessionStorage.setItem('tableStyle', this.tableStyle);
        if (this.colorPallete) sessionStorage.setItem('colorPallete', JSON.stringify(this.colorPallete));
        if (this.theme) sessionStorage.setItem('theme', this.theme);
        if (this.view) sessionStorage.setItem('view', this.view);
        if (this.viewParams) sessionStorage.setItem('viewParams', JSON.stringify(this.viewParams));

        if (this.currentChatTagId > 0) {
            sessionStorage.setItem('currentChatTagId', this.currentChatTagId.toString());
        } else {
            sessionStorage.removeItem('currentChatTagId');
        }

        if (this.tagDraft) sessionStorage.setItem('tagDraft', JSON.stringify(this.tagDraft));
        if (this.summaryTagDraft) sessionStorage.setItem('summaryTagDraft', JSON.stringify(this.summaryTagDraft));
    }

    /**
     * Resets the store (for logout).
     */
    public reset(): void {
        this.jwt = '';
        this.UserRole = {};
        this.userId = 0;
        this.currentChatTagId = -1;
        this.view = 'Home';
        this.viewParams = {};
        this.tagDraft = { name: '', description: '', prompt: '', saveGlobally: true, isAvailableForAll: true, selectedSponsors: [], selectedSources: [] };
        this.summaryTagDraft = { name: '', description: '', prompt: '', saveGlobally: true, isAvailableForAll: true, selectedSponsors: [], selectedSources: [] };
    }
}
