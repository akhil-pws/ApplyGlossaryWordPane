
export class StoreService {
    private static instance: StoreService;

    // State Variables
    public jwt: string = '';
    public UserRole: any = {};
    public documentID: string = '';
    public organizationName: string = '';
    public aiTagList: any[] = [];
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

    private constructor() { }

    public static getInstance(): StoreService {
        if (!StoreService.instance) {
            StoreService.instance = new StoreService();
        }
        return StoreService.instance;
    }
}
