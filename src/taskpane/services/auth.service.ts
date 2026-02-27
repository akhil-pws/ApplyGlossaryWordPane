import { CONFIG } from "../utils/config";
import { loginUser } from "../draft/draft.api";
import { PersistenceService } from "./persistence.service";
import { StoreService } from "./store.service";

export class AuthService {
    private static readonly USER_ROLE_KEY = 'userRole';
    private static readonly STYLE_KEY = 'tableStyle';
    private static readonly PALETTE_KEY = 'colorPallete';
    private static readonly TAG_ID_KEY = 'currentChatTagId';

    static getStoredToken(): string | null {
        // We still check sessionStorage for immediate token access if needed, 
        // but preferred way is restoreSession
        return sessionStorage.getItem('token');
    }

    /**
     * Restores session. Priority: sessionStorage (refresh) > PersistenceService & Office.settings (reopen).
     */
    static async restoreSession(): Promise<any> {
        // Check SessionStorage first (survives refresh)
        const sessionToken = sessionStorage.getItem('token');
        if (sessionToken) {
            console.log("AuthService: Restoring session from sessionStorage");
            return {
                jwt: sessionToken,
                userRole: JSON.parse(sessionStorage.getItem(this.USER_ROLE_KEY) || '{}'),
                tableStyle: sessionStorage.getItem(this.STYLE_KEY),
                colorPallete: JSON.parse(sessionStorage.getItem(this.PALETTE_KEY) || 'null'),
                userId: sessionStorage.getItem('userId'),
                mode: sessionStorage.getItem('mode') || 'Home',
                currentChatTagId: sessionStorage.getItem(this.TAG_ID_KEY) ? Number(sessionStorage.getItem(this.TAG_ID_KEY)) : null,
                tagDraft: JSON.parse(sessionStorage.getItem('tagDraft') || 'null'),
                summaryTagDraft: JSON.parse(sessionStorage.getItem('summaryTagDraft') || 'null'),
                view: sessionStorage.getItem('view') || 'Home',
                viewParams: JSON.parse(sessionStorage.getItem('viewParams') || '{}')
            };
        }

        // Check PersistenceService & Office.settings (survives closing)
        const state = await PersistenceService.loadState();
        if (state) {
            console.log("AuthService: Restoring session from decrypted localStorage");

            // Load document-specific settings
            const settingsKeys = [this.STYLE_KEY, this.PALETTE_KEY, this.TAG_ID_KEY, 'theme', 'tagDraft', 'summaryTagDraft', 'mode'];
            const docSettings = PersistenceService.loadSettings(settingsKeys);

            const mergedState = {
                ...state,
                ...docSettings,
                // Ensure correct types
                tableStyle: docSettings[this.STYLE_KEY] || state.tableStyle,
                colorPallete: docSettings[this.PALETTE_KEY] || state.colorPallete,
                currentChatTagId: docSettings[this.TAG_ID_KEY] !== undefined ? docSettings[this.TAG_ID_KEY] : state.currentChatTagId,
                mode: docSettings['mode'] || state.mode || 'Home'
            };

            // For legacy components that might still read from sessionStorage
            this.syncToSessionStorage(mergedState);
            return mergedState;
        }
        return null;
    }

    private static syncToSessionStorage(state: any): void {
        if (state.jwt) sessionStorage.setItem('token', state.jwt);
        if (state.userRole) sessionStorage.setItem(this.USER_ROLE_KEY, JSON.stringify(state.userRole));
        if (state.userId) sessionStorage.setItem('userId', state.userId.toString());
        if (state.tableStyle) sessionStorage.setItem(this.STYLE_KEY, state.tableStyle);
        if (state.colorPallete) sessionStorage.setItem(this.PALETTE_KEY, JSON.stringify(state.colorPallete));
        if (state.currentChatTagId !== undefined) {
            if (state.currentChatTagId === null || state.currentChatTagId === -1) {
                sessionStorage.removeItem(this.TAG_ID_KEY);
            } else {
                sessionStorage.setItem(this.TAG_ID_KEY, state.currentChatTagId.toString());
            }
        }
        if (state.tagDraft) sessionStorage.setItem('tagDraft', JSON.stringify(state.tagDraft));
        if (state.summaryTagDraft) sessionStorage.setItem('summaryTagDraft', JSON.stringify(state.summaryTagDraft));
        if (state.view) sessionStorage.setItem('view', state.view);
        if (state.viewParams) sessionStorage.setItem('viewParams', JSON.stringify(state.viewParams));
        if (state.mode) sessionStorage.setItem('mode', state.mode);
        if (state.theme) sessionStorage.setItem('theme', state.theme);
    }

    static async login(organization: string, username: string, password: string): Promise<{ success: boolean, message?: string, data?: any }> {
        try {
            console.log(`Logging in to ${CONFIG.dataUrl}`);
            const data = await loginUser(organization, username, password);

            if (data.Status === true && data['Data']) {
                if (data['Data'].ResponseStatus) {
                    const jwt = data.Data.Token;
                    const userRole = data.Data.UserRole;
                    const userId = data.Data.UserID;

                    const store = StoreService.getInstance();

                    const state = {
                        jwt: jwt,
                        userRole: userRole,
                        userId: userId,
                        mode: store.mode || 'Home',
                        tableStyle: store.tableStyle,
                        colorPallete: store.colorPallete,
                        organizationName: organization,
                        WorkbenchID: store.WorkbenchID
                    };

                    // Store interactions in Persistence (Encrypted LocalStorage)
                    await PersistenceService.saveState(state);

                    // Also sync to Session for legacy support
                    this.syncToSessionStorage(state);

                    return {
                        success: true,
                        data: {
                            token: jwt,
                            userRole: userRole,
                            userId: userId,
                            raw: data.Data
                        }
                    };
                } else {
                    return { success: false, message: "An error occurred during login. Please try again." };
                }
            } else {
                return { success: false, message: "An error occurred during login. Please try again." };
            }
        } catch (error) {
            console.error('Error during login:', error);
            return { success: false, message: "An error occurred during login. Please try again." };
        }
    }

    static logout(): void {
        PersistenceService.clearState();
        sessionStorage.clear();
        console.log("Logged out and state cleared");
    }
}
