import { CONFIG } from "../utils/config";
import { UserProfile } from "../models/user.model";
import { loginUser, checkLoginType, ssoLogin, ssoComplete, logoutUser, pingUser } from "../draft/draft.api";
import { StoreService } from "./store.service";
import { DocStorage } from "../utils/doc-storage";

export class AuthService {
    private static readonly TOKEN_KEY = 'user_token';
    private static readonly USER_ROLE_KEY = 'userRole';
    private static readonly STYLE_KEY = 'tableStyle';
    private static readonly PALETTE_KEY = 'colorPallete';
    private static readonly TEXT_STYLE_KEY = 'defaultTextStyle';
    private static readonly LOGIN_ID_KEY = 'loginId';
    private static pingIntervalId: any = null;

    static getStoredToken(): string | null {
        return DocStorage.getItem('token');
    }

    static startPingTimer(loginId?: string | number | null, token?: string | null): void {
        this.stopPingTimer();

        const currentLoginId = loginId || DocStorage.getItem(this.LOGIN_ID_KEY) || StoreService.getInstance().loginId;
        if (!currentLoginId) {
            return;
        }

        this.pingIntervalId = setInterval(async () => {
            const activeLoginId = DocStorage.getItem(this.LOGIN_ID_KEY) || StoreService.getInstance().loginId;
            const activeToken = DocStorage.getItem('token') || StoreService.getInstance().jwt;
            if (activeLoginId) {
                try {
                    await pingUser(activeLoginId, activeToken);
                } catch (error) {
                    console.error('Error during ping API call:', error);
                }
            } else {
                this.stopPingTimer();
            }
        }, 60000);
    }

    static stopPingTimer(): void {
        if (this.pingIntervalId) {
            clearInterval(this.pingIntervalId);
            this.pingIntervalId = null;
        }
    }

    static restoreSession(): any {
        const sessionToken = DocStorage.getItem('token');
        if (sessionToken) {
            // Check expiry (24 hours)
            const tokenLastUpdated = DocStorage.getItem('tokenLastUpdated');
            const now = new Date();
            if (tokenLastUpdated) {
                const lastUpdatedTime = new Date(tokenLastUpdated).getTime();
                const differenceInHours = (now.getTime() - lastUpdatedTime) / (1000 * 60 * 60);
                if (differenceInHours >= 24) {
                    console.log("JWT token expired (passed 24 hours). Logging out.");
                    DocStorage.removeItem('token');
                    this.logout();
                    return null;
                } else {
                    // Update timestamp to now, resetting the 24-hour expiration window
                    DocStorage.setItem('tokenLastUpdated', now.toISOString());
                }
            } else {
                // Initialize timestamp for legacy sessions
                DocStorage.setItem('tokenLastUpdated', now.toISOString());
            }

            const loginId = DocStorage.getItem(this.LOGIN_ID_KEY);
            if (loginId) {
                this.startPingTimer(loginId, sessionToken);
            }

            return {
                jwt: sessionToken,
                userRole: JSON.parse(DocStorage.getItem(this.USER_ROLE_KEY) || '{}'),
                tableStyle: DocStorage.getItem(this.STYLE_KEY),
                colorPallete: JSON.parse(DocStorage.getItem(this.PALETTE_KEY) || 'null'),
                userId: DocStorage.getItem('userId'),
                loginId: loginId,
                defaultTextStyle: DocStorage.getItem(this.TEXT_STYLE_KEY)
            };
        }
        return null;
    }

    static async login(organization: string, username: string, password: string): Promise<{ success: boolean, message?: string, data?: any }> {
        try {
            console.log(`Logging in to ${CONFIG.dataUrl}`);
            const data = await loginUser(organization, username, password);

            if (data.Status === true && data['Data']) {
                if (data['Data'].ResponseStatus) {
                    const jwt = data.Data.Token;
                    const userRole = data.Data.UserRole;
                    const userId = data.Data.ID;
                    const loginId = data.Data.LoginID;

                    // Store interactions in Local Storage for persistence (scoped to document)
                    DocStorage.setItem('token', jwt);
                    DocStorage.setItem(this.USER_ROLE_KEY, JSON.stringify(userRole));
                    DocStorage.setItem('userId', userId);
                    if (loginId !== undefined && loginId !== null) {
                        DocStorage.setItem(this.LOGIN_ID_KEY, String(loginId));
                    }
                    DocStorage.setItem('tokenLastUpdated', new Date().toISOString());

                    // Sync with StoreService and persist
                    const store = StoreService.getInstance();
                    store.clearStorage();
                    store.jwt = jwt;
                    store.UserRole = userRole;
                    store.userId = userId;
                    if (loginId !== undefined && loginId !== null) {
                        store.loginId = loginId;
                    }
                    store.saveToStorage();

                    if (loginId !== undefined && loginId !== null) {
                        this.startPingTimer(loginId, jwt);
                    }

                    return {
                        success: true,
                        data: {
                            token: jwt,
                            userRole: userRole,
                            userId: userId,
                            loginId: loginId,
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

    static async checkLoginType(organization: string, username: string): Promise<string> {
        try {
            const data = await checkLoginType(organization, username);
            if (data && data.Status === true && data.Data) {
                return data.Data.AuthenticationType || 'TrialAssure';
            }
            return 'TrialAssure';
        } catch (error) {
            console.error('Error during checkLoginType:', error);
            return 'TrialAssure';
        }
    }

    static async ssoLogin(organization: string, username: string): Promise<{ success: boolean, redirectUrl?: string, message?: string }> {
        try {
            const data = await ssoLogin(organization, username);
            if (data && data.Status === true && data.Data && data.Data.RedirectUrl) {
                return {
                    success: true,
                    redirectUrl: data.Data.RedirectUrl
                };
            } else {
                return {
                    success: false,
                    message: (data && data.Message) || 'Something went wrong during SSO login'
                };
            }
        } catch (error) {
            console.error('Error during ssoLogin:', error);
            return {
                success: false,
                message: 'Connection Lost'
            };
        }
    }

    static async ssoComplete(key: string): Promise<{ success: boolean, message?: string, data?: any }> {
        try {
            const data = await ssoComplete(key);
            if (data && data.Status === true && data.Data) {
                const jwt = data.Data.Token;
                const userRole = data.Data.UserRole;
                const userId = data.Data.ID;
                const loginId = data.Data.LoginID;

                // Store interactions in Local Storage for persistence (scoped to document)
                DocStorage.setItem('token', jwt);
                DocStorage.setItem(this.USER_ROLE_KEY, JSON.stringify(userRole));
                DocStorage.setItem('userId', userId);
                if (loginId !== undefined && loginId !== null) {
                    DocStorage.setItem(this.LOGIN_ID_KEY, String(loginId));
                }
                DocStorage.setItem('tokenLastUpdated', new Date().toISOString());

                // Sync with StoreService and persist
                const store = StoreService.getInstance();
                store.clearStorage();
                store.jwt = jwt;
                store.UserRole = userRole;
                store.userId = userId;
                if (loginId !== undefined && loginId !== null) {
                    store.loginId = loginId;
                }
                store.saveToStorage();

                if (loginId !== undefined && loginId !== null) {
                    this.startPingTimer(loginId, jwt);
                }

                return {
                    success: true,
                    data: {
                        token: jwt,
                        userRole: userRole,
                        userId: userId,
                        loginId: loginId,
                        raw: data.Data
                    }
                };
            } else {
                return {
                    success: false,
                    message: (data && data.Message) || 'SSO complete failed'
                };
            }
        } catch (error) {
            console.error('Error during ssoComplete:', error);
            return {
                success: false,
                message: 'Connection Lost'
            };
        }
    }

    static async logout(): Promise<void> {
        this.stopPingTimer();
        const token = DocStorage.getItem('token') || StoreService.getInstance().jwt;
        const loginId = DocStorage.getItem(this.LOGIN_ID_KEY) || StoreService.getInstance().loginId;

        try {
            if (token) {
                await logoutUser(loginId, token);
            }
        } catch (error) {
            console.error('Error during logout API call:', error);
        } finally {
            DocStorage.removeItem('token');
            DocStorage.removeItem(this.USER_ROLE_KEY);
            DocStorage.removeItem('userId');
            DocStorage.removeItem(this.LOGIN_ID_KEY);
            DocStorage.removeItem('tokenLastUpdated');

            const store = StoreService.getInstance();
            store.clearStorage();
            console.log("Logged out");
        }
    }
}

