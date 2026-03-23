import { CONFIG } from "../utils/config";
import { UserProfile } from "../models/user.model";
import { loginUser } from "../draft/draft.api";
import { StoreService } from "./store.service";

export class AuthService {
    private static readonly TOKEN_KEY = 'user_token';
    private static readonly USER_ROLE_KEY = 'userRole';
    private static readonly STYLE_KEY = 'tableStyle';
    private static readonly PALETTE_KEY = 'colorPallete';

    static getStoredToken(): string | null {
        return localStorage.getItem('token'); 
    }

    static restoreSession(): any {
        const sessionToken = localStorage.getItem('token');
        if (sessionToken) {
            return {
                jwt: sessionToken,
                userRole: JSON.parse(localStorage.getItem(this.USER_ROLE_KEY) || '{}'),
                tableStyle: localStorage.getItem(this.STYLE_KEY),
                colorPallete: JSON.parse(localStorage.getItem(this.PALETTE_KEY) || 'null'),
                userId: localStorage.getItem('userId')
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

                    // Store interactions in Local Storage for persistence
                    localStorage.setItem('token', jwt);
                    localStorage.setItem(this.USER_ROLE_KEY, JSON.stringify(userRole));
                    localStorage.setItem('userId', userId);

                    // Sync with StoreService and persist
                    const store = StoreService.getInstance();
                    store.jwt = jwt;
                    store.UserRole = userRole;
                    store.userId = userId;
                    store.saveToStorage();

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
        localStorage.removeItem('token');
        localStorage.removeItem(this.USER_ROLE_KEY);
        localStorage.removeItem('userId');
        
        const store = StoreService.getInstance();
        store.clearStorage();
        console.log("Logged out");
    }
}
