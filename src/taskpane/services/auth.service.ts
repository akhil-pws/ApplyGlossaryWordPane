import { CONFIG } from "../utils/config";
import { UserProfile } from "../models/user.model";
import { loginUser } from "../draft/draft.api";
import { StoreService } from "./store.service";
import { DocStorage } from "../utils/doc-storage";

export class AuthService {
    private static readonly TOKEN_KEY = 'user_token';
    private static readonly USER_ROLE_KEY = 'userRole';
    private static readonly STYLE_KEY = 'tableStyle';
    private static readonly PALETTE_KEY = 'colorPallete';

    static getStoredToken(): string | null {
        return DocStorage.getItem('token'); 
    }

    static restoreSession(): any {
        const sessionToken = DocStorage.getItem('token');
        if (sessionToken) {
            return {
                jwt: sessionToken,
                userRole: JSON.parse(DocStorage.getItem(this.USER_ROLE_KEY) || '{}'),
                tableStyle: DocStorage.getItem(this.STYLE_KEY),
                colorPallete: JSON.parse(DocStorage.getItem(this.PALETTE_KEY) || 'null'),
                userId: DocStorage.getItem('userId')
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

                    // Store interactions in Local Storage for persistence (scoped to document)
                    DocStorage.setItem('token', jwt);
                    DocStorage.setItem(this.USER_ROLE_KEY, JSON.stringify(userRole));
                    DocStorage.setItem('userId', userId);

                    // Sync with StoreService and persist
                    const store = StoreService.getInstance();
                    store.clearStorage();
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
        DocStorage.removeItem('token');
        DocStorage.removeItem(this.USER_ROLE_KEY);
        DocStorage.removeItem('userId');
        
        const store = StoreService.getInstance();
        store.clearStorage();
        console.log("Logged out");
    }
}
