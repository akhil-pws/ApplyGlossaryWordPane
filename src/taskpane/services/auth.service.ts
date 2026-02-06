import { CONFIG } from "../utils/config";
import { UserProfile } from "../models/user.model";
import { loginUser } from "../draft/draft.api";

export class AuthService {
    private static readonly TOKEN_KEY = 'user_token';
    private static readonly USER_ROLE_KEY = 'userRole';
    private static readonly STYLE_KEY = 'tableStyle';
    private static readonly PALETTE_KEY = 'colorPallete';

    static getStoredToken(): string | null {
        return sessionStorage.getItem('token'); // Legacy code used sessionStorage for token
    }

    static restoreSession(): any {
        const sessionToken = sessionStorage.getItem('token');
        if (sessionToken) {
            return {
                jwt: sessionToken,
                userRole: JSON.parse(sessionStorage.getItem(this.USER_ROLE_KEY) || '{}'),
                tableStyle: sessionStorage.getItem(this.STYLE_KEY),
                colorPallete: JSON.parse(sessionStorage.getItem(this.PALETTE_KEY) || 'null'),
                userId: sessionStorage.getItem('userId')
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

                    // Store interactions in Session
                    sessionStorage.setItem('token', jwt);
                    sessionStorage.setItem(this.USER_ROLE_KEY, JSON.stringify(userRole));
                    sessionStorage.setItem('userId', userId);

                    // Note: tableStyle and colorPallete are typically stored sequentially or retrieved from profile
                    // For now, we just ensure session is clean or updated as per legacy flow which checked sessionStorage items?
                    // Legacy flow: 
                    // const style = sessionStorage.getItem('tableStyle');
                    // const localPallete = sessionStorage.getItem('colorPallete');
                    // It seems legacy flow READS from session storage if available, it doesn't SET them from login response?
                    // Actually, it seemed to just re-read them.

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
        sessionStorage.removeItem('token');
        sessionStorage.removeItem(this.USER_ROLE_KEY);
        // Add other keys
        console.log("Logged out");
    }
}
