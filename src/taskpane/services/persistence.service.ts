
export class PersistenceService {
    private static readonly STORAGE_KEY = 'office_addin_state';
    private static readonly EXPIRY_MS = 60 * 60 * 1000; // 1 hour
    private static readonly SECRET_KEY_STUB = 'TrialAssure-LINK-AI-Persistence'; // In a real app, this should be more secure

    /**
     * Saves the current application state to localStorage after encryption.
     */
    static async saveState(state: any): Promise<void> {
        try {
            const dataToSave = {
                ...state,
                timestamp: Date.now()
            };
            const jsonString = JSON.stringify(dataToSave);
            const encryptedData = await this.encrypt(jsonString);
            localStorage.setItem(this.STORAGE_KEY, encryptedData);
        } catch (error) {
            console.error("PersistenceService: Failed to save state", error);
        }
    }

    /**
     * Loads, decrypts, and validates the state from localStorage.
     * Returns null if state is invalid, expired, or doesn't exist.
     */
    static async loadState(): Promise<any | null> {
        try {
            const encryptedData = localStorage.getItem(this.STORAGE_KEY);
            if (!encryptedData) return null;

            const jsonString = await this.decrypt(encryptedData);
            if (!jsonString) {
                this.clearState();
                return null;
            }

            const state = JSON.parse(jsonString);
            if (this.isExpired(state.timestamp)) {
                console.log("PersistenceService: State expired");
                this.clearState();
                return null;
            }

            return state;
        } catch (error) {
            console.error("PersistenceService: Failed to load state", error);
            this.clearState();
            return null;
        }
    }

    /**
     * Clears the stored state.
     */
    static clearState(): void {
        localStorage.removeItem(this.STORAGE_KEY);
    }

    /**
     * Checks if the state has expired (older than 1 hour).
     */
    private static isExpired(timestamp: number): boolean {
        if (!timestamp) return true;
        const currentTime = Date.now();
        return (currentTime - timestamp) > this.EXPIRY_MS;
    }

    /**
     * Encryption logic using Web Crypto API (AES-GCM)
     */
    private static async encrypt(text: string): Promise<string> {
        try {
            const encoder = new TextEncoder();
            const data = encoder.encode(text);

            const key = await this.deriveKey();
            const iv = crypto.getRandomValues(new Uint8Array(12));
            const encrypted = await crypto.subtle.encrypt(
                { name: 'AES-GCM', iv },
                key,
                data
            );

            // Combine IV and Encrypted data
            const combined = new Uint8Array(iv.length + encrypted.byteLength);
            combined.set(iv);
            combined.set(new Uint8Array(encrypted), iv.length);

            // Convert to Base64 (safely for older TS/environments)
            let binary = '';
            for (let i = 0; i < combined.byteLength; i++) {
                binary += String.fromCharCode(combined[i]);
            }
            return btoa(binary);
        } catch (error) {
            console.error("Encryption error", error);
            throw error;
        }
    }

    /**
     * Decryption logic using Web Crypto API (AES-GCM)
     */
    private static async decrypt(base64Data: string): Promise<string | null> {
        try {
            const combined = new Uint8Array(
                atob(base64Data)
                    .split('')
                    .map(char => char.charCodeAt(0))
            );

            const iv = combined.slice(0, 12);
            const data = combined.slice(12);

            const key = await this.deriveKey();
            const decrypted = await crypto.subtle.decrypt(
                { name: 'AES-GCM', iv },
                key,
                data
            );

            const decoder = new TextDecoder();
            return decoder.decode(decrypted);
        } catch (error) {
            console.error("Decryption error", error);
            return null;
        }
    }

    /**
     * Derives a CryptoKey from the static secret stub.
     * In a production environment, this should be more robust.
     */
    private static async deriveKey(): Promise<CryptoKey> {
        const encoder = new TextEncoder();
        const keyMaterial = await crypto.subtle.importKey(
            'raw',
            encoder.encode(this.SECRET_KEY_STUB),
            { name: 'PBKDF2' },
            false,
            ['deriveKey']
        );

        return await crypto.subtle.deriveKey(
            {
                name: 'PBKDF2',
                salt: encoder.encode('TrialAssure-Salt'),
                iterations: 100000,
                hash: 'SHA-256'
            },
            keyMaterial,
            { name: 'AES-GCM', length: 256 },
            false,
            ['encrypt', 'decrypt']
        );
    }
}
