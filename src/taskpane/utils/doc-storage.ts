/**
 * DocStorage – a localStorage wrapper that scopes every key to the current
 * document ID, so that two add-in instances opened against different documents
 * never share credentials or settings.
 *
 * Usage:  DocStorage.getItem('token')
 *         DocStorage.setItem('token', jwt)
 *         DocStorage.removeItem('token')
 *         DocStorage.clearAll()   ← clears only THIS document's keys
 *
 * The document ID is read lazily from StoreService each time a method is
 * called, so it is always up-to-date even if the store is initialised after
 * the first import.
 */

let _docId: string = '';

/** Called once in taskpane.ts after documentID is known. */
export function setDocStorageId(id: string): void {
    _docId = id || 'default';
}

function prefix(key: string): string {
    return `${_docId || 'default'}__${key}`;
}

export const DocStorage = {
    getItem(key: string): string | null {
        return localStorage.getItem(prefix(key));
    },

    setItem(key: string, value: string): void {
        localStorage.setItem(prefix(key), value);
    },

    removeItem(key: string): void {
        localStorage.removeItem(prefix(key));
    },

    /** Remove every localStorage entry that belongs to the current document. */
    clearAll(): void {
        const pfx = prefix('');
        const toDelete: string[] = [];
        for (let i = 0; i < localStorage.length; i++) {
            const k = localStorage.key(i);
            if (k && k.startsWith(pfx)) {
                toDelete.push(k);
            }
        }
        toDelete.forEach(k => localStorage.removeItem(k));
    }
};
