import { GlossaryTerm } from "../models/glossary.model";
import { CONFIG } from "../utils/config";

export class GlossaryService {

    static async fetchGlossary(): Promise<GlossaryTerm[]> {
        console.log(`Fetching glossary from ${CONFIG.dataUrl}`);
        // Logic from fetchGlossary in taskpane.ts
        // Placeholder return
        return [];
    }

    static async applyGlossary(terms: GlossaryTerm[]): Promise<void> {
        await Word.run(async (context) => {
            const body = context.document.body;
            terms.forEach(term => {
                // Logic to search and replace or highlight
                // This is a complex logic in taskpane.ts, needs careful porting
            });
            await context.sync();
        });
    }
}
