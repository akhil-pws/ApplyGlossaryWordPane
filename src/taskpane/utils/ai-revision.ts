/**
 * Attributes content inserted by the add-in to "LINK AI Assistant" in Word's
 * Track Changes / Revisions pane instead of the signed-in Word user.
 *
 * Office.js has no API to set a revision's author, so the approach is:
 *   1. Temporarily switch Track Changes off (so Word does not record the
 *      add-in's own edits as the current user).
 *   2. Insert content as usual.
 *   3. Round-trip each inserted paragraph's content through OOXML, wrapping
 *      every run in <w:ins w:author="LINK AI Assistant">, and write it back.
 *   4. Restore the original Track Changes mode, so later manual edits are
 *      attributed to the real user as normal.
 *
 * If Track Changes was off to begin with, nothing is rewritten and the
 * insertion behaves exactly as before.
 *
 * The author name comes from the document's "ChangeTracker" custom property
 * when it is set; otherwise it defaults to "LINK AI Assistant".
 */

export const AI_REVISION_AUTHOR = "LINK AI Assistant";
export const AI_REVISION_AUTHOR_PROPERTY = "ChangeTracker";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

type TrackingMode = Word.Document["changeTrackingMode"];

let revisionId = 900000 + Math.floor(Math.random() * 90000);

export class AIRevisionSession {
  private constructor(
    private readonly context: Word.RequestContext,
    private readonly previousMode: TrackingMode | null,
    private readonly author: string = AI_REVISION_AUTHOR
  ) { }

  /** True when Track Changes was on and inserted content will be re-attributed. */
  get active(): boolean {
    return this.previousMode !== null;
  }

  /** Turns Track Changes off (remembering the previous mode) for the duration of the insertion. */
  static async begin(context: Word.RequestContext): Promise<AIRevisionSession> {
    if (!Office.context.requirements.isSetSupported("WordApi", "1.4")) {
      return new AIRevisionSession(context, null);
    }
    try {
      const doc = context.document;
      doc.load("changeTrackingMode");
      await context.sync();

      const mode = doc.changeTrackingMode;
      if (!mode || mode === Word.ChangeTrackingMode.off) {
        return new AIRevisionSession(context, null);
      }

      const author = await getRevisionAuthor(context);

      doc.changeTrackingMode = Word.ChangeTrackingMode.off;
      await context.sync();
      return new AIRevisionSession(context, mode, author);
    } catch (e) {
      console.warn("AI revision: unable to read/suspend change tracking", e);
      return new AIRevisionSession(context, null);
    }
  }

  /**
   * Re-inserts the content of `range` as a tracked insertion authored by
   * "LINK AI Assistant". Paragraph/table/style formatting is preserved because
   * only the inline content of each paragraph is replaced.
   *
   * Returns a range covering the (re-inserted) content; use it for anything
   * applied afterwards, such as bookmarks.
   */
  async markAsAI(range: Word.Range): Promise<Word.Range> {
    if (!this.active || !range) return range;
    const context = this.context;

    try {
      const paragraphs = range.paragraphs;
      paragraphs.load("items");
      await context.sync();

      // Only touch the part of each paragraph that lies inside the inserted
      // range, so pre-existing user text in the same paragraph is untouched.
      const pieces = paragraphs.items.map(p =>
        p.getRange(Word.RangeLocation.content).intersectWithOrNullObject(range)
      );
      pieces.forEach(r => r.load("isNullObject"));
      await context.sync();

      const targets = pieces.filter(r => !r.isNullObject);
      const ooxmls = targets.map(r => r.getOoxml());
      await context.sync();

      const date = new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
      const replaced: Word.Range[] = [];
      targets.forEach((r, i) => {
        const transformed = wrapRunsAsInsertion(ooxmls[i].value, date, this.author);
        if (transformed) {
          replaced.push(r.insertOoxml(transformed, Word.InsertLocation.replace));
        }
      });
      await context.sync();

      if (!replaced.length) return range;
      const first = replaced[0];
      const last = replaced[replaced.length - 1];
      try {
        const covering = range.expandTo(first).expandTo(last);
        await context.sync();
        return covering;
      } catch {
        // Original range no longer resolvable; fall back to the re-inserted span.
        return first.expandTo(last);
      }
    } catch (e) {
      // Never fail the insertion itself because attribution failed.
      console.warn("AI revision: unable to attribute inserted content", e);
      return range;
    }
  }

  /** Restores the Track Changes mode that was active before `begin`. */
  async end(): Promise<void> {
    if (!this.active) return;
    try {
      this.context.document.changeTrackingMode = this.previousMode!;
      await this.context.sync();
    } catch (e) {
      console.warn("AI revision: unable to restore change tracking", e);
    }
  }
}

/**
 * Reads the "ChangeTracker" custom document property; falls back to
 * "LINK AI Assistant" when it is missing, empty or cannot be read.
 */
async function getRevisionAuthor(context: Word.RequestContext): Promise<string> {
  if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
    return AI_REVISION_AUTHOR;
  }
  try {
    const prop = context.document.properties.customProperties
      .getItemOrNullObject(AI_REVISION_AUTHOR_PROPERTY);
    prop.load("value");
    await context.sync();

    const value = prop.isNullObject ? "" : String(prop.value ?? "").trim();
    return value || AI_REVISION_AUTHOR;
  } catch (e) {
    console.warn("AI revision: unable to read ChangeTracker property", e);
    return AI_REVISION_AUTHOR;
  }
}

/**
 * Wraps every <w:r> in the package's document body with a
 * <w:ins w:author="..."> element. Returns null when there is nothing to wrap
 * (e.g. an empty paragraph).
 */
function wrapRunsAsInsertion(ooxml: string, date: string, author: string): string | null {
  const xml = new DOMParser().parseFromString(ooxml, "application/xml");
  if (xml.getElementsByTagName("parsererror").length) return null;

  const bodies = xml.getElementsByTagNameNS(W_NS, "body");
  if (!bodies.length) return null;

  const runs = Array.from(bodies[0].getElementsByTagNameNS(W_NS, "r"));
  let wrapped = 0;

  for (const run of runs) {
    if (hasRevisionAncestor(run)) continue;

    const ins = xml.createElementNS(W_NS, "w:ins");
    ins.setAttributeNS(W_NS, "w:id", String(revisionId++));
    ins.setAttributeNS(W_NS, "w:author", author);
    ins.setAttributeNS(W_NS, "w:date", date);

    run.parentNode!.insertBefore(ins, run);
    ins.appendChild(run);
    wrapped++;
  }

  return wrapped ? new XMLSerializer().serializeToString(xml) : null;
}

function hasRevisionAncestor(node: Element): boolean {
  for (let p = node.parentElement; p; p = p.parentElement) {
    if (p.namespaceURI === W_NS && ["ins", "del", "moveFrom", "moveTo"].includes(p.localName)) {
      return true;
    }
  }
  return false;
}
