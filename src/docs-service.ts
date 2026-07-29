import { google, docs_v1, drive_v3 } from "googleapis";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface DocParagraph {
  startIndex: number;
  endIndex: number;
  style: string;
  bulleted: boolean;
  text: string;
}

export interface DocOutline {
  documentId: string;
  title: string;
  revisionId: string;
  paragraphs: DocParagraph[];
  tableCount: number;
  indexHint: string;
}

export interface DocComment {
  id: string;
  author: string;
  content: string;
  quotedText: string | null;
  resolved: boolean;
  createdTime: string;
  replies: { author: string; content: string; createdTime: string }[];
}

export interface DocEditResult {
  documentId: string;
  revisionId: string;
  applied: number;
  detail: string;
}

// Named styles the Docs API accepts for updateParagraphStyle
export const NAMED_STYLES = [
  "NORMAL_TEXT",
  "TITLE",
  "SUBTITLE",
  "HEADING_1",
  "HEADING_2",
  "HEADING_3",
  "HEADING_4",
  "HEADING_5",
  "HEADING_6",
] as const;

export type NamedStyle = (typeof NAMED_STYLES)[number];

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function paragraphText(el: docs_v1.Schema$StructuralElement): string {
  const runs = el.paragraph?.elements ?? [];
  return runs.map((r) => r.textRun?.content ?? "").join("");
}

// ---------------------------------------------------------------------------
// Docs Service — one instance per access token (per request)
// ---------------------------------------------------------------------------

export class DocsService {
  private docs: docs_v1.Docs;
  private drive: drive_v3.Drive;

  constructor(accessToken: string) {
    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    this.docs = google.docs({ version: "v1", auth });
    this.drive = google.drive({ version: "v3", auth });
  }

  // -----------------------------------------------------------------------
  // read — returns paragraph-level structure WITH indices, which is what
  // every positional edit below depends on.
  // -----------------------------------------------------------------------

  async readDocument(documentId: string): Promise<DocOutline> {
    const res = await this.docs.documents.get({ documentId });
    const doc = res.data;

    const paragraphs: DocParagraph[] = [];
    let tableCount = 0;

    for (const el of doc.body?.content ?? []) {
      if (el.table) {
        tableCount++;
        continue;
      }
      if (!el.paragraph) continue;

      const text = paragraphText(el);
      paragraphs.push({
        startIndex: el.startIndex ?? 0,
        endIndex: el.endIndex ?? 0,
        style: el.paragraph.paragraphStyle?.namedStyleType ?? "NORMAL_TEXT",
        bulleted: !!el.paragraph.bullet,
        text,
      });
    }

    return {
      documentId,
      title: doc.title ?? "",
      revisionId: doc.revisionId ?? "",
      paragraphs,
      tableCount,
      indexHint:
        "startIndex/endIndex are character offsets into the document. To target text inside a paragraph, add the offset of that text within paragraph.text to paragraph.startIndex. Indices shift after every edit — re-read the document before issuing a second positional edit.",
    };
  }

  // -----------------------------------------------------------------------
  // batchUpdate wrapper
  // -----------------------------------------------------------------------

  private async batch(
    documentId: string,
    requests: docs_v1.Schema$Request[],
    detail: string
  ): Promise<DocEditResult> {
    const res = await this.docs.documents.batchUpdate({
      documentId,
      requestBody: { requests },
    });

    return {
      documentId,
      revisionId: res.data.writeControl?.requiredRevisionId ?? "",
      applied: requests.length,
      detail,
    };
  }

  // -----------------------------------------------------------------------
  // replace text — index-free, safest edit
  // -----------------------------------------------------------------------

  async replaceText(
    documentId: string,
    find: string,
    replace: string,
    matchCase: boolean = true
  ): Promise<DocEditResult & { occurrencesChanged: number }> {
    const res = await this.docs.documents.batchUpdate({
      documentId,
      requestBody: {
        requests: [
          {
            replaceAllText: {
              containsText: { text: find, matchCase },
              replaceText: replace,
            },
          },
        ],
      },
    });

    const changed =
      res.data.replies?.[0]?.replaceAllText?.occurrencesChanged ?? 0;

    return {
      documentId,
      revisionId: res.data.writeControl?.requiredRevisionId ?? "",
      applied: 1,
      occurrencesChanged: changed,
      detail:
        changed === 0
          ? `No occurrences of "${find}" found — nothing changed.`
          : `Replaced ${changed} occurrence(s).`,
    };
  }

  // -----------------------------------------------------------------------
  // insert / delete at index
  // -----------------------------------------------------------------------

  async insertText(
    documentId: string,
    index: number,
    text: string
  ): Promise<DocEditResult> {
    return this.batch(
      documentId,
      [{ insertText: { location: { index }, text } }],
      `Inserted ${text.length} characters at index ${index}.`
    );
  }

  async appendText(documentId: string, text: string): Promise<DocEditResult> {
    return this.batch(
      documentId,
      [{ insertText: { endOfSegmentLocation: {}, text } }],
      `Appended ${text.length} characters to the end of the document.`
    );
  }

  async deleteRange(
    documentId: string,
    startIndex: number,
    endIndex: number
  ): Promise<DocEditResult> {
    return this.batch(
      documentId,
      [{ deleteContentRange: { range: { startIndex, endIndex } } }],
      `Deleted characters ${startIndex}–${endIndex}.`
    );
  }

  // -----------------------------------------------------------------------
  // formatting
  // -----------------------------------------------------------------------

  async formatText(
    documentId: string,
    startIndex: number,
    endIndex: number,
    opts: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      strikethrough?: boolean;
      fontSize?: number;
      linkUrl?: string;
    }
  ): Promise<DocEditResult> {
    const textStyle: docs_v1.Schema$TextStyle = {};
    const fields: string[] = [];

    if (opts.bold !== undefined) {
      textStyle.bold = opts.bold;
      fields.push("bold");
    }
    if (opts.italic !== undefined) {
      textStyle.italic = opts.italic;
      fields.push("italic");
    }
    if (opts.underline !== undefined) {
      textStyle.underline = opts.underline;
      fields.push("underline");
    }
    if (opts.strikethrough !== undefined) {
      textStyle.strikethrough = opts.strikethrough;
      fields.push("strikethrough");
    }
    if (opts.fontSize !== undefined) {
      textStyle.fontSize = { magnitude: opts.fontSize, unit: "PT" };
      fields.push("fontSize");
    }
    if (opts.linkUrl !== undefined) {
      textStyle.link = { url: opts.linkUrl };
      fields.push("link");
    }

    if (fields.length === 0) {
      throw new Error(
        "No formatting options supplied. Pass at least one of bold, italic, underline, strikethrough, font_size, link_url."
      );
    }

    return this.batch(
      documentId,
      [
        {
          updateTextStyle: {
            range: { startIndex, endIndex },
            textStyle,
            fields: fields.join(","),
          },
        },
      ],
      `Applied ${fields.join(", ")} to characters ${startIndex}–${endIndex}.`
    );
  }

  async setParagraphStyle(
    documentId: string,
    startIndex: number,
    endIndex: number,
    namedStyle: NamedStyle
  ): Promise<DocEditResult> {
    return this.batch(
      documentId,
      [
        {
          updateParagraphStyle: {
            range: { startIndex, endIndex },
            paragraphStyle: { namedStyleType: namedStyle },
            fields: "namedStyleType",
          },
        },
      ],
      `Set paragraph style ${namedStyle} on characters ${startIndex}–${endIndex}.`
    );
  }

  async setBullets(
    documentId: string,
    startIndex: number,
    endIndex: number,
    numbered: boolean
  ): Promise<DocEditResult> {
    return this.batch(
      documentId,
      [
        {
          createParagraphBullets: {
            range: { startIndex, endIndex },
            bulletPreset: numbered
              ? "NUMBERED_DECIMAL_ALPHA_ROMAN"
              : "BULLET_DISC_CIRCLE_SQUARE",
          },
        },
      ],
      `Applied ${numbered ? "numbered" : "bulleted"} list to characters ${startIndex}–${endIndex}.`
    );
  }

  // -----------------------------------------------------------------------
  // tables & breaks
  // -----------------------------------------------------------------------

  async insertTable(
    documentId: string,
    index: number,
    rows: number,
    columns: number
  ): Promise<DocEditResult> {
    return this.batch(
      documentId,
      [{ insertTable: { location: { index }, rows, columns } }],
      `Inserted a ${rows}×${columns} table at index ${index}.`
    );
  }

  async insertPageBreak(
    documentId: string,
    index: number
  ): Promise<DocEditResult> {
    return this.batch(
      documentId,
      [{ insertPageBreak: { location: { index } } }],
      `Inserted a page break at index ${index}.`
    );
  }

  // -----------------------------------------------------------------------
  // comments — Drive API, not Docs API. The Docs API has no comment support.
  // -----------------------------------------------------------------------

  async addComment(
    documentId: string,
    content: string,
    quotedText?: string
  ): Promise<{ id: string; content: string; quotedText: string | null }> {
    const res = await this.drive.comments.create({
      fileId: documentId,
      fields: "id, content, quotedFileContent",
      requestBody: {
        content,
        ...(quotedText
          ? { quotedFileContent: { mimeType: "text/plain", value: quotedText } }
          : {}),
      },
    });

    return {
      id: res.data.id ?? "",
      content: res.data.content ?? content,
      quotedText: res.data.quotedFileContent?.value ?? null,
    };
  }

  async listComments(
    documentId: string,
    includeResolved: boolean = false
  ): Promise<DocComment[]> {
    const res = await this.drive.comments.list({
      fileId: documentId,
      fields:
        "comments(id, author(displayName), content, quotedFileContent(value), resolved, createdTime, replies(author(displayName), content, createdTime))",
      pageSize: 100,
      includeDeleted: false,
    });

    return (res.data.comments ?? [])
      .filter((c) => includeResolved || !c.resolved)
      .map((c) => ({
        id: c.id ?? "",
        author: c.author?.displayName ?? "unknown",
        content: c.content ?? "",
        quotedText: c.quotedFileContent?.value ?? null,
        resolved: c.resolved ?? false,
        createdTime: c.createdTime ?? "",
        replies: (c.replies ?? []).map((r) => ({
          author: r.author?.displayName ?? "unknown",
          content: r.content ?? "",
          createdTime: r.createdTime ?? "",
        })),
      }));
  }

  async replyToComment(
    documentId: string,
    commentId: string,
    content: string
  ): Promise<{ id: string; content: string }> {
    const res = await this.drive.replies.create({
      fileId: documentId,
      commentId,
      fields: "id, content",
      requestBody: { content },
    });

    return { id: res.data.id ?? "", content: res.data.content ?? content };
  }

  async resolveComment(
    documentId: string,
    commentId: string
  ): Promise<{ id: string; resolved: boolean }> {
    // Drive resolves a comment by posting a reply with action=resolve
    await this.drive.replies.create({
      fileId: documentId,
      commentId,
      fields: "id",
      requestBody: { action: "resolve" },
    });

    return { id: commentId, resolved: true };
  }
}
