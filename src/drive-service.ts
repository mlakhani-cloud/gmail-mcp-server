import { google, drive_v3 } from "googleapis";
import { Readable } from "node:stream";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface DriveFileSummary {
  id: string;
  name: string;
  mimeType: string;
  modifiedTime: string;
  size: string | null;
  owners: string[];
  webViewLink: string | null;
  parents: string[];
  starred: boolean;
  trashed: boolean;
}

export interface DriveFileContent extends DriveFileSummary {
  content: string;
  contentTruncated: boolean;
  exportedAs: string | null;
}

export interface DriveWriteResult {
  id: string;
  name: string;
  mimeType: string;
  webViewLink: string | null;
}

// ---------------------------------------------------------------------------
// Google-native MIME types and how to export them as text
// ---------------------------------------------------------------------------

const GOOGLE_EXPORT_MAP: Record<string, string> = {
  "application/vnd.google-apps.document": "text/plain",
  "application/vnd.google-apps.spreadsheet": "text/csv",
  "application/vnd.google-apps.presentation": "text/plain",
  "application/vnd.google-apps.script": "application/vnd.google-apps.script+json",
};

// Plain-text-ish types we can safely read directly
const TEXT_LIKE = [
  "text/",
  "application/json",
  "application/xml",
  "application/javascript",
  "application/x-yaml",
];

const MAX_CONTENT_CHARS = 100_000;

const FILE_FIELDS =
  "id, name, mimeType, modifiedTime, size, owners(emailAddress), webViewLink, parents, starred, trashed";

function toSummary(f: drive_v3.Schema$File): DriveFileSummary {
  return {
    id: f.id ?? "",
    name: f.name ?? "(unnamed)",
    mimeType: f.mimeType ?? "",
    modifiedTime: f.modifiedTime ?? "",
    size: f.size ?? null,
    owners: (f.owners ?? [])
      .map((o) => o.emailAddress ?? "")
      .filter(Boolean),
    webViewLink: f.webViewLink ?? null,
    parents: f.parents ?? [],
    starred: f.starred ?? false,
    trashed: f.trashed ?? false,
  };
}

// ---------------------------------------------------------------------------
// Drive Service — one instance per access token (per request)
// ---------------------------------------------------------------------------

export class DriveService {
  private drive: drive_v3.Drive;

  constructor(accessToken: string) {
    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    this.drive = google.drive({ version: "v3", auth });
  }

  // -----------------------------------------------------------------------
  // search
  // -----------------------------------------------------------------------

  async searchFiles(
    query?: string,
    maxResults: number = 20,
    includeTrashed: boolean = false
  ): Promise<DriveFileSummary[]> {
    // If the caller passed bare words rather than Drive query syntax, treat it
    // as a full-text search. Drive query syntax always contains an operator.
    const looksLikeQuerySyntax =
      !!query &&
      /(\bcontains\b|\bin parents\b|mimeType\s*=|name\s*=|modifiedTime\s*[<>]|trashed\s*=|starred\s*=|\bowners\b|\bsharedWithMe\b)/i.test(
        query
      );

    let q: string | undefined;
    if (query && looksLikeQuerySyntax) {
      q = query;
    } else if (query) {
      q = `fullText contains '${query.replace(/'/g, "\\'")}'`;
    }

    if (!includeTrashed) {
      q = q ? `(${q}) and trashed = false` : "trashed = false";
    }

    const res = await this.drive.files.list({
      q,
      pageSize: Math.min(Math.max(maxResults, 1), 100),
      fields: `files(${FILE_FIELDS})`,
      orderBy: "modifiedTime desc",
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      corpora: "allDrives",
    });

    return (res.data.files ?? []).map(toSummary);
  }

  // -----------------------------------------------------------------------
  // read
  // -----------------------------------------------------------------------

  async getFile(fileId: string): Promise<DriveFileContent> {
    const meta = await this.drive.files.get({
      fileId,
      fields: FILE_FIELDS,
      supportsAllDrives: true,
    });

    const summary = toSummary(meta.data);
    const mimeType = summary.mimeType;

    let content = "";
    let exportedAs: string | null = null;

    const exportMime = GOOGLE_EXPORT_MAP[mimeType];

    if (exportMime) {
      const res = await this.drive.files.export(
        { fileId, mimeType: exportMime },
        { responseType: "text" }
      );
      content = typeof res.data === "string" ? res.data : String(res.data);
      exportedAs = exportMime;
    } else if (mimeType === "application/vnd.google-apps.folder") {
      const children = await this.drive.files.list({
        q: `'${fileId}' in parents and trashed = false`,
        pageSize: 100,
        fields: `files(${FILE_FIELDS})`,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
      });
      content = JSON.stringify(
        (children.data.files ?? []).map(toSummary),
        null,
        2
      );
      exportedAs = "folder-listing";
    } else if (TEXT_LIKE.some((prefix) => mimeType.startsWith(prefix))) {
      const res = await this.drive.files.get(
        { fileId, alt: "media", supportsAllDrives: true },
        { responseType: "text" }
      );
      content = typeof res.data === "string" ? res.data : String(res.data);
    } else {
      content = `[Binary file — ${mimeType}. Not converted to text. Open it at ${summary.webViewLink ?? "Drive"}.]`;
    }

    const truncated = content.length > MAX_CONTENT_CHARS;
    if (truncated) content = content.slice(0, MAX_CONTENT_CHARS);

    return {
      ...summary,
      content,
      contentTruncated: truncated,
      exportedAs,
    };
  }

  // -----------------------------------------------------------------------
  // create
  // -----------------------------------------------------------------------

  async createFile(
    name: string,
    content: string,
    options: {
      mimeType?: string;
      folderId?: string;
      asGoogleDoc?: boolean;
    } = {}
  ): Promise<DriveWriteResult> {
    const sourceMime = options.mimeType ?? "text/plain";

    // Converting to a Google Doc/Sheet means uploading text and asking Drive
    // to convert it into the native type.
    let targetMime: string | undefined;
    if (options.asGoogleDoc) {
      targetMime =
        sourceMime === "text/csv"
          ? "application/vnd.google-apps.spreadsheet"
          : "application/vnd.google-apps.document";
    }

    const res = await this.drive.files.create({
      requestBody: {
        name,
        ...(targetMime ? { mimeType: targetMime } : {}),
        ...(options.folderId ? { parents: [options.folderId] } : {}),
      },
      media: {
        mimeType: sourceMime,
        body: Readable.from([content]),
      },
      fields: "id, name, mimeType, webViewLink",
      supportsAllDrives: true,
    });

    return {
      id: res.data.id ?? "",
      name: res.data.name ?? name,
      mimeType: res.data.mimeType ?? sourceMime,
      webViewLink: res.data.webViewLink ?? null,
    };
  }

  async createFolder(
    name: string,
    parentId?: string
  ): Promise<DriveWriteResult> {
    const res = await this.drive.files.create({
      requestBody: {
        name,
        mimeType: "application/vnd.google-apps.folder",
        ...(parentId ? { parents: [parentId] } : {}),
      },
      fields: "id, name, mimeType, webViewLink",
      supportsAllDrives: true,
    });

    return {
      id: res.data.id ?? "",
      name: res.data.name ?? name,
      mimeType: "application/vnd.google-apps.folder",
      webViewLink: res.data.webViewLink ?? null,
    };
  }

  // -----------------------------------------------------------------------
  // update
  // -----------------------------------------------------------------------

  async updateFileContent(
    fileId: string,
    content: string,
    mimeType: string = "text/plain"
  ): Promise<DriveWriteResult> {
    const res = await this.drive.files.update({
      fileId,
      media: {
        mimeType,
        body: Readable.from([content]),
      },
      fields: "id, name, mimeType, webViewLink",
      supportsAllDrives: true,
    });

    return {
      id: res.data.id ?? fileId,
      name: res.data.name ?? "",
      mimeType: res.data.mimeType ?? mimeType,
      webViewLink: res.data.webViewLink ?? null,
    };
  }

  async renameFile(fileId: string, name: string): Promise<DriveWriteResult> {
    const res = await this.drive.files.update({
      fileId,
      requestBody: { name },
      fields: "id, name, mimeType, webViewLink",
      supportsAllDrives: true,
    });

    return {
      id: res.data.id ?? fileId,
      name: res.data.name ?? name,
      mimeType: res.data.mimeType ?? "",
      webViewLink: res.data.webViewLink ?? null,
    };
  }

  async moveFile(fileId: string, folderId: string): Promise<DriveWriteResult> {
    const current = await this.drive.files.get({
      fileId,
      fields: "parents",
      supportsAllDrives: true,
    });
    const previousParents = (current.data.parents ?? []).join(",");

    const res = await this.drive.files.update({
      fileId,
      addParents: folderId,
      ...(previousParents ? { removeParents: previousParents } : {}),
      fields: "id, name, mimeType, webViewLink",
      supportsAllDrives: true,
    });

    return {
      id: res.data.id ?? fileId,
      name: res.data.name ?? "",
      mimeType: res.data.mimeType ?? "",
      webViewLink: res.data.webViewLink ?? null,
    };
  }

  // -----------------------------------------------------------------------
  // trash (deliberately not permanent delete)
  // -----------------------------------------------------------------------

  async trashFile(fileId: string): Promise<{ id: string; trashed: boolean; name: string }> {
    const res = await this.drive.files.update({
      fileId,
      requestBody: { trashed: true },
      fields: "id, name, trashed",
      supportsAllDrives: true,
    });

    return {
      id: res.data.id ?? fileId,
      name: res.data.name ?? "",
      trashed: res.data.trashed ?? true,
    };
  }
}
