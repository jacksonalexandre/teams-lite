import { getAccessToken, type TeamsConfig } from "./msal";

const GRAPH = "https://graph.microsoft.com/v1.0";

// =========== Throttling (HTTP 429) ===========
// One gate for the whole app: after a 429, every Graph call fails fast until
// the window passes, so polling stops hitting the API instead of piling on.

let throttledUntil = 0;
let throttleStreak = 0;

function isThrottled() {
  return Date.now() < throttledUntil;
}

function registerThrottle(res: Response) {
  const header = Number(res.headers.get("Retry-After"));
  throttleStreak++;
  const fallback = Math.min(30 * 2 ** (throttleStreak - 1), 300);
  throttledUntil = Date.now() + (Number.isFinite(header) && header > 0 ? header : fallback) * 1000;
}

async function authorizedFetch(cfg: TeamsConfig, url: string, init: RequestInit = {}) {
  if (isThrottled()) throw new Error("Graph 429: aguardando o fim do throttling");
  const token = await getAccessToken(cfg);
  const res = await fetch(url, {
    ...init,
    headers: { Authorization: `Bearer ${token}`, ...(init.headers || {}) },
  });
  if (res.status === 429) registerThrottle(res);
  else if (res.ok) throttleStreak = 0;
  return res;
}

async function graphFetch(cfg: TeamsConfig, path: string, init: RequestInit = {}) {
  const res = await authorizedFetch(cfg, `${GRAPH}${path}`, {
    ...init,
    headers: { "Content-Type": "application/json", ...(init.headers || {}) },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph ${res.status}: ${body}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

// =========== Blob caches (object URLs, LRU) ===========
// Object URLs pin their blob in memory until revoked, so the caches are
// bounded and revoke whatever they evict. A loader returns null for content
// that does not exist (cached) and throws on transient failures (not cached).

class BlobUrlCache {
  private entries = new Map<string, Promise<string | null>>();
  constructor(private max: number) {}

  get(key: string, load: () => Promise<Blob | null>): Promise<string | null> {
    const hit = this.entries.get(key);
    if (hit) {
      this.entries.delete(key);
      this.entries.set(key, hit);
      return hit;
    }
    const p = load()
      .then((blob) => (blob ? URL.createObjectURL(blob) : null))
      .catch(() => {
        this.entries.delete(key);
        return null;
      });
    this.entries.set(key, p);
    this.evict();
    return p;
  }

  private evict() {
    while (this.entries.size > this.max) {
      const [key, p] = this.entries.entries().next().value!;
      this.entries.delete(key);
      p.then((url) => { if (url) URL.revokeObjectURL(url); });
    }
  }
}

// =========== User photos ===========

const photoCache = new BlobUrlCache(200);

export function getUserPhotoUrl(cfg: TeamsConfig, userId: string): Promise<string | null> {
  if (!userId) return Promise.resolve(null);
  return photoCache.get(userId, async () => {
    // 96x96 covers the largest avatar (40px) on 2x screens; fall back to the
    // original only when Graph has no resized version.
    let res = await authorizedFetch(cfg, `${GRAPH}/users/${userId}/photos/96x96/$value`);
    if (res.status === 404) res = await authorizedFetch(cfg, `${GRAPH}/users/${userId}/photo/$value`);
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`Graph ${res.status}`);
    return res.blob();
  });
}

// =========== Hosted contents (inline images in messages) ===========

const hostedContentCache = new BlobUrlCache(60);

export function getHostedContentUrl(cfg: TeamsConfig, url: string): Promise<string | null> {
  if (!url) return Promise.resolve(null);
  return hostedContentCache.get(url, async () => {
    const res = await authorizedFetch(cfg, url);
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`Graph ${res.status}`);
    return res.blob();
  });
}

// =========== Concurrency ===========

export async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const out = new Array<R>(items.length);
  let next = 0;
  const worker = async () => {
    while (next < items.length) {
      const i = next++;
      out[i] = await fn(items[i]);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return out;
}

export type GraphChatMember = {
  displayName?: string | null;
  userId?: string | null;
  email?: string | null;
};

export type GraphChat = {
  id: string;
  topic: string | null;
  chatType: "oneOnOne" | "group" | "meeting" | string;
  lastUpdatedDateTime: string;
  members?: GraphChatMember[];
  viewpoint?: { isHidden?: boolean; lastMessageReadDateTime?: string } | null;
  lastMessagePreview?: {
    createdDateTime?: string;
    from?: {
      user?: { displayName?: string; id?: string };
    } | null;
    body?: { content?: string; contentType?: string };
  } | null;
};

export type GraphAttachment = {
  id: string;
  contentType: string;
  contentUrl?: string | null;
  name?: string | null;
  content?: string | null;
  thumbnailUrl?: string | null;
};

export type GraphMessage = {
  id: string;
  createdDateTime: string;
  lastModifiedDateTime?: string;
  from?: {
    user?: { displayName?: string; id?: string };
  } | null;
  body: { contentType: "html" | "text"; content: string };
  attachments?: GraphAttachment[];
};

export async function listChats(cfg: TeamsConfig): Promise<GraphChat[]> {
  const data = await graphFetch(
    cfg,
    "/me/chats?$expand=members,lastMessagePreview&$orderby=lastMessagePreview/createdDateTime desc&$top=50",
  );
  return data.value as GraphChat[];
}

export async function hideChat(
  cfg: TeamsConfig,
  chatId: string,
  hide: boolean,
  user: { id: string; tenantId: string },
) {
  await graphFetch(cfg, `/chats/${chatId}/hideForUser`, {
    method: "POST",
    body: JSON.stringify({
      user: { id: user.id, tenantId: user.tenantId },
      hideForUser: hide,
    }),
  });
}

export async function listMessages(cfg: TeamsConfig, chatId: string): Promise<GraphMessage[]> {
  const data = await graphFetch(cfg, `/me/chats/${chatId}/messages?$top=50`);
  // Graph returns newest first; reverse for chronological display
  return (data.value as GraphMessage[]).slice().reverse();
}

export async function sendMessage(cfg: TeamsConfig, chatId: string, text: string) {
  return graphFetch(cfg, `/me/chats/${chatId}/messages`, {
    method: "POST",
    body: JSON.stringify({
      body: { contentType: "text", content: text },
    }),
  });
}

export function chatTitle(chat: GraphChat, meId?: string, meName?: string): string {
  if (chat.topic) return chat.topic;
  const others = (chat.members ?? []).filter((m) => !meId || m.userId !== meId);
  const names = others.map((m) => m.displayName).filter(Boolean) as string[];
  if (names.length > 0) {
    if (names.length <= 2) return names.join(", ");
    return `${names.slice(0, 2).join(", ")} +${names.length - 2}`;
  }
  const previewName = chat.lastMessagePreview?.from?.user?.displayName;
  if (previewName && previewName !== meName) return previewName;
  const emails = others.map((m) => m.email).filter(Boolean) as string[];
  if (emails.length > 0) return emails.join(", ");
  return "Usuário externo";
}

export function buildTeamsVideoLink(chat: GraphChat, meId?: string): string | null {
  const others = (chat.members ?? []).filter((m) => !meId || m.userId !== meId);
  const emails = others.map((m) => m.email).filter((e): e is string => !!e);
  if (emails.length === 0) return null;
  return `https://teams.microsoft.com/l/call/0/0?users=${encodeURIComponent(emails.join(","))}&withVideo=true`;
}

// ============ Teams / Channels ============

export type GraphTeam = {
  id: string;
  displayName: string;
  description?: string;
};

export type GraphChannel = {
  id: string;
  displayName: string;
  description?: string;
  membershipType?: string;
};

export async function listJoinedTeams(cfg: TeamsConfig): Promise<GraphTeam[]> {
  const data = await graphFetch(cfg, "/me/joinedTeams");
  return data.value as GraphTeam[];
}

export async function listChannels(cfg: TeamsConfig, teamId: string): Promise<GraphChannel[]> {
  const data = await graphFetch(cfg, `/teams/${teamId}/channels`);
  return data.value as GraphChannel[];
}

export async function listChannelMessages(
  cfg: TeamsConfig,
  teamId: string,
  channelId: string,
): Promise<GraphMessage[]> {
  const data = await graphFetch(
    cfg,
    `/teams/${teamId}/channels/${channelId}/messages?$top=50`,
  );
  return (data.value as GraphMessage[]).slice().reverse();
}

export async function getChannelLastMessageDate(
  cfg: TeamsConfig,
  teamId: string,
  channelId: string,
): Promise<string | null> {
  try {
    const data = await graphFetch(
      cfg,
      `/teams/${teamId}/channels/${channelId}/messages?$top=1`,
    );
    const first = (data.value as GraphMessage[])[0];
    return first?.lastModifiedDateTime ?? first?.createdDateTime ?? null;
  } catch {
    return null;
  }
}

export async function sendChannelMessage(
  cfg: TeamsConfig,
  teamId: string,
  channelId: string,
  text: string,
) {
  return graphFetch(cfg, `/teams/${teamId}/channels/${channelId}/messages`, {
    method: "POST",
    body: JSON.stringify({
      body: { contentType: "text", content: text },
    }),
  });
}

