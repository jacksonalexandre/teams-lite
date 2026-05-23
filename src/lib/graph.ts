import { getAccessToken, type TeamsConfig } from "./msal";

const GRAPH = "https://graph.microsoft.com/v1.0";

async function graphFetch(cfg: TeamsConfig, path: string, init: RequestInit = {}) {
  const token = await getAccessToken(cfg);
  const res = await fetch(`${GRAPH}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(init.headers || {}),
    },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph ${res.status}: ${body}`);
  }
  return res.json();
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

export type GraphMessage = {
  id: string;
  createdDateTime: string;
  from?: {
    user?: { displayName?: string; id?: string };
  } | null;
  body: { contentType: "html" | "text"; content: string };
};

export async function listChats(cfg: TeamsConfig): Promise<GraphChat[]> {
  const data = await graphFetch(
    cfg,
    "/me/chats?$expand=members&$orderby=lastMessagePreview/createdDateTime desc&$top=50",
  );
  return data.value as GraphChat[];
}

export async function hideChat(cfg: TeamsConfig, chatId: string, hide: boolean) {
  const token = await getAccessToken(cfg);
  const res = await fetch(`${GRAPH}/me/chats/${chatId}/hideForUser`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      user: { "@odata.id": `https://graph.microsoft.com/v1.0/me` },
      hideForUser: hide,
    }),
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph ${res.status}: ${body}`);
  }
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

