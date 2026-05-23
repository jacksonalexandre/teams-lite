import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import { useServerFn } from "@tanstack/react-start";
import { Archive, ArchiveRestore } from "lucide-react";
import {
  signIn,
  signOut,
  ensureInit,
  getCurrentAccount,
  type TeamsConfig,
} from "@/lib/msal";
import { getTeamsConfig } from "@/lib/config.functions";
import {
  listChats,
  listMessages,
  sendMessage,
  chatTitle,
  hideChat,
  listJoinedTeams,
  listChannels,
  listChannelMessages,
  sendChannelMessage,
  type GraphChat,
  type GraphMessage,
  type GraphTeam,
  type GraphChannel,
} from "@/lib/graph";

export const Route = createFileRoute("/")({
  component: TeamsLite,
});

type Mode = "chats" | "channels";
type Selection =
  | { kind: "chat"; chatId: string }
  | { kind: "channel"; teamId: string; channelId: string }
  | null;

function TeamsLite() {
  const fetchConfig = useServerFn(getTeamsConfig);
  const [config, setConfig] = useState<TeamsConfig | null>(null);
  const [account, setAccount] = useState<{ name?: string; username: string; oid?: string } | null>(null);
  const [mode, setMode] = useState<Mode>("chats");

  // Chats state
  const [chats, setChats] = useState<GraphChat[]>([]);
  const [loadingChats, setLoadingChats] = useState(false);
  const [visibleCount, setVisibleCount] = useState(7);
  const [showHidden, setShowHidden] = useState(false);
  const [hidingId, setHidingId] = useState<string | null>(null);

  // Channels state
  const [teams, setTeams] = useState<GraphTeam[]>([]);
  const [loadingTeams, setLoadingTeams] = useState(false);
  const [expandedTeam, setExpandedTeam] = useState<string | null>(null);
  const [channelsByTeam, setChannelsByTeam] = useState<Record<string, GraphChannel[]>>({});
  const [loadingChannels, setLoadingChannels] = useState<Record<string, boolean>>({});

  // Selection + messages
  const [selection, setSelection] = useState<Selection>(null);
  const [messages, setMessages] = useState<GraphMessage[]>([]);
  const [draft, setDraft] = useState("");
  const [loadingMsgs, setLoadingMsgs] = useState(false);
  const [sending, setSending] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [msgsVisibleCount, setMsgsVisibleCount] = useState(7);
  const scrollRef = useRef<HTMLDivElement>(null);
  const prevSelKeyRef = useRef<string | null>(null);
  const prevLastMsgIdRef = useRef<string | null>(null);

  useEffect(() => {
    fetchConfig()
      .then((cfg) => {
        setConfig(cfg);
        return ensureInit(cfg).then(() => getCurrentAccount(cfg));
      })
      .then((acc) => {
        if (acc) setAccount({ name: acc.name, username: acc.username, oid: (acc as any).idTokenClaims?.oid || acc.localAccountId });
      })
      .catch((e) => setError(String(e)));
  }, []); // eslint-disable-line

  // Load chats
  useEffect(() => {
    if (!config || !account || mode !== "chats") return;
    setLoadingChats(true);
    setVisibleCount(7);
    listChats(config)
      .then((cs) => {
        setChats(cs);
      })
      .catch((e) => setError(String(e)))
      .finally(() => setLoadingChats(false));
  }, [config, account, mode]); // eslint-disable-line

  // Load teams
  useEffect(() => {
    if (!config || !account || mode !== "channels") return;
    setLoadingTeams(true);
    listJoinedTeams(config)
      .then((ts) => setTeams(ts))
      .catch((e) => setError(String(e)))
      .finally(() => setLoadingTeams(false));
  }, [config, account, mode]);

  async function toggleTeam(teamId: string) {
    if (!config) return;
    if (expandedTeam === teamId) {
      setExpandedTeam(null);
      return;
    }
    setExpandedTeam(teamId);
    if (!channelsByTeam[teamId]) {
      setLoadingChannels((p) => ({ ...p, [teamId]: true }));
      try {
        const chs = await listChannels(config, teamId);
        setChannelsByTeam((p) => ({ ...p, [teamId]: chs }));
      } catch (e) {
        setError(String(e));
      } finally {
        setLoadingChannels((p) => ({ ...p, [teamId]: false }));
      }
    }
  }

  // Load messages (chat or channel) + poll
  useEffect(() => {
    if (!config || !selection) {
      setMessages([]);
      return;
    }
    let cancelled = false;
    const load = async (showLoader: boolean) => {
      if (showLoader) setLoadingMsgs(true);
      try {
        const msgs =
          selection.kind === "chat"
            ? await listMessages(config, selection.chatId)
            : await listChannelMessages(config, selection.teamId, selection.channelId);
        if (!cancelled) setMessages(msgs);
      } catch (e) {
        if (!cancelled) setError(String(e));
      } finally {
        if (!cancelled && showLoader) setLoadingMsgs(false);
      }
    };
    load(true);
    const t = setInterval(() => load(false), 5000);
    return () => {
      cancelled = true;
      clearInterval(t);
    };
  }, [config, selection]);

  const selKey = selection
    ? selection.kind === "chat"
      ? `c:${selection.chatId}`
      : `ch:${selection.teamId}:${selection.channelId}`
    : null;

  useEffect(() => {
    setMsgsVisibleCount(7);
  }, [selKey]);

  useEffect(() => {
    const lastId = messages[messages.length - 1]?.id ?? null;
    const selChanged = prevSelKeyRef.current !== selKey;
    const newMessage = prevLastMsgIdRef.current !== lastId;
    if (selChanged || newMessage) {
      scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight });
    }
    prevSelKeyRef.current = selKey;
    prevLastMsgIdRef.current = lastId;
  }, [messages, selKey]);

  const meId = useMemo(() => account?.oid, [account]);

  async function toggleHide(chatId: string, currentlyHidden: boolean) {
    if (!config || hidingId) return;
    setHidingId(chatId);
    try {
      await hideChat(config, chatId, !currentlyHidden);
      setChats((prev) =>
        prev.map((c) =>
          c.id === chatId
            ? { ...c, viewpoint: { ...(c.viewpoint ?? {}), isHidden: !currentlyHidden } }
            : c,
        ),
      );
      if (selection?.kind === "chat" && selection.chatId === chatId) {
        setSelection(null);
      }
    } catch (e) {
      setError(String(e));
    } finally {
      setHidingId(null);
    }
  }

  const filteredChats = useMemo(() => {
    return chats.filter((c) => {
      const hidden = !!c.viewpoint?.isHidden;
      return showHidden ? hidden : !hidden;
    });
  }, [chats, showHidden]);

  async function handleSignIn() {
    if (!config) return;
    setError(null);
    try {
      const acc = await signIn(config);
      setAccount({ name: acc.name, username: acc.username, oid: (acc as any).idTokenClaims?.oid || acc.localAccountId });
    } catch (e) {
      setError(String(e));
    }
  }

  async function handleSignOut() {
    if (!config) return;
    try {
      await signOut(config);
      setAccount(null);
      setChats([]);
      setTeams([]);
      setMessages([]);
      setSelection(null);
    } catch (e) {
      setError(String(e));
    }
  }

  async function handleSend() {
    if (!config || !selection || !draft.trim() || sending) return;
    setSending(true);
    const text = draft;
    setDraft("");
    try {
      if (selection.kind === "chat") {
        await sendMessage(config, selection.chatId, text);
        const msgs = await listMessages(config, selection.chatId);
        setMessages(msgs);
      } else {
        await sendChannelMessage(config, selection.teamId, selection.channelId, text);
        const msgs = await listChannelMessages(config, selection.teamId, selection.channelId);
        setMessages(msgs);
      }
    } catch (e) {
      setError(String(e));
      setDraft(text);
    } finally {
      setSending(false);
    }
  }

  const headerTitle = useMemo(() => {
    if (!selection) return "";
    if (selection.kind === "chat") {
      const c = chats.find((x) => x.id === selection.chatId);
      return c ? chatTitle(c, meId, account?.name) : "";
    }
    const team = teams.find((t) => t.id === selection.teamId);
    const ch = channelsByTeam[selection.teamId]?.find((c) => c.id === selection.channelId);
    return team && ch ? `${team.displayName} · ${ch.displayName}` : "";
  }, [selection, chats, teams, channelsByTeam, meId, account]);

  return (
    <div className="flex h-screen flex-col bg-background text-foreground">
      <header className="flex items-center justify-between border-b border-border px-5 py-3">
        <div className="flex items-center gap-2">
          <div className="grid h-7 w-7 place-items-center rounded-md bg-primary text-primary-foreground text-sm font-semibold">
            T
          </div>
          <h1 className="text-base font-semibold tracking-tight">Teams Lite</h1>
        </div>
        <div className="flex items-center gap-2 text-sm">
          {account ? (
            <>
              <span className="text-muted-foreground hidden sm:inline">{account.name ?? account.username}</span>
              <button
                onClick={handleSignOut}
                className="rounded-md border border-border px-3 py-1.5 text-xs font-medium hover:bg-muted"
              >
                Sair
              </button>
            </>
          ) : (
            config && (
              <button
                onClick={handleSignIn}
                className="rounded-md bg-primary px-3 py-1.5 text-xs font-medium text-primary-foreground hover:opacity-90"
              >
                Entrar com Microsoft
              </button>
            )
          )}
        </div>
      </header>

      {error && (
        <div className="border-b border-destructive/30 bg-destructive/10 px-5 py-2 text-xs text-destructive">
          <div className="flex items-start justify-between gap-3">
            <pre className="whitespace-pre-wrap break-all font-mono">{error}</pre>
            <button onClick={() => setError(null)} className="shrink-0 underline">
              fechar
            </button>
          </div>
        </div>
      )}

      <div className="flex min-h-0 flex-1">
        {/* Sidebar */}
        <aside className="flex w-72 flex-col border-r border-border bg-card">
          <div className="flex border-b border-border">
            <button
              onClick={() => setMode("chats")}
              className={`flex-1 px-4 py-2.5 text-xs font-medium uppercase tracking-wider transition-colors ${
                mode === "chats" ? "bg-muted text-foreground" : "text-muted-foreground hover:bg-muted/60"
              }`}
            >
              Conversas
            </button>
            <button
              onClick={() => setMode("channels")}
              className={`flex-1 px-4 py-2.5 text-xs font-medium uppercase tracking-wider transition-colors ${
                mode === "channels" ? "bg-muted text-foreground" : "text-muted-foreground hover:bg-muted/60"
              }`}
            >
              Canais
            </button>
          </div>

          {mode === "chats" && account && (
            <div className="flex items-center justify-between border-b border-border px-4 py-2">
              <span className="text-[11px] text-muted-foreground">
                {showArchived ? "Arquivados" : "Ativos"} ({filteredChats.length})
              </span>
              <button
                onClick={() => {
                  setShowArchived((v) => !v);
                  setVisibleCount(7);
                }}
                className="text-[11px] font-medium text-primary hover:underline"
              >
                {showArchived ? "Ver ativos" : `Ver arquivados (${archived.size})`}
              </button>
            </div>
          )}

          <div className="flex-1 overflow-y-auto">
            {!account ? (
              <EmptyHint text="Faça login para começar." />
            ) : mode === "chats" ? (
              loadingChats ? (
                <EmptyHint text="Carregando…" />
              ) : filteredChats.length === 0 ? (
                <EmptyHint text={showArchived ? "Nenhum chat arquivado." : "Nenhum chat encontrado."} />
              ) : (
                <>
                  {filteredChats.slice(0, visibleCount).map((c) => {
                    const title = chatTitle(c, meId, account.name);
                    const active = selection?.kind === "chat" && selection.chatId === c.id;
                    const isArch = archived.has(c.id);
                    return (
                      <div
                        key={c.id}
                        className={`group flex w-full items-start gap-2 border-b border-border px-3 py-3 text-left text-sm transition-colors ${
                          active ? "bg-muted" : "hover:bg-muted/60"
                        }`}
                      >
                        <button
                          onClick={() => setSelection({ kind: "chat", chatId: c.id })}
                          className="flex min-w-0 flex-1 flex-col items-start gap-0.5"
                        >
                          <span className="line-clamp-1 font-medium">{title}</span>
                          <span className="text-[11px] text-muted-foreground">
                            {formatDateTime(c.lastMessagePreview?.createdDateTime ?? c.lastUpdatedDateTime)}
                            {c.chatType === "group" ? " · Grupo" : ""}
                          </span>
                        </button>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            toggleArchive(c.id);
                          }}
                          title={isArch ? "Desarquivar" : "Arquivar"}
                          className="opacity-0 group-hover:opacity-100 shrink-0 rounded px-1.5 py-1 text-[10px] font-medium text-muted-foreground hover:bg-background hover:text-foreground transition-opacity"
                        >
                          {isArch ? "↩" : "📁"}
                        </button>
                      </div>
                    );
                  })}
                  {filteredChats.length > visibleCount && (
                    <button
                      onClick={() => setVisibleCount((n) => n + 7)}
                      className="flex w-full items-center justify-center border-b border-border px-4 py-2 text-xs font-medium text-muted-foreground hover:bg-muted transition-colors"
                    >
                      Carregar mais ({filteredChats.length - visibleCount} restantes)
                    </button>
                  )}
                </>
              )
            ) : loadingTeams ? (
              <EmptyHint text="Carregando equipes…" />
            ) : teams.length === 0 ? (
              <EmptyHint text="Nenhuma equipe encontrada." />
            ) : (
              teams.map((t) => {
                const open = expandedTeam === t.id;
                const chs = channelsByTeam[t.id];
                return (
                  <div key={t.id} className="border-b border-border">
                    <button
                      onClick={() => toggleTeam(t.id)}
                      className="flex w-full items-center justify-between px-4 py-3 text-left text-sm font-medium hover:bg-muted/60"
                    >
                      <span className="line-clamp-1">{t.displayName}</span>
                      <span className="text-xs text-muted-foreground">{open ? "▾" : "▸"}</span>
                    </button>
                    {open && (
                      <div className="bg-background/40">
                        {loadingChannels[t.id] ? (
                          <EmptyHint text="Carregando canais…" />
                        ) : chs && chs.length > 0 ? (
                          chs.map((ch) => {
                            const active =
                              selection?.kind === "channel" &&
                              selection.teamId === t.id &&
                              selection.channelId === ch.id;
                            return (
                              <button
                                key={ch.id}
                                onClick={() =>
                                  setSelection({ kind: "channel", teamId: t.id, channelId: ch.id })
                                }
                                className={`flex w-full items-center gap-2 px-6 py-2 text-left text-sm transition-colors ${
                                  active ? "bg-muted" : "hover:bg-muted/60"
                                }`}
                              >
                                <span className="text-muted-foreground">#</span>
                                <span className="line-clamp-1">{ch.displayName}</span>
                              </button>
                            );
                          })
                        ) : (
                          <EmptyHint text="Sem canais." />
                        )}
                      </div>
                    )}
                  </div>
                );
              })
            )}
          </div>
        </aside>

        {/* Main */}
        <main className="flex min-w-0 flex-1 flex-col">
          {!selection || !account ? (
            <div className="grid flex-1 place-items-center text-sm text-muted-foreground">
              {account
                ? mode === "chats"
                  ? "Selecione uma conversa"
                  : "Selecione um canal"
                : config
                  ? "Não autenticado"
                  : "Carregando configuração…"}
            </div>
          ) : (
            <>
              {headerTitle && (
                <div className="border-b border-border bg-card px-6 py-2.5 text-sm font-medium">
                  {headerTitle}
                </div>
              )}
              <div ref={scrollRef} className="flex-1 space-y-3 overflow-y-auto px-6 py-5">
                {loadingMsgs && messages.length === 0 ? (
                  <EmptyHint text="Carregando mensagens…" />
                ) : messages.length === 0 ? (
                  <EmptyHint text="Sem mensagens ainda." />
                ) : (
                  <>
                    {messages.length > msgsVisibleCount && (
                      <button
                        onClick={() => setMsgsVisibleCount((n) => n + 7)}
                        className="mx-auto block rounded-md border border-border px-3 py-1.5 text-xs font-medium text-muted-foreground hover:bg-muted transition-colors"
                      >
                        Carregar mais ({messages.length - msgsVisibleCount} restantes)
                      </button>
                    )}
                    {messages.slice(-msgsVisibleCount).map((m) => (
                      <MessageBubble key={m.id} m={m} meName={account.name} />
                    ))}
                  </>
                )}
              </div>
              <div className="border-t border-border bg-card px-4 py-3">
                <div className="flex items-end gap-2">
                  <textarea
                    value={draft}
                    onChange={(e) => setDraft(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter" && !e.shiftKey) {
                        e.preventDefault();
                        handleSend();
                      }
                    }}
                    placeholder="Escreva uma mensagem… (Enter envia, Shift+Enter quebra linha)"
                    rows={1}
                    className="max-h-40 min-h-[40px] flex-1 resize-none rounded-md border border-input bg-background px-3 py-2 text-sm outline-none focus:border-ring focus:ring-2 focus:ring-ring/30"
                  />
                  <button
                    onClick={handleSend}
                    disabled={!draft.trim() || sending}
                    className="rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground transition-opacity hover:opacity-90 disabled:opacity-40"
                  >
                    {sending ? "…" : "Enviar"}
                  </button>
                </div>
              </div>
            </>
          )}
        </main>
      </div>
    </div>
  );
}

function EmptyHint({ text }: { text: string }) {
  return <div className="px-4 py-10 text-center text-xs text-muted-foreground">{text}</div>;
}

function MessageBubble({ m, meName }: { m: GraphMessage; meName?: string }) {
  const author = m.from?.user?.displayName ?? "Sistema";
  const mine = !!meName && author === meName;
  const text = m.body.contentType === "html" ? stripHtml(m.body.content) : m.body.content;
  if (!text.trim()) return null;
  return (
    <div className={`flex ${mine ? "justify-end" : "justify-start"}`}>
      <div
        className={`max-w-[75%] rounded-2xl px-4 py-2 text-sm shadow-sm ${
          mine
            ? "bg-primary text-primary-foreground rounded-br-sm"
            : "bg-card text-card-foreground border border-border rounded-bl-sm"
        }`}
      >
        {!mine && <div className="mb-0.5 text-[11px] font-medium opacity-70">{author}</div>}
        <div className="whitespace-pre-wrap break-words">{text}</div>
        <div className={`mt-1 text-[10px] ${mine ? "text-primary-foreground/70" : "text-muted-foreground"}`}>
          {formatDateTime(m.createdDateTime)}
        </div>
      </div>
    </div>
  );
}

function stripHtml(html: string) {
  if (typeof window === "undefined") return html.replace(/<[^>]*>/g, "");
  const div = document.createElement("div");
  div.innerHTML = html;
  return div.textContent || div.innerText || "";
}

function formatDateTime(iso: string) {
  try {
    const d = new Date(iso);
    const today = new Date();
    const time = d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    if (d.toDateString() === today.toDateString()) return time;
    const date = d.toLocaleDateString();
    return `${date} ${time}`;
  } catch {
    return "";
  }
}
