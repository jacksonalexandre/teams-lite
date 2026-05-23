import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import { useServerFn } from "@tanstack/react-start";
import { Archive, ArchiveRestore, Loader2, AlertCircle, X } from "lucide-react";
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
  getChannelLastMessageDate,
  getUserPhotoUrl,
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
  const [hidingId, setHidingId] = useState<string | null>(null);

  // Channels state (flat list across all joined teams, sorted by last activity)
  const [channelList, setChannelList] = useState<
    Array<{ team: GraphTeam; channel: GraphChannel; lastDate: string | null }>
  >([]);
  const [loadingChannels, setLoadingChannels] = useState(false);

  // Selection + messages
  const [selection, setSelection] = useState<Selection>(null);
  const [messages, setMessages] = useState<GraphMessage[]>([]);
  const [draft, setDraft] = useState("");
  const [loadingMsgs, setLoadingMsgs] = useState(false);
  const [sending, setSending] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [msgsVisibleCount, setMsgsVisibleCount] = useState(7);
  type Pending = { id: string; selKey: string; text: string; status: "sending" | "error"; error?: string };
  const [pending, setPending] = useState<Pending[]>([]);
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
    listChats(config)
      .then((cs) => {
        setChats(cs);
      })
      .catch((e) => setError(String(e)))
      .finally(() => setLoadingChats(false));
  }, [config, account, mode]); // eslint-disable-line

  // Load all channels across joined teams, sorted by last activity
  useEffect(() => {
    if (!config || !account || mode !== "channels") return;
    let cancelled = false;
    setLoadingChannels(true);
    (async () => {
      try {
        const ts = await listJoinedTeams(config);
        const channelsPerTeam = await Promise.all(
          ts.map((t) =>
            listChannels(config, t.id)
              .then((chs) => chs.map((channel) => ({ team: t, channel })))
              .catch(() => [] as Array<{ team: GraphTeam; channel: GraphChannel }>),
          ),
        );
        const flat = channelsPerTeam.flat();
        const withDates = await Promise.all(
          flat.map(async (item) => ({
            ...item,
            lastDate: await getChannelLastMessageDate(config, item.team.id, item.channel.id),
          })),
        );
        withDates.sort((a, b) => {
          const ad = a.lastDate ? new Date(a.lastDate).getTime() : 0;
          const bd = b.lastDate ? new Date(b.lastDate).getTime() : 0;
          return bd - ad;
        });
        const top12 = withDates.filter((x) => x.lastDate).slice(0, 12);
        if (!cancelled) setChannelList(top12);
      } catch (e) {
        if (!cancelled) setError(String(e));
      } finally {
        if (!cancelled) setLoadingChannels(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [config, account, mode]);


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

  const [atBottom, setAtBottom] = useState(true);
  const [hasNewBelow, setHasNewBelow] = useState(false);

  useEffect(() => {
    setMsgsVisibleCount(7);
    setHasNewBelow(false);
    setAtBottom(true);
  }, [selKey]);

  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const onScroll = () => {
      const near = el.scrollHeight - el.scrollTop - el.clientHeight < 40;
      setAtBottom(near);
      if (near) setHasNewBelow(false);
    };
    el.addEventListener("scroll", onScroll);
    return () => el.removeEventListener("scroll", onScroll);
  }, [selKey]);

  useEffect(() => {
    const lastId = messages[messages.length - 1]?.id ?? null;
    const selChanged = prevSelKeyRef.current !== selKey;
    const newMessage = prevLastMsgIdRef.current !== lastId;
    if (selChanged) {
      scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight });
    } else if (newMessage) {
      if (atBottom) {
        scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight, behavior: "smooth" });
      } else if (prevLastMsgIdRef.current !== null) {
        setHasNewBelow(true);
      }
    }
    prevSelKeyRef.current = selKey;
    prevLastMsgIdRef.current = lastId;
  }, [messages, selKey, atBottom]);

  function scrollToBottom() {
    const el = scrollRef.current;
    if (!el) return;
    el.scrollTo({ top: el.scrollHeight, behavior: "smooth" });
    setHasNewBelow(false);
  }


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

  const filteredChats = useMemo(
    () => chats.filter((c) => !c.viewpoint?.isHidden),
    [chats],
  );

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
      setChannelList([]);
      setMessages([]);
      setSelection(null);
    } catch (e) {
      setError(String(e));
    }
  }

  async function doSend(p: Pending) {
    if (!config) return;
    const sel = selection;
    if (!sel) return;
    setPending((prev) => prev.map((x) => (x.id === p.id ? { ...x, status: "sending", error: undefined } : x)));
    try {
      if (sel.kind === "chat") {
        await sendMessage(config, sel.chatId, p.text);
        const msgs = await listMessages(config, sel.chatId);
        setMessages(msgs);
      } else {
        await sendChannelMessage(config, sel.teamId, sel.channelId, p.text);
        const msgs = await listChannelMessages(config, sel.teamId, sel.channelId);
        setMessages(msgs);
      }
      setPending((prev) => prev.filter((x) => x.id !== p.id));
    } catch (e) {
      setPending((prev) => prev.map((x) => (x.id === p.id ? { ...x, status: "error", error: String(e) } : x)));
    }
  }

  async function handleSend() {
    if (!config || !selection || !draft.trim() || sending) return;
    const text = draft.trim();
    setDraft("");
    setSending(true);
    const p: Pending = {
      id: `p-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      selKey: selection.kind === "chat" ? `c:${selection.chatId}` : `ch:${selection.teamId}:${selection.channelId}`,
      text,
      status: "sending",
    };
    setPending((prev) => [...prev, p]);
    try {
      await doSend(p);
    } finally {
      setSending(false);
    }
  }

  function retryPending(id: string) {
    const p = pending.find((x) => x.id === id);
    if (p) doSend(p);
  }

  function discardPending(id: string) {
    setPending((prev) => prev.filter((x) => x.id !== id));
  }

  const headerTitle = useMemo(() => {
    if (!selection) return "";
    if (selection.kind === "chat") {
      const c = chats.find((x) => x.id === selection.chatId);
      return c ? chatTitle(c, meId, account?.name) : "";
    }
    const item = channelList.find(
      (x) => x.team.id === selection.teamId && x.channel.id === selection.channelId,
    );
    return item ? `${item.team.displayName} · ${item.channel.displayName}` : "";
  }, [selection, chats, channelList, meId, account]);

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
        <aside className="flex w-80 flex-col border-r border-border bg-card">
          <div className="flex border-b border-border">
            <button
              onClick={() => setMode("chats")}
              className={`flex-1 px-4 py-2.5 text-xs font-medium uppercase tracking-wider transition-colors ${
                mode === "chats" ? "bg-muted text-foreground" : "text-muted-foreground hover:bg-muted/60"
              }`}
            >
              Chats
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


          <div className="flex-1 overflow-y-auto">
            {!account ? (
              <EmptyHint text="Faça login para começar." />
            ) : mode === "chats" ? (
              loadingChats ? (
                <EmptyHint text="Carregando…" />
              ) : filteredChats.length === 0 ? (
                <EmptyHint text="Nenhum chat encontrado." />
              ) : (
                filteredChats.map((c) => {
                  const title = chatTitle(c, meId, account.name);
                  const active = selection?.kind === "chat" && selection.chatId === c.id;
                  const isHidden = !!c.viewpoint?.isHidden;
                  const isBusy = hidingId === c.id;
                  const preview = c.lastMessagePreview;
                  const previewAuthor = preview?.from?.user?.displayName;
                  const previewIsMe = !!account.name && previewAuthor === account.name;
                  const previewText = preview?.body?.content
                    ? preview.body.contentType === "html"
                      ? stripHtml(preview.body.content)
                      : preview.body.content
                    : "";
                  const previewPrefix = previewIsMe
                    ? "Você: "
                    : previewAuthor
                      ? `${previewAuthor.split(" ")[0]}: `
                      : "";
                  const isGroup = c.chatType === "group" || c.chatType === "meeting";
                  const otherUserId = !isGroup
                    ? (c.members ?? []).find((m) => m.userId && m.userId !== meId)?.userId ?? undefined
                    : undefined;
                  return (
                    <div
                      key={c.id}
                      onClick={() => setSelection({ kind: "chat", chatId: c.id })}
                      className={`group flex w-full cursor-pointer items-center gap-3 border-b border-border px-3 py-2.5 text-sm transition-colors ${
                        active ? "bg-muted" : "hover:bg-muted/60"
                      }`}
                      title={title}
                    >
                      <Avatar name={title} isGroup={isGroup} userId={otherUserId} cfg={config} />
                      <div className="flex min-w-0 flex-1 flex-col">
                        <div className="flex items-baseline justify-between gap-2">
                          <span className="truncate text-left font-semibold text-foreground" title={title}>
                            {title}
                          </span>
                          <span className="shrink-0 text-[11px] text-muted-foreground">
                            {formatListDate(preview?.createdDateTime ?? c.lastUpdatedDateTime)}
                          </span>
                        </div>
                        <div className="flex items-center justify-between gap-2">
                          <span className="truncate text-left text-xs text-muted-foreground">
                            {previewPrefix}
                            {previewText || (isGroup ? "Grupo" : "")}
                          </span>
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              toggleHide(c.id, isHidden);
                            }}
                            disabled={isBusy}
                            title={isHidden ? "Reexibir no Teams" : "Ocultar no Teams"}
                            className="opacity-0 group-hover:opacity-100 focus:opacity-100 shrink-0 rounded p-1 text-muted-foreground hover:bg-background hover:text-foreground transition-opacity disabled:opacity-40"
                          >
                            {isHidden ? <ArchiveRestore size={14} /> : <Archive size={14} />}
                          </button>
                        </div>
                      </div>
                    </div>
                  );
                })
              )
            ) : loadingChannels ? (
              <EmptyHint text="Carregando canais…" />
            ) : channelList.length === 0 ? (
              <EmptyHint text="Nenhum canal encontrado." />
            ) : (
              channelList.map(({ team, channel, lastDate }) => {
                const active =
                  selection?.kind === "channel" &&
                  selection.teamId === team.id &&
                  selection.channelId === channel.id;
                const fullTitle = `${team.displayName} · #${channel.displayName}`;
                return (
                  <div
                    key={`${team.id}:${channel.id}`}
                    onClick={() =>
                      setSelection({ kind: "channel", teamId: team.id, channelId: channel.id })
                    }
                    title={fullTitle}
                    className={`flex w-full cursor-pointer items-center gap-3 border-b border-border px-3 py-2.5 text-sm transition-colors ${
                      active ? "bg-muted" : "hover:bg-muted/60"
                    }`}
                  >
                    <Avatar name={team.displayName} isGroup />
                    <div className="flex min-w-0 flex-1 flex-col">
                      <div className="flex items-baseline justify-between gap-2">
                        <span className="truncate text-left font-semibold text-foreground" title={fullTitle}>
                          {team.displayName}
                        </span>
                        <span className="shrink-0 text-[11px] text-muted-foreground">
                          {lastDate ? formatListDate(lastDate) : ""}
                        </span>
                      </div>
                      <span className="truncate text-left text-xs text-muted-foreground" title={fullTitle}>
                        #{channel.displayName}
                      </span>
                    </div>
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
              <div className="relative flex min-h-0 flex-1 flex-col">
              <div ref={scrollRef} className="flex-1 space-y-3 overflow-y-auto px-6 py-5">
                {loadingMsgs && messages.length === 0 ? (
                  <EmptyHint text="Carregando mensagens…" />
                ) : messages.length === 0 && pending.filter((p) => p.selKey === selKey).length === 0 ? (
                  <EmptyHint text="Sem mensagens ainda." />
                ) : (
                  <>
                    {messages.length > 0 && (
                      <button
                        onClick={() => {
                          if (messages.length <= msgsVisibleCount) return;
                          const el = scrollRef.current;
                          const prevHeight = el?.scrollHeight ?? 0;
                          const prevTop = el?.scrollTop ?? 0;
                          setMsgsVisibleCount((n) => n + 7);
                          requestAnimationFrame(() => {
                            const el2 = scrollRef.current;
                            if (!el2) return;
                            el2.scrollTop = prevTop + (el2.scrollHeight - prevHeight);
                          });
                        }}
                        disabled={messages.length <= msgsVisibleCount}
                        className="mx-auto block rounded-md border border-border px-3 py-1.5 text-xs font-medium text-muted-foreground hover:bg-muted transition-colors disabled:cursor-not-allowed disabled:opacity-40"
                      >
                        {messages.length > msgsVisibleCount
                          ? `Carregar mais (${messages.length - msgsVisibleCount} restantes)`
                          : "Todas as mensagens carregadas"}
                      </button>
                    )}
                    {messages.slice(-msgsVisibleCount).map((m, i, arr) => {
                      const prev = arr[i - 1];
                      const showDate = !prev || toDateKey(m.createdDateTime) !== toDateKey(prev.createdDateTime);
                      return (
                        <div key={m.id} className="flex flex-col gap-3">
                          {showDate && <DateSeparator iso={m.createdDateTime} />}
                          <MessageBubble m={m} meName={account.name} />
                        </div>
                      );
                    })}
                    {pending
                      .filter((p) => p.selKey === selKey)
                      .map((p) => (
                        <PendingBubble
                          key={p.id}
                          p={p}
                          onRetry={() => retryPending(p.id)}
                          onDiscard={() => discardPending(p.id)}
                        />
                      ))}
                  </>
                )}
              </div>
                {hasNewBelow && !atBottom && (
                  <button
                    onClick={scrollToBottom}
                    className="absolute bottom-3 left-1/2 -translate-x-1/2 rounded-full border border-border bg-card px-3 py-1.5 text-xs font-medium text-foreground shadow-md hover:bg-muted"
                  >
                    ↓ Novas mensagens
                  </button>
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
                    className="inline-flex items-center gap-1.5 rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground transition-opacity hover:opacity-90 disabled:opacity-40"
                  >
                    {sending && <Loader2 size={14} className="animate-spin" />}
                    {sending ? "Enviando…" : "Enviar"}
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

function PendingBubble({
  p,
  onRetry,
  onDiscard,
}: {
  p: { text: string; status: "sending" | "error"; error?: string };
  onRetry: () => void;
  onDiscard: () => void;
}) {
  const isError = p.status === "error";
  return (
    <div className="flex justify-end">
      <div
        className={`max-w-[75%] rounded-2xl rounded-br-sm px-4 py-2 text-sm shadow-sm ${
          isError
            ? "border border-destructive/40 bg-destructive/10 text-foreground"
            : "bg-primary/70 text-primary-foreground"
        }`}
      >
        <div className="whitespace-pre-wrap break-words">{p.text}</div>
        {isError ? (
          <div className="mt-1.5 flex items-center gap-2 text-[11px]">
            <AlertCircle size={12} className="text-destructive" />
            <span className="text-destructive">Falha ao enviar</span>
            <button onClick={onRetry} className="ml-auto rounded px-2 py-0.5 font-medium underline hover:no-underline">
              Reenviar
            </button>
            <button
              onClick={onDiscard}
              className="rounded p-0.5 text-muted-foreground hover:text-foreground"
              title="Descartar"
            >
              <X size={12} />
            </button>
          </div>
        ) : (
          <div className="mt-1 flex items-center gap-1.5 text-[10px] text-primary-foreground/80">
            <Loader2 size={10} className="animate-spin" />
            <span>Enviando…</span>
          </div>
        )}
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

function toDateKey(iso: string) {
  try {
    return new Date(iso).toDateString();
  } catch {
    return "";
  }
}

function DateSeparator({ iso }: { iso: string }) {
  const label = (() => {
    try {
      const d = new Date(iso);
      const today = new Date();
      const yesterday = new Date();
      yesterday.setDate(today.getDate() - 1);
      if (d.toDateString() === today.toDateString()) return "Hoje";
      if (d.toDateString() === yesterday.toDateString()) return "Ontem";
      return d.toLocaleDateString("pt-BR", {
        weekday: "long",
        year: "numeric",
        month: "long",
        day: "numeric",
      });
    } catch {
      return "";
    }
  })();
  return (
    <div className="flex items-center justify-center py-1">
      <span className="rounded-full bg-muted px-3 py-1 text-[10px] font-medium text-muted-foreground uppercase tracking-wide">
        {label}
      </span>
    </div>
  );
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

function formatListDate(iso?: string | null) {
  if (!iso) return "";
  try {
    const d = new Date(iso);
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const dDay = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const diffDays = Math.round((today.getTime() - dDay.getTime()) / 86400000);
    if (diffDays === 0) {
      return d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    }
    if (diffDays === 1) return "Ontem";
    if (diffDays < 7) {
      return d.toLocaleDateString("pt-BR", { weekday: "short" }).replace(".", "");
    }
    return d.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" });
  } catch {
    return "";
  }
}

function initials(name: string) {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "?";
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function colorFromName(name: string) {
  let h = 0;
  for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) >>> 0;
  const hue = h % 360;
  return `hsl(${hue} 55% 45%)`;
}

function Avatar({ name, isGroup }: { name: string; isGroup?: boolean }) {
  const bg = colorFromName(name || "?");
  return (
    <div
      className="grid h-10 w-10 shrink-0 place-items-center rounded-full text-xs font-semibold text-white"
      style={{ backgroundColor: bg }}
      aria-hidden
    >
      {isGroup ? (
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" />
          <circle cx="8.5" cy="7" r="4" />
          <path d="M22 21v-2a4 4 0 0 0-3-3.87" />
          <path d="M16 3.13a4 4 0 0 1 0 7.75" />
        </svg>
      ) : (
        initials(name)
      )}
    </div>
  );
}

