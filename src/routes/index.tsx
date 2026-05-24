import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { useServerFn } from "@tanstack/react-start";
import { Loader2, Video } from "lucide-react";
import { signIn, signOut, ensureInit, getCurrentAccount, type TeamsConfig } from "@/lib/msal";
import { getTeamsConfig } from "@/lib/config.functions";
import { chatTitle, buildTeamsVideoLink } from "@/lib/graph";
import { useChats } from "@/hooks/useChats";
import { useChannels } from "@/hooks/useChannels";
import { useMessages, useScrollBehavior } from "@/hooks/useMessages";
import { ChatItem } from "@/components/ChatItem";
import { ChannelItem } from "@/components/ChannelItem";
import { MessageBubble, PendingBubble } from "@/components/MessageBubble";
import { DateSeparator } from "@/components/DateSeparator";
import { Avatar } from "@/components/Avatar";
import { toDateKey } from "@/lib/format";

export const Route = createFileRoute("/")({
  component: TeamsLite,
});

type Mode = "chats" | "channels";
type Selection =
  | { kind: "chat"; chatId: string }
  | { kind: "channel"; teamId: string; channelId: string }
  | null;

function EmptyHint({ text }: { text: string }) {
  return <div className="px-4 py-10 text-center text-xs text-muted-foreground">{text}</div>;
}

function TeamsLite() {
  const fetchConfig = useServerFn(getTeamsConfig);
  const [config, setConfig] = useState<TeamsConfig | null>(null);
  const [account, setAccount] = useState<{ name?: string; username: string; oid?: string; tid?: string } | null>(null);
  const [mode, setMode] = useState<Mode>("chats");
  const [selection, setSelection] = useState<Selection>(null);
  const [draft, setDraft] = useState("");
  const [msgsVisibleCount, setMsgsVisibleCount] = useState(7);
  const [initError, setInitError] = useState<string | null>(null);

  const { chats, loading: loadingChats, hidingId, error: chatsError, toggleHide } = useChats(config, account, mode === "chats");
  const { channelList, loading: loadingChannels, error: channelsError } = useChannels(config, account, mode === "channels");
  const { messages, loading: loadingMsgs, sending, pending, error: msgsError, selKey, send, retryPending, discardPending } = useMessages(config, selection);
  const { scrollRef, atBottom, hasNewBelow, scrollToBottom } = useScrollBehavior(messages, selKey);

  const error = initError ?? chatsError ?? channelsError ?? msgsError;

  useEffect(() => {
    fetchConfig()
      .then((cfg) => {
        setConfig(cfg);
        return ensureInit(cfg).then(() => getCurrentAccount(cfg));
      })
      .then((acc) => {
        if (acc) {
          const claims = (acc as any).idTokenClaims ?? {};
          setAccount({ name: acc.name, username: acc.username, oid: claims.oid || acc.localAccountId, tid: claims.tid || (acc as any).tenantId });
        }
      })
      .catch((e) => setInitError(String(e)));
  }, []); // eslint-disable-line

  useEffect(() => { setMsgsVisibleCount(7); }, [selKey]);

  const meId = useMemo(() => account?.oid, [account]);
  const filteredChats = useMemo(() => chats.filter((c) => !c.viewpoint?.isHidden), [chats]);

  const headerTitle = useMemo(() => {
    if (!selection) return "";
    if (selection.kind === "chat") {
      const c = chats.find((x) => x.id === selection.chatId);
      return c ? chatTitle(c, meId, account?.name) : "";
    }
    const item = channelList.find((x) => x.team.id === selection.teamId && x.channel.id === selection.channelId);
    return item ? `${item.team.displayName} · ${item.channel.displayName}` : "";
  }, [selection, chats, channelList, meId, account]);

  async function handleSignIn() {
    if (!config) return;
    setInitError(null);
    try {
      const acc = await signIn(config);
      const claims = (acc as any).idTokenClaims ?? {};
      setAccount({ name: acc.name, username: acc.username, oid: claims.oid || acc.localAccountId, tid: claims.tid || (acc as any).tenantId });
    } catch (e) {
      setInitError(String(e));
    }
  }

  async function handleSignOut() {
    if (!config) return;
    try {
      await signOut(config);
      setAccount(null);
      setSelection(null);
    } catch (e) {
      setInitError(String(e));
    }
  }

  async function handleSend() {
    if (!draft.trim()) return;
    const text = draft.trim();
    setDraft("");
    await send(text);
  }

  return (
    <div className="flex h-screen flex-col bg-background text-foreground">
      <header className="flex items-center justify-between border-b border-border px-5 py-3">
        <div className="flex items-center gap-2">
          <div className="grid h-7 w-7 place-items-center rounded-md bg-primary text-primary-foreground text-sm font-semibold">T</div>
          <h1 className="text-base font-semibold tracking-tight">Teams Lite</h1>
        </div>
        <div className="flex items-center gap-2 text-sm">
          {account ? (
            <>
              <span className="text-muted-foreground hidden sm:inline">{account.name ?? account.username}</span>
              <button onClick={handleSignOut} className="rounded-md border border-border px-3 py-1.5 text-xs font-medium hover:bg-muted">
                Sair
              </button>
            </>
          ) : (
            config && (
              <button onClick={handleSignIn} className="rounded-md bg-primary px-3 py-1.5 text-xs font-medium text-primary-foreground hover:opacity-90">
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
            <button onClick={() => setInitError(null)} className="shrink-0 underline">fechar</button>
          </div>
        </div>
      )}

      <div className="flex min-h-0 flex-1">
        <aside className="flex w-80 flex-col border-r border-border bg-card">
          <div className="flex border-b border-border">
            <button
              onClick={() => setMode("chats")}
              className={`flex-1 px-4 py-2.5 text-xs font-medium uppercase tracking-wider transition-colors ${mode === "chats" ? "bg-muted text-foreground" : "text-muted-foreground hover:bg-muted/60"}`}
            >
              Chats
            </button>
            <button
              onClick={() => setMode("channels")}
              className={`flex-1 px-4 py-2.5 text-xs font-medium uppercase tracking-wider transition-colors ${mode === "channels" ? "bg-muted text-foreground" : "text-muted-foreground hover:bg-muted/60"}`}
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
                filteredChats.map((c) => (
                  <ChatItem
                    key={c.id}
                    chat={c}
                    active={selection?.kind === "chat" && selection.chatId === c.id}
                    meId={meId}
                    meName={account.name}
                    hidingId={hidingId}
                    cfg={config}
                    onSelect={() => setSelection({ kind: "chat", chatId: c.id })}
                    onToggleHide={async (id, hidden) => {
                      const ok = await toggleHide(id, hidden);
                      if (ok && selection?.kind === "chat" && selection.chatId === id) setSelection(null);
                    }}
                  />
                ))
              )
            ) : loadingChannels ? (
              <EmptyHint text="Carregando canais…" />
            ) : channelList.length === 0 ? (
              <EmptyHint text="Nenhum canal encontrado." />
            ) : (
              channelList.map(({ team, channel, lastDate }) => (
                <ChannelItem
                  key={`${team.id}:${channel.id}`}
                  team={team}
                  channel={channel}
                  lastDate={lastDate}
                  active={selection?.kind === "channel" && selection.teamId === team.id && selection.channelId === channel.id}
                  onSelect={() => setSelection({ kind: "channel", teamId: team.id, channelId: channel.id })}
                />
              ))
            )}
          </div>
        </aside>

        <main className="flex min-w-0 flex-1 flex-col">
          {!selection || !account ? (
            <div className="grid flex-1 place-items-center text-sm text-muted-foreground">
              {account
                ? mode === "chats" ? "Selecione uma conversa" : "Selecione um canal"
                : config ? "Não autenticado" : "Carregando configuração…"}
            </div>
          ) : (
            <>
              {headerTitle && (
                <div className="flex items-center gap-3 border-b border-border bg-card px-6 py-2.5 text-sm font-medium">
                  {(() => {
                    if (selection?.kind === "chat") {
                      const c = chats.find((x) => x.id === selection.chatId);
                      const isGroup = c?.chatType === "group" || c?.chatType === "meeting";
                      const otherId = c && !isGroup ? (c.members ?? []).find((m) => m.userId && m.userId !== meId)?.userId ?? undefined : undefined;
                      return <Avatar name={headerTitle} isGroup={isGroup} userId={otherId} cfg={config} size={32} />;
                    }
                    return <Avatar name={headerTitle} isGroup size={32} />;
                  })()}
                  <span>{headerTitle}</span>
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
                            <MessageBubble m={m} meName={account.name} cfg={config} />
                          </div>
                        );
                      })}
                      {pending
                        .filter((p) => p.selKey === selKey)
                        .map((p) => (
                          <PendingBubble key={p.id} p={p} onRetry={() => retryPending(p.id)} onDiscard={() => discardPending(p.id)} />
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
                    onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSend(); } }}
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
