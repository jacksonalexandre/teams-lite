import { useState, useEffect, useRef } from "react";
import type { TeamsConfig } from "@/lib/msal";
import { listMessages, listChannelMessages, sendMessage, sendChannelMessage, type GraphMessage } from "@/lib/graph";

type Selection =
  | { kind: "chat"; chatId: string }
  | { kind: "channel"; teamId: string; channelId: string }
  | null;

type Pending = { id: string; selKey: string; text: string; status: "sending" | "error"; error?: string };

export function useMessages(config: TeamsConfig | null, selection: Selection) {
  const [messages, setMessages] = useState<GraphMessage[]>([]);
  const [loading, setLoading] = useState(false);
  const [sending, setSending] = useState(false);
  const [pending, setPending] = useState<Pending[]>([]);
  const [error, setError] = useState<string | null>(null);

  const selKey = selection
    ? selection.kind === "chat" ? `c:${selection.chatId}` : `ch:${selection.teamId}:${selection.channelId}`
    : null;

  useEffect(() => {
    if (!config || !selection) { setMessages([]); return; }
    let cancelled = false;
    const load = async (showLoader: boolean) => {
      if (showLoader) setLoading(true);
      try {
        const msgs = selection.kind === "chat"
          ? await listMessages(config, selection.chatId)
          : await listChannelMessages(config, selection.teamId, selection.channelId);
        if (!cancelled) setMessages(msgs);
      } catch (e) {
        if (!cancelled && showLoader) setError(String(e));
      } finally {
        if (!cancelled && showLoader) setLoading(false);
      }
    };
    load(true);
    const tick = () => { if (typeof document !== "undefined" && document.hidden) return; load(false); };
    const interval = setInterval(tick, 4000);
    const onVisible = () => { if (typeof document !== "undefined" && !document.hidden) load(false); };
    document.addEventListener("visibilitychange", onVisible);
    return () => { cancelled = true; clearInterval(interval); document.removeEventListener("visibilitychange", onVisible); };
  }, [config, selection]);

  async function doSend(p: Pending) {
    if (!config || !selection) return;
    setPending((prev) => prev.map((x) => (x.id === p.id ? { ...x, status: "sending", error: undefined } : x)));
    try {
      if (selection.kind === "chat") {
        await sendMessage(config, selection.chatId, p.text);
        setMessages(await listMessages(config, selection.chatId));
      } else {
        await sendChannelMessage(config, selection.teamId, selection.channelId, p.text);
        setMessages(await listChannelMessages(config, selection.teamId, selection.channelId));
      }
      setPending((prev) => prev.filter((x) => x.id !== p.id));
    } catch (e) {
      setPending((prev) => prev.map((x) => (x.id === p.id ? { ...x, status: "error", error: String(e) } : x)));
    }
  }

  async function send(text: string) {
    if (!config || !selection || !text.trim() || sending) return;
    setSending(true);
    const p: Pending = {
      id: `p-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      selKey: selKey!,
      text,
      status: "sending",
    };
    setPending((prev) => [...prev, p]);
    try { await doSend(p); } finally { setSending(false); }
  }

  function retryPending(id: string) {
    const p = pending.find((x) => x.id === id);
    if (p) doSend(p);
  }

  function discardPending(id: string) {
    setPending((prev) => prev.filter((x) => x.id !== id));
  }

  return { messages, loading, sending, pending, error, selKey, send, retryPending, discardPending };
}

export function useScrollBehavior(messages: GraphMessage[], selKey: string | null) {
  const scrollRef = useRef<HTMLDivElement>(null);
  const [atBottom, setAtBottom] = useState(true);
  const [hasNewBelow, setHasNewBelow] = useState(false);
  const prevSelKeyRef = useRef<string | null>(null);
  const prevLastMsgIdRef = useRef<string | null>(null);

  useEffect(() => { setHasNewBelow(false); setAtBottom(true); }, [selKey]);

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

  return { scrollRef, atBottom, hasNewBelow, scrollToBottom };
}
