import { useState, useEffect } from "react";
import type { TeamsConfig } from "@/lib/msal";
import { listChats, hideChat, type GraphChat } from "@/lib/graph";

export function useChats(config: TeamsConfig | null, account: { oid?: string; tid?: string } | null, active: boolean) {
  const [chats, setChats] = useState<GraphChat[]>([]);
  const [loading, setLoading] = useState(false);
  const [hidingId, setHidingId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!config || !account || !active) return;
    let cancelled = false;
    const load = async (showLoader: boolean) => {
      if (showLoader) setLoading(true);
      try {
        const cs = await listChats(config);
        if (!cancelled) setChats(cs);
      } catch (e) {
        if (!cancelled && showLoader) setError(String(e));
      } finally {
        if (!cancelled && showLoader) setLoading(false);
      }
    };
    load(true);
    const tick = () => { if (typeof document !== "undefined" && document.hidden) return; load(false); };
    const interval = setInterval(tick, 15000);
    const onVisible = () => { if (typeof document !== "undefined" && !document.hidden) load(false); };
    document.addEventListener("visibilitychange", onVisible);
    return () => { cancelled = true; clearInterval(interval); document.removeEventListener("visibilitychange", onVisible); };
  }, [config, account, active]);

  async function toggleHide(chatId: string, currentlyHidden: boolean): Promise<boolean> {
    if (!config || hidingId || !account?.oid || !account?.tid) return false;
    setHidingId(chatId);
    try {
      await hideChat(config, chatId, !currentlyHidden, { id: account.oid, tenantId: account.tid });
      setChats((prev) =>
        prev.map((c) =>
          c.id === chatId ? { ...c, viewpoint: { ...(c.viewpoint ?? {}), isHidden: !currentlyHidden } } : c,
        ),
      );
      return true;
    } catch (e) {
      setError(String(e));
      return false;
    } finally {
      setHidingId(null);
    }
  }

  return { chats, loading, hidingId, error, toggleHide, setChats };
}
