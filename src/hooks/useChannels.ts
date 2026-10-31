import { useState, useEffect } from "react";
import type { TeamsConfig } from "@/lib/msal";
import {
  listJoinedTeams,
  listChannels,
  getChannelLastMessageDate,
  mapLimit,
  type GraphTeam,
  type GraphChannel,
} from "@/lib/graph";

export type ChannelEntry = { team: GraphTeam; channel: GraphChannel; lastDate: string | null };

// Building the list costs one request per team plus one per channel, so it is
// kept across tab switches and only rebuilt after the TTL.
const CACHE_TTL = 5 * 60 * 1000;
const MAX_CONCURRENT = 4;
let cache: { accountId: string; at: number; list: ChannelEntry[] } | null = null;

function cachedFor(accountId: string | undefined) {
  if (!cache || !accountId || cache.accountId !== accountId) return null;
  return cache;
}

export function useChannels(config: TeamsConfig | null, account: { oid?: string } | null, active: boolean) {
  const [channelList, setChannelList] = useState<ChannelEntry[]>(() => cachedFor(account?.oid)?.list ?? []);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!config || !account || !active) return;
    const accountId = account.oid ?? "";
    const hit = cachedFor(accountId);
    if (hit && Date.now() - hit.at < CACHE_TTL) return;
    let cancelled = false;
    (async () => {
      if (!hit) setLoading(true);
      try {
        const teams = await listJoinedTeams(config);
        const perTeam = await mapLimit(teams, MAX_CONCURRENT, (t) =>
          listChannels(config, t.id)
            .then((chs) => chs.map((channel) => ({ team: t, channel })))
            .catch(() => [] as Array<{ team: GraphTeam; channel: GraphChannel }>),
        );
        const flat = perTeam.flat();
        const withDates = await mapLimit(flat, MAX_CONCURRENT, async (item) => ({
          ...item,
          lastDate: await getChannelLastMessageDate(config, item.team.id, item.channel.id),
        }));
        withDates.sort((a, b) => {
          const ad = a.lastDate ? new Date(a.lastDate).getTime() : 0;
          const bd = b.lastDate ? new Date(b.lastDate).getTime() : 0;
          return bd - ad;
        });
        const list = withDates.filter((x) => x.lastDate).slice(0, 12);
        cache = { accountId, at: Date.now(), list };
        if (!cancelled) setChannelList(list);
      } catch (e) {
        if (!cancelled) setError(String(e));
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [config, account, active]);

  return { channelList, loading, error, setChannelList };
}
