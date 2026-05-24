import { useState, useEffect } from "react";
import type { TeamsConfig } from "@/lib/msal";
import {
  listJoinedTeams,
  listChannels,
  getChannelLastMessageDate,
  type GraphTeam,
  type GraphChannel,
} from "@/lib/graph";

export type ChannelEntry = { team: GraphTeam; channel: GraphChannel; lastDate: string | null };

export function useChannels(config: TeamsConfig | null, account: unknown, active: boolean) {
  const [channelList, setChannelList] = useState<ChannelEntry[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!config || !account || !active) return;
    let cancelled = false;
    setLoading(true);
    (async () => {
      try {
        const teams = await listJoinedTeams(config);
        const perTeam = await Promise.all(
          teams.map((t) =>
            listChannels(config, t.id)
              .then((chs) => chs.map((channel) => ({ team: t, channel })))
              .catch(() => [] as Array<{ team: GraphTeam; channel: GraphChannel }>),
          ),
        );
        const flat = perTeam.flat();
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
        if (!cancelled) setChannelList(withDates.filter((x) => x.lastDate).slice(0, 12));
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
