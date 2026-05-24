import type { GraphTeam, GraphChannel } from "@/lib/graph";
import { Avatar } from "./Avatar";
import { formatListDate } from "@/lib/format";

export function ChannelItem({
  team,
  channel,
  lastDate,
  active,
  onSelect,
}: {
  team: GraphTeam;
  channel: GraphChannel;
  lastDate: string | null;
  active: boolean;
  onSelect: () => void;
}) {
  const fullTitle = `${team.displayName} · #${channel.displayName}`;
  return (
    <div
      onClick={onSelect}
      title={fullTitle}
      className={`flex w-full cursor-pointer items-center gap-3 border-b border-border px-3 py-2.5 text-sm transition-colors ${
        active ? "bg-muted" : "hover:bg-muted/60"
      }`}
    >
      <Avatar name={team.displayName} isGroup />
      <div className="flex min-w-0 flex-1 flex-col">
        <div className="flex items-baseline justify-between gap-2">
          <span className="truncate text-left font-semibold text-foreground">{team.displayName}</span>
          <span className="shrink-0 text-[11px] text-muted-foreground">{formatListDate(lastDate)}</span>
        </div>
        <span className="truncate text-left text-xs text-muted-foreground">#{channel.displayName}</span>
      </div>
    </div>
  );
}
