import { Eye, EyeOff } from "lucide-react";
import type { GraphChat } from "@/lib/graph";
import type { TeamsConfig } from "@/lib/msal";
import { Avatar } from "./Avatar";
import { stripHtml, formatListDate } from "@/lib/format";
import { chatTitle } from "@/lib/graph";

export function ChatItem({
  chat,
  active,
  meId,
  meName,
  hidingId,
  cfg,
  onSelect,
  onToggleHide,
}: {
  chat: GraphChat;
  active: boolean;
  meId?: string;
  meName?: string;
  hidingId: string | null;
  cfg?: TeamsConfig | null;
  onSelect: () => void;
  onToggleHide: (id: string, hidden: boolean) => void;
}) {
  const isHidden = !!chat.viewpoint?.isHidden;
  const isBusy = hidingId === chat.id;
  const preview = chat.lastMessagePreview;
  const previewAuthor = preview?.from?.user?.displayName;
  const previewIsMe = !!meName && previewAuthor === meName;
  const rawPreview = preview?.body?.content ?? "";
  const previewStripped = preview?.body?.content
    ? preview.body.contentType === "html" ? stripHtml(preview.body.content) : preview.body.content
    : "";
  const previewText = previewStripped.trim() || (/<img\b/i.test(rawPreview) ? "📷 Imagem" : "");
  const previewPrefix = previewIsMe ? "Você: " : previewAuthor ? `${previewAuthor.split(" ")[0]}: ` : "";
  const isGroup = chat.chatType === "group" || chat.chatType === "meeting";
  const otherUserId = !isGroup ? (chat.members ?? []).find((m) => m.userId && m.userId !== meId)?.userId ?? undefined : undefined;
  const title = chatTitle(chat, meId, meName);

  return (
    <div
      onClick={onSelect}
      className={`group flex w-full cursor-pointer items-center gap-3 border-b border-border px-3 py-2.5 text-sm transition-colors ${active ? "bg-muted" : "hover:bg-muted/60"}`}
      title={title}
    >
      <Avatar name={title} isGroup={isGroup} userId={otherUserId} cfg={cfg} />
      <div className="flex min-w-0 flex-1 flex-col">
        <div className="flex items-baseline justify-between gap-2">
          <span className="truncate text-left font-semibold text-foreground">{title}</span>
          <span className="shrink-0 text-[11px] text-muted-foreground">
            {formatListDate(preview?.createdDateTime ?? chat.lastUpdatedDateTime)}
          </span>
        </div>
        <div className="flex items-center justify-between gap-2">
          <span className="truncate text-left text-xs text-muted-foreground">
            {previewPrefix}{previewText || (isGroup ? "Grupo" : "")}
          </span>
          <button
            onClick={(e) => { e.stopPropagation(); onToggleHide(chat.id, isHidden); }}
            disabled={isBusy}
            title={isHidden ? "Mostrar no Teams" : "Ocultar no Teams"}
            className="opacity-0 group-hover:opacity-100 focus:opacity-100 shrink-0 rounded p-1 text-muted-foreground hover:bg-background hover:text-foreground transition-opacity disabled:opacity-40"
          >
            {isHidden ? <Eye size={14} /> : <EyeOff size={14} />}
          </button>
        </div>
      </div>
    </div>
  );
}
