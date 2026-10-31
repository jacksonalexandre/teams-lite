import { useState } from "react";
import { Loader2 } from "lucide-react";

// The draft lives here, not in the page, so typing re-renders only the
// composer instead of the chat list and every message.
export function Composer({ sending, onSend }: { sending: boolean; onSend: (text: string) => Promise<void> }) {
  const [draft, setDraft] = useState("");

  async function handleSend() {
    if (!draft.trim()) return;
    const text = draft.trim();
    setDraft("");
    await onSend(text);
  }

  return (
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
  );
}
