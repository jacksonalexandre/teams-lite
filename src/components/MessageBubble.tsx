import { useState, useEffect, type CSSProperties } from "react";
import { AlertCircle, Loader2, X } from "lucide-react";
import type { GraphMessage } from "@/lib/graph";
import type { TeamsConfig } from "@/lib/msal";
import { getHostedContentUrl } from "@/lib/graph";
import { formatDateTime } from "@/lib/format";
import { Avatar } from "./Avatar";

type ParsedImage = { src: string; width?: number; height?: number; hosted: boolean };

function parseMessageHtml(html: string): { text: string; images: ParsedImage[] } {
  if (typeof window === "undefined") return { text: html.replace(/<[^>]*>/g, ""), images: [] };
  const doc = new DOMParser().parseFromString(html, "text/html");
  doc.querySelectorAll("attachment").forEach((n) => n.remove());
  const images: ParsedImage[] = [];
  doc.querySelectorAll("img").forEach((img) => {
    const src = img.getAttribute("src") || "";
    if (!src) return;
    const hosted = /graph\.microsoft\.com\/.+\/hostedContents\//i.test(src);
    const w = parseInt(img.getAttribute("width") || "", 10);
    const h = parseInt(img.getAttribute("height") || "", 10);
    images.push({ src, width: Number.isFinite(w) ? w : undefined, height: Number.isFinite(h) ? h : undefined, hosted });
    img.remove();
  });
  return { text: (doc.body.textContent || "").trim(), images };
}

function MessageImage({ img, cfg }: { img: ParsedImage; cfg?: TeamsConfig | null }) {
  const [url, setUrl] = useState<string | null>(img.hosted ? null : img.src);
  const [failed, setFailed] = useState(false);

  useEffect(() => {
    if (!img.hosted || !cfg) return;
    let cancelled = false;
    (async () => {
      const u = await getHostedContentUrl(cfg, img.src);
      if (cancelled) return;
      if (u) setUrl(u); else setFailed(true);
    })();
    return () => { cancelled = true; };
  }, [img.hosted, img.src, cfg]);

  const style: CSSProperties = { maxWidth: 320, maxHeight: 320, width: img.width ? Math.min(img.width, 320) : undefined, height: "auto" };

  if (failed) return (
    <div className="rounded-lg border border-border bg-muted/40 px-2.5 py-2 text-[11px] text-muted-foreground">
      Não foi possível carregar a imagem.
    </div>
  );
  if (!url) return (
    <div className="flex items-center justify-center rounded-lg border border-border bg-muted/40" style={{ width: style.width ?? 240, height: 140 }}>
      <Loader2 size={16} className="animate-spin text-muted-foreground" />
    </div>
  );
  return (
    <a href={url} target="_blank" rel="noopener noreferrer" className="block">
      <img src={url} alt="" style={style} className="rounded-lg border border-border object-contain" onError={() => setFailed(true)} />
    </a>
  );
}

function loopLabel(a: { contentType: string; content?: string | null; name?: string | null }) {
  const ct = (a.contentType || "").toLowerCase();
  let componentType = "";
  try {
    if (a.content) {
      const parsed = JSON.parse(a.content);
      componentType = String(parsed?.componentType || parsed?.LoopComponentType || "").toLowerCase();
    }
  } catch { /* ignore */ }
  const isLoop = ct.includes("fluidembedcard") || ct.includes("loopcomponent") || ct.includes("fluid") || componentType.startsWith("fluid");
  if (!isLoop) return null;
  const map: Record<string, string> = { fluidlist: "Lista", fluidtable: "Tabela", fluidtask: "Tarefas", fluidparagraph: "Parágrafo", fluidchecklist: "Checklist" };
  const kind = map[componentType] || "Componente do Loop";
  return { kind, name: a.name || kind };
}

function AttachmentChip({ a, mine }: { a: NonNullable<GraphMessage["attachments"]>[number]; mine: boolean }) {
  const loop = loopLabel(a);
  const baseCls = mine ? "border-primary-foreground/30 bg-primary-foreground/10 hover:bg-primary-foreground/20" : "border-border bg-muted/40 hover:bg-muted";
  const subCls = mine ? "text-primary-foreground/70" : "text-muted-foreground";

  if (loop) return (
    <a href={a.contentUrl || "#"} target="_blank" rel="noopener noreferrer" className={`flex items-start gap-2 rounded-lg border px-2.5 py-2 transition-colors ${baseCls}`}>
      <div className="mt-0.5 flex h-7 w-7 shrink-0 items-center justify-center rounded-md bg-[#7B83EB]/15 text-[#7B83EB]">
        <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor" aria-hidden>
          <path d="M12 2a10 10 0 1 0 10 10A10.01 10.01 0 0 0 12 2Zm4.5 11h-3v3a1.5 1.5 0 0 1-3 0v-3h-3a1.5 1.5 0 0 1 0-3h3V7a1.5 1.5 0 0 1 3 0v3h3a1.5 1.5 0 0 1 0 3Z" />
        </svg>
      </div>
      <div className="min-w-0 flex-1">
        <div className="truncate text-[13px] font-medium">{loop.name}</div>
        <div className={`text-[11px] ${subCls}`}>{loop.kind} · Abrir no Microsoft Loop ↗</div>
      </div>
    </a>
  );

  const ct = (a.contentType || "").toLowerCase();
  const label = ct.includes("reference") ? "Arquivo" : "Anexo";
  return (
    <a href={a.contentUrl || "#"} target="_blank" rel="noopener noreferrer" className={`flex items-start gap-2 rounded-lg border px-2.5 py-2 transition-colors ${baseCls}`}>
      <div className="mt-0.5 flex h-7 w-7 shrink-0 items-center justify-center rounded-md bg-foreground/10">
        <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden>
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
          <polyline points="14 2 14 8 20 8" />
        </svg>
      </div>
      <div className="min-w-0 flex-1">
        <div className="truncate text-[13px] font-medium">{a.name || label}</div>
        <div className={`text-[11px] ${subCls}`}>{a.contentUrl ? "Abrir ↗" : label}</div>
      </div>
    </a>
  );
}

export function MessageBubble({ m, meName, cfg }: { m: GraphMessage; meName?: string; cfg?: TeamsConfig | null }) {
  const author = m.from?.user?.displayName ?? "Sistema";
  const mine = !!meName && author === meName;
  const { text, images } = m.body.contentType === "html" ? parseMessageHtml(m.body.content) : { text: m.body.content, images: [] as ParsedImage[] };
  const attachments = (m.attachments ?? []).filter((a) => !!a);
  if (!text.trim() && attachments.length === 0 && images.length === 0) return null;
  const fromUserId = m.from?.user?.id ?? undefined;
  return (
    <div className={`flex items-end gap-2 ${mine ? "justify-end" : "justify-start"}`}>
      {!mine && <Avatar name={author} userId={fromUserId} cfg={cfg} size={28} />}
      <div className={`max-w-[75%] rounded-2xl px-4 py-2 text-sm shadow-sm ${mine ? "bg-primary text-primary-foreground rounded-br-sm" : "bg-card text-card-foreground border border-border rounded-bl-sm"}`}>
        <div className={`mb-1 text-[11px] ${mine ? "text-primary-foreground/70" : "text-muted-foreground"}`}>
          {mine ? formatDateTime(m.createdDateTime) : (
            <>
              <span className="font-medium opacity-90">{author}</span>
              <span className="opacity-60"> · {formatDateTime(m.createdDateTime)}</span>
            </>
          )}
        </div>
        {text.trim() && <div className="whitespace-pre-wrap break-words">{text}</div>}
        {images.length > 0 && (
          <div className="mt-2 flex flex-col gap-1.5">
            {images.map((img, i) => <MessageImage key={`${img.src}-${i}`} img={img} cfg={cfg} />)}
          </div>
        )}
        {attachments.length > 0 && (
          <div className="mt-2 flex flex-col gap-1.5">
            {attachments.map((a) => <AttachmentChip key={a.id} a={a} mine={mine} />)}
          </div>
        )}
      </div>
    </div>
  );
}

export function PendingBubble({
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
      <div className={`max-w-[75%] rounded-2xl rounded-br-sm px-4 py-2 text-sm shadow-sm ${isError ? "border border-destructive/40 bg-destructive/10 text-foreground" : "bg-primary/70 text-primary-foreground"}`}>
        <div className="whitespace-pre-wrap break-words">{p.text}</div>
        {isError ? (
          <div className="mt-1.5 flex items-center gap-2 text-[11px]">
            <AlertCircle size={12} className="text-destructive" />
            <span className="text-destructive">Falha ao enviar</span>
            <button onClick={onRetry} className="ml-auto rounded px-2 py-0.5 font-medium underline hover:no-underline">Reenviar</button>
            <button onClick={onDiscard} className="rounded p-0.5 text-muted-foreground hover:text-foreground" title="Descartar"><X size={12} /></button>
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
