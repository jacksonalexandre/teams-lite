import { useState, useEffect } from "react";
import type { TeamsConfig } from "@/lib/msal";
import { getUserPhotoUrl } from "@/lib/graph";

function initials(name: string) {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "?";
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function colorFromName(name: string) {
  let h = 0;
  for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) >>> 0;
  return `hsl(${h % 360} 55% 45%)`;
}

function useUserPhoto(cfg: TeamsConfig | null, userId?: string | null) {
  const [url, setUrl] = useState<string | null>(null);
  useEffect(() => {
    let cancelled = false;
    if (!cfg || !userId) { setUrl(null); return; }
    getUserPhotoUrl(cfg, userId).then((u) => { if (!cancelled) setUrl(u); });
    return () => { cancelled = true; };
  }, [cfg, userId]);
  return url;
}

export function Avatar({
  name,
  isGroup,
  userId,
  cfg,
  size = 40,
}: {
  name: string;
  isGroup?: boolean;
  userId?: string;
  cfg?: TeamsConfig | null;
  size?: number;
}) {
  const bg = colorFromName(name || "?");
  const photo = useUserPhoto(cfg ?? null, isGroup ? null : userId);
  const dim = { width: size, height: size };
  const textSize = size <= 28 ? "text-[10px]" : "text-xs";

  if (photo) {
    return (
      <img src={photo} alt="" aria-hidden style={dim} className="shrink-0 rounded-full object-cover" />
    );
  }
  return (
    <div
      className={`grid shrink-0 place-items-center rounded-full font-semibold text-white ${textSize}`}
      style={{ backgroundColor: bg, ...dim }}
      aria-hidden
    >
      {isGroup ? (
        <svg width={size * 0.45} height={size * 0.45} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
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
