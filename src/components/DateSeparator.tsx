export function DateSeparator({ iso }: { iso: string }) {
  const label = (() => {
    try {
      const d = new Date(iso);
      const today = new Date();
      const yesterday = new Date();
      yesterday.setDate(today.getDate() - 1);
      if (d.toDateString() === today.toDateString()) return "Hoje";
      if (d.toDateString() === yesterday.toDateString()) return "Ontem";
      return d.toLocaleDateString("pt-BR", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
    } catch {
      return "";
    }
  })();
  return (
    <div className="flex items-center justify-center py-1">
      <span className="rounded-full bg-muted px-3 py-1 text-[10px] font-medium text-muted-foreground uppercase tracking-wide">
        {label}
      </span>
    </div>
  );
}
