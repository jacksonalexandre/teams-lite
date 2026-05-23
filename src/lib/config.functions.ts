import { createServerFn } from "@tanstack/react-start";
import type { TeamsConfig } from "@/lib/msal";

export const getTeamsConfig = createServerFn({ method: "GET" }).handler(
  async (): Promise<TeamsConfig> => {
    const clientId = process.env.TEAMS_CLIENT_ID;
    const tenantId = process.env.TEAMS_TENANT_ID;
    if (!clientId) throw new Error("TEAMS_CLIENT_ID não configurado");
    return { clientId, tenantId: tenantId || "common" };
  }
);
