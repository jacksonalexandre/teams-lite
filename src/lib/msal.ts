import { PublicClientApplication, type Configuration, InteractionRequiredAuthError } from "@azure/msal-browser";

const STORAGE_KEY = "teamslite.config";

export type TeamsConfig = {
  clientId: string;
  tenantId: string; // "common" | "organizations" | tenant GUID
};

export function loadConfig(): TeamsConfig | null {
  if (typeof window === "undefined") return null;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? (JSON.parse(raw) as TeamsConfig) : null;
  } catch {
    return null;
  }
}

export function saveConfig(cfg: TeamsConfig) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
}

export function clearConfig() {
  localStorage.removeItem(STORAGE_KEY);
}

export const GRAPH_SCOPES = [
  "User.Read",
  "Chat.ReadWrite",
  "ChatMessage.Send",
];

let pca: PublicClientApplication | null = null;
let initialized = false;

export function getMsal(cfg: TeamsConfig) {
  if (!pca) {
    const config: Configuration = {
      auth: {
        clientId: cfg.clientId,
        authority: `https://login.microsoftonline.com/${cfg.tenantId || "common"}`,
        redirectUri: `${window.location.origin}/auth-callback.html`,
        postLogoutRedirectUri: window.location.origin,
      },
      cache: {
        cacheLocation: "localStorage",
      },
    };
    pca = new PublicClientApplication(config);
  }
  return pca;
}

export async function ensureInit(cfg: TeamsConfig) {
  const app = getMsal(cfg);
  if (!initialized) {
    await app.initialize();
    await app.handleRedirectPromise();
    initialized = true;
  }
  return app;
}

export async function getAccessToken(cfg: TeamsConfig): Promise<string> {
  const app = await ensureInit(cfg);
  const accounts = app.getAllAccounts();
  if (accounts.length === 0) throw new Error("Not signed in");
  try {
    const res = await app.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account: accounts[0],
    });
    return res.accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      const res = await app.acquireTokenPopup({ scopes: GRAPH_SCOPES });
      return res.accessToken;
    }
    throw e;
  }
}

export async function signIn(cfg: TeamsConfig) {
  const app = await ensureInit(cfg);
  const res = await app.loginPopup({ scopes: GRAPH_SCOPES, prompt: "select_account" });
  app.setActiveAccount(res.account);
  return res.account;
}

export async function signOut(cfg: TeamsConfig) {
  const app = await ensureInit(cfg);
  const account = app.getAllAccounts()[0];
  if (account) {
    await app.logoutPopup({ account });
  }
}

export async function getCurrentAccount(cfg: TeamsConfig) {
  const app = await ensureInit(cfg);
  return app.getAllAccounts()[0] ?? null;
}
