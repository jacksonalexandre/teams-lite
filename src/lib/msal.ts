import {
  PublicClientApplication,
  type Configuration,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

const POPUP_REDIRECT_PATH = "/auth-callback";

export type TeamsConfig = {
  clientId: string;
  tenantId: string; // "common" | "organizations" | tenant GUID
};

export const GRAPH_SCOPES = [
  "User.Read",
  "User.ReadBasic.All",
  "Chat.ReadWrite",
  "ChatMessage.Send",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "ChannelMessage.Send",
];

let pca: PublicClientApplication | null = null;
let initialized = false;

function getMsal(cfg: TeamsConfig) {
  if (!pca) {
    const config: Configuration = {
      auth: {
        clientId: cfg.clientId,
        authority: `https://login.microsoftonline.com/${cfg.tenantId || "common"}`,
        redirectUri: `${window.location.origin}${POPUP_REDIRECT_PATH}`,
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
