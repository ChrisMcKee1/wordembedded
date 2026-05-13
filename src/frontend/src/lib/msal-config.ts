import { LogLevel, type Configuration } from "@azure/msal-browser";
import { apiScopes } from "./auth-scopes";

const tenantId = process.env.NEXT_PUBLIC_AAD_TENANT_ID ?? "common";
const clientId = process.env.NEXT_PUBLIC_AAD_CLIENT_ID ?? "";
const browserOrigin = typeof window === "undefined" ? "http://localhost:3000" : window.location.origin;

export const msalConfig: Configuration = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: `${browserOrigin}/auth/callback`,
    postLogoutRedirectUri: browserOrigin,
  },
  cache: { cacheLocation: "sessionStorage" },
  system: {
    allowPlatformBroker: false,
    loggerOptions: { logLevel: LogLevel.Warning },
  },
};

export { apiScopes };
export const loginScopes = ["openid", "profile", "User.Read", ...apiScopes];
