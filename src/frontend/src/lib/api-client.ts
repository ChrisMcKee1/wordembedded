import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { apiScopes } from "./auth-scopes";
import { initializeMsal, msalInstance } from "./msal-instance";

const baseUrl = process.env.NEXT_PUBLIC_API_BASE_URL ?? "https://localhost:7071";

async function getToken(): Promise<string> {
  await initializeMsal();

  const account = msalInstance.getActiveAccount() ?? msalInstance.getAllAccounts()[0];
  if (!account) throw new Error("Not signed in");
  if (apiScopes.length === 0) throw new Error("NEXT_PUBLIC_API_SCOPE is not configured");

  try {
    const result = await msalInstance.acquireTokenSilent({ scopes: apiScopes, account });
    return result.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const result = await msalInstance.acquireTokenPopup({ scopes: apiScopes, account });
      return result.accessToken;
    }

    throw error;
  }
}

export async function apiFetch<T = unknown>(path: string, init?: RequestInit): Promise<T> {
  const token = await getToken();
  const headers = new Headers(init?.headers);
  headers.set("Authorization", `Bearer ${token}`);

  const response = await fetch(`${baseUrl}${path}`, { ...init, headers });
  if (!response.ok) {
    const text = await response.text().catch(() => "");
    throw new Error(`${response.status} ${response.statusText}: ${text}`);
  }

  const contentType = response.headers.get("content-type");
  return (contentType?.includes("application/json") ? await response.json() : await response.text()) as T;
}
