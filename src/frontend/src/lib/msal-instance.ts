import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig } from "./msal-config";

export const msalInstance = new PublicClientApplication(msalConfig);

let initializationPromise: Promise<void> | undefined;

export function initializeMsal(): Promise<void> {
  initializationPromise ??= msalInstance.initialize().then(async () => {
    const redirectResponse = await msalInstance.handleRedirectPromise({ navigateToLoginRequestUrl: true });

    if (redirectResponse?.account) {
      msalInstance.setActiveAccount(redirectResponse.account);
      return;
    }

    if (!msalInstance.getActiveAccount()) {
      const [firstAccount] = msalInstance.getAllAccounts();
      if (firstAccount) msalInstance.setActiveAccount(firstAccount);
    }
  });

  return initializationPromise;
}
