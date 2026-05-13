"use client";

import { EventType, type AuthenticationResult, type EventMessage } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { FluentProvider, Spinner, webDarkTheme, webLightTheme } from "@fluentui/react-components";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type { ReactNode } from "react";
import { useEffect, useState } from "react";
import { AppToasterProvider } from "@/components/common/Toaster";
import { ThemeControllerProvider, useAppTheme } from "@/components/theme/ThemeContext";
import { initializeMsal, msalInstance } from "@/lib/msal-instance";

function isAuthenticationResult(payload: EventMessage["payload"]): payload is AuthenticationResult {
  return Boolean(payload && typeof payload === "object" && "account" in payload);
}

function FluentThemeProvider({ children }: { children: ReactNode }) {
  const { theme } = useAppTheme();
  const fluentTheme = theme === "dark" ? webDarkTheme : webLightTheme;

  return (
    <FluentProvider className="fluent-root" theme={fluentTheme}>
      {children}
    </FluentProvider>
  );
}

export function Providers({ children }: { children: ReactNode }) {
  const [queryClient] = useState(
    () =>
      new QueryClient({
        defaultOptions: { queries: { staleTime: 30_000, retry: 1 } },
      }),
  );
  const [isMsalReady, setIsMsalReady] = useState(false);

  useEffect(() => {
    let isMounted = true;

    const callbackId = msalInstance.addEventCallback((event) => {
      if (event.eventType === EventType.LOGIN_SUCCESS && isAuthenticationResult(event.payload)) {
        msalInstance.setActiveAccount(event.payload.account);
      }
    });

    initializeMsal()
      .catch((error: unknown) => {
        console.error("MSAL initialization failed", error);
      })
      .finally(() => {
        if (isMounted) setIsMsalReady(true);
      });

    return () => {
      isMounted = false;
      if (callbackId) msalInstance.removeEventCallback(callbackId);
    };
  }, []);

  return (
    <MsalProvider instance={msalInstance}>
      <QueryClientProvider client={queryClient}>
        <ThemeControllerProvider>
          <FluentThemeProvider>
            <AppToasterProvider>{isMsalReady ? children : <Spinner className="centered-page" label="Preparing sign in" />}</AppToasterProvider>
          </FluentThemeProvider>
        </ThemeControllerProvider>
      </QueryClientProvider>
    </MsalProvider>
  );
}
