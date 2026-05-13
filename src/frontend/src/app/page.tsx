"use client";

import { InteractionStatus } from "@azure/msal-browser";
import { Spinner } from "@fluentui/react-components";
import { useIsAuthenticated, useMsal } from "@azure/msal-react";
import { useRouter } from "next/navigation";
import { useEffect } from "react";

export default function HomePage() {
  const router = useRouter();
  const isAuthenticated = useIsAuthenticated();
  const { inProgress } = useMsal();

  useEffect(() => {
    if (inProgress === InteractionStatus.None) {
      router.replace(isAuthenticated ? "/files" : "/signin");
    }
  }, [inProgress, isAuthenticated, router]);

  return (
    <main className="centered-page" aria-live="polite">
      <Spinner label="Opening Wordembedded" />
    </main>
  );
}
