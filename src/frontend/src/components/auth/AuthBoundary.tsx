"use client";

import { InteractionStatus } from "@azure/msal-browser";
import { Button, Card, Spinner, Text } from "@fluentui/react-components";
import { PersonKeyRegular } from "@fluentui/react-icons";
import { useIsAuthenticated, useMsal } from "@azure/msal-react";
import { useRouter } from "next/navigation";
import type { ReactNode } from "react";
import { useEffect } from "react";
import { loginScopes } from "@/lib/msal-config";

export function AuthBoundary({ children }: { children: ReactNode }) {
  const router = useRouter();
  const isAuthenticated = useIsAuthenticated();
  const { inProgress, instance } = useMsal();

  useEffect(() => {
    if (!isAuthenticated && inProgress === InteractionStatus.None) {
      router.replace("/signin");
    }
  }, [inProgress, isAuthenticated, router]);

  if (inProgress !== InteractionStatus.None) {
    return (
      <main className="centered-page" aria-live="polite">
        <Spinner label="Checking sign-in" />
      </main>
    );
  }

  if (!isAuthenticated) {
    return (
      <main className="centered-page">
        <Card className="signin-card" appearance="outline">
          <div className="stack-sm">
            <Text as="h1" className="page-title" weight="semibold">
              Sign in required
            </Text>
            <Text className="page-copy">Use your Microsoft work account to continue.</Text>
          </div>
          <Button appearance="primary" icon={<PersonKeyRegular />} onClick={() => void instance.loginRedirect({ scopes: loginScopes })}>
            Sign in with Microsoft
          </Button>
        </Card>
      </main>
    );
  }

  return <>{children}</>;
}
