"use client";

import { InteractionStatus } from "@azure/msal-browser";
import { Button, Card, Spinner, Text } from "@fluentui/react-components";
import { PersonKeyRegular } from "@fluentui/react-icons";
import { useIsAuthenticated, useMsal } from "@azure/msal-react";
import { useRouter } from "next/navigation";
import { useEffect } from "react";
import { ThemeToggle } from "@/components/theme/ThemeToggle";
import { loginScopes } from "@/lib/msal-config";

export default function SignInPage() {
  const router = useRouter();
  const isAuthenticated = useIsAuthenticated();
  const { inProgress, instance } = useMsal();

  useEffect(() => {
    if (isAuthenticated && inProgress === InteractionStatus.None) {
      router.replace("/files");
    }
  }, [inProgress, isAuthenticated, router]);

  const signIn = () => {
    void instance.loginRedirect({ scopes: loginScopes });
  };

  return (
    <main className="centered-page signin-page">
      <Card className="signin-card" appearance="outline">
        <div className="page-toolbar">
          <Text size={200} weight="semibold">
            Register
          </Text>
          <ThemeToggle />
        </div>
        <div className="stack-lg">
          <div className="stack-sm">
            <Text as="h1" className="page-title" weight="semibold">
              Sign in to your document workspace
            </Text>
            <Text className="page-copy">
              Use your Microsoft work account to access SharePoint Embedded containers through the Register API.
            </Text>
          </div>
          <Button
            appearance="primary"
            disabled={inProgress !== InteractionStatus.None}
            icon={inProgress === InteractionStatus.None ? <PersonKeyRegular /> : undefined}
            onClick={signIn}
            size="large"
          >
            {inProgress === InteractionStatus.None ? "Sign in with Microsoft" : "Sign-in in progress"}
          </Button>
          {inProgress !== InteractionStatus.None ? <Spinner size="tiny" label="Waiting for Microsoft sign-in" /> : null}
        </div>
      </Card>
    </main>
  );
}
