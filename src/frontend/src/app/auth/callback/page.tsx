"use client";

import { Spinner, Text } from "@fluentui/react-components";

export default function AuthCallbackPage() {
  return (
    <main className="centered-page" aria-live="polite">
      <div className="stack-sm centered-copy">
        <Spinner label="Signing you in" />
        <Text className="page-copy">Microsoft sign-in is completing. You will be redirected shortly.</Text>
      </div>
    </main>
  );
}
