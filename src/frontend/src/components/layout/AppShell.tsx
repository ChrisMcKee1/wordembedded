"use client";

import { Avatar, Button, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, SearchBox, Text, Toolbar, makeStyles, shorthands, tokens } from "@fluentui/react-components";
import { SignOutRegular } from "@fluentui/react-icons";
import { useMsal } from "@azure/msal-react";
import { usePathname, useRouter, useSearchParams } from "next/navigation";
import { type ReactNode, useState } from "react";
import { ThemeToggle } from "@/components/theme/ThemeToggle";

const useStyles = makeStyles({
  shell: {
    minHeight: "100vh",
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
  },
  header: {
    position: "sticky",
    top: 0,
    zIndex: 2,
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderBottom(tokens.strokeWidthThin, "solid", tokens.colorNeutralStroke2),
  },
  toolbar: {
    width: "min(100%, 1280px)",
    marginRight: "auto",
    marginLeft: "auto",
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    paddingRight: tokens.spacingHorizontalL,
    paddingLeft: tokens.spacingHorizontalL,
    display: "grid",
    gridTemplateColumns: "auto minmax(220px, 440px) auto",
    gap: tokens.spacingHorizontalL,
    alignItems: "center",
    "@media (max-width: 720px)": {
      gridTemplateColumns: "1fr",
      alignItems: "stretch",
    },
  },
  brand: {
    display: "grid",
    gap: tokens.spacingVerticalXXS,
  },
  actions: {
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-end",
    gap: tokens.spacingHorizontalS,
    "@media (max-width: 720px)": {
      justifyContent: "space-between",
    },
  },
  main: {
    width: "min(100%, 1280px)",
    marginRight: "auto",
    marginLeft: "auto",
    padding: tokens.spacingHorizontalL,
  },
  userButton: {
    maxWidth: "260px",
  },
});

export function AppShell({ children }: { children: ReactNode }) {
  const styles = useStyles();
  const router = useRouter();
  const pathname = usePathname();
  const searchParams = useSearchParams();
  const initialQuery = pathname === "/search" ? (searchParams.get("q") ?? "") : "";
  const [query, setQuery] = useState(initialQuery);
  const { accounts, instance } = useMsal();
  const account = instance.getActiveAccount() ?? accounts[0];
  const displayName = account?.name ?? account?.username ?? "User";
  const email = account?.username ?? "";

  const submitSearch = () => {
    const value = query.trim();
    if (value) router.push(`/search?q=${encodeURIComponent(value)}`);
  };

  const signOut = () => {
    void instance.logoutRedirect({ account, postLogoutRedirectUri: window.location.origin });
  };

  return (
    <div className={styles.shell}>
      <header className={styles.header}>
        <Toolbar className={styles.toolbar} aria-label="Application toolbar">
          <div className={styles.brand}>
            <Text weight="semibold" size={500}>Wordembedded</Text>
            <Text size={200}>SharePoint Embedded workspace</Text>
          </div>
          <SearchBox
            aria-label="Search documents"
            placeholder="Search documents"
            value={query}
            onChange={(_, data) => setQuery(data.value)}
            onKeyDown={(event) => {
              if (event.key === "Enter") submitSearch();
            }}
          />
          <div className={styles.actions}>
            <ThemeToggle />
            <Menu positioning="below-end">
              <MenuTrigger disableButtonEnhancement>
                <Button className={styles.userButton} appearance="subtle" icon={<Avatar name={displayName} size={28} />}>
                  {displayName}
                </Button>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem disabled>{email || displayName}</MenuItem>
                  <MenuItem icon={<SignOutRegular />} onClick={signOut}>Sign out</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </div>
        </Toolbar>
      </header>
      <main className={styles.main}>{children}</main>
    </div>
  );
}
