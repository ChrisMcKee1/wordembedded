# Product

## Register

product

## Users

Knowledge workers inside a single organization who already live in Microsoft 365. They have Outlook, Teams, and OneDrive open all day; identity, licensing, and permissions are familiar terrain, not something to teach.

Three working roles, from the PRD:

- **Readers** browse and preview documents inline. They want to find the right file and read it without leaving the app.
- **Editors** open Word, Excel, or PowerPoint in Office for the web to co-author in real time. They want one click to get there and one click to come back.
- **Owners and tenant admins** share documents, manage permissions, and apply organizational governance (sensitivity labels, retention, DLP).

The product is used at a desk, in a browser, alongside other M365 surfaces. People won't tolerate a slower or noisier experience than the SharePoint or OneDrive surfaces it replaces — they have those tools one tab over.

## Product Purpose

A custom React workspace that owns the file-collaboration UX (browsing, organizing, sharing, AI-assisted workflows) while delegating storage, identity, permissions, and Office co-authoring to Microsoft 365 via SharePoint Embedded.

Audience scope is **internal LOB for one organization**. Multi-tenant ISV remains a stated future possibility in the PRD, but it is not the design target for this iteration.

Success looks like: people in the org prefer this app over navigating raw SharePoint or OneDrive for the documents that live in its containers, and they don't have to think about *which* Microsoft surface they're in to get co-authoring, AutoSave, version history, presence, and comments — those just work.

## Brand Personality

**Clear. Fast. Grown-up.**

Microsoft-native by deliberate choice — Fluent v9 vocabulary, dual light and dark themes mirroring the Office family — but visibly tighter and quicker than the SharePoint surfaces it sits next to. The tone is adult: no consumer illustrations, no animated mascots, no marketing flourishes inside the workspace. The voice is plain and direct in error states and sharing copy.

The user is fluent in M365. The product greets that fluency instead of explaining it.

## Anti-references

- **Windows 11 Mica / acrylic glassmorphism.** No decorative blur, no glass cards by default, no translucent panes used "for depth." If a surface is blurred, there is a functional reason that survives a code-review challenge.
- Also worth restating from the shared design laws: no side-stripe accent borders, no gradient text, no hero-metric template, no identical card grids, no modal as first thought. These are absolute bans, not preferences.

## Design Principles

1. **Respect M365 muscle memory.** Sharing, permissions, presence, and the editor handoff behave the way someone who already uses M365 expects. Don't reinvent vocabulary or flows that Microsoft has trained users on for fifteen years.
2. **Office owns Office.** Never simulate a Word, Excel, or PowerPoint editor. Preview inline when read-only is enough; hand off to Office for the web in a new window when editing is needed. The handoff is the feature, not a workaround.
3. **The chrome is invisible.** The document, the list, and the action are the foreground. The app's own identity is incidental. Brand expression lives in restraint, not ornament.
4. **Fast beats fancy.** Every navigation, list render, and preview load should feel snappier than the SharePoint surface it replaces. Animations exist to confirm an action, not to entertain.
5. **Make access boundaries legible.** SPE permission layering (container membership → driveItem permissions) is the actual mental model. Show *who can see this* and *how it was shared* in the moment of action, not buried in a settings page.

## Accessibility & Inclusion

- **WCAG 2.2 AA** is the floor. Audit against it before any release.
- **Keyboard-first.** File workflows are list- and grid-heavy; every action reachable by mouse must be reachable by keyboard with a visible focus ring that survives both themes.
- **Reduced motion respected.** `prefers-reduced-motion` collapses non-essential motion to instant state changes. Confirmation animation is not essential; presence updates and AutoSave indicators are.
- **Both themes meet the same contrast bar.** Light and dark are equal citizens — neither is the "less polished" mode. Verify contrast in both with the actual production palette, not the design tokens in isolation.
- **No color-only signal.** Permission state, error state, and presence use color *plus* shape, icon, or text.
