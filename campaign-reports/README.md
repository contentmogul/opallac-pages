# Campaign reports

Dedicated, single-campaign report pages — one HTML file per report, no shared
nav/sidebar, meant to be sent directly to a client/brand.

This is distinct from the root-level tools (`drive-assets.html`,
`campaign-plans.html`), which are internal, multi-campaign browsing UIs gated
by each client's live, client-approval-facing registry Sheet. A report here
draws its authorized Drive folder from a client's separate
`historicalRegistrySheetId` registry instead, so a past/closed campaign never
shows up in the approval-facing list.
