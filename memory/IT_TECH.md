# IT Tech Memory

Focus: scripts, automation, integrations, hosting.

Update rules
- Save technical decisions here (with date).
- Keep secrets out of repo (use env vars or local files).

TBD: add setup notes after athlete onboarding.

2026-02-09
- Report UI: improved chart legibility and explainability.
- Charts now render inside `chart-card` blocks with:
  - fixed height for readability (responsive),
  - consistent axis/legend/tooltip styling for dark background,
  - multi-axis wellness charts (sleep on left; HRV/FC on right) to avoid flat/illegible lines,
  - an expandable "O que este grafico significa" section per chart,
  - an "Ampliar" modal to view the chart larger (also opens on canvas click).
