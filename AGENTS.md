# Agent Workflow Guidelines

These instructions apply to the entire repository.

## Code style
- Prefer modern JavaScript/Apps Script syntax (const/let over var, arrow functions when appropriate).
- Keep functions small and single-purpose; extract helpers in the same file when logic grows.
- Avoid adding external services or dependencies; keep solutions Apps Script–native unless absolutely necessary.
- Use descriptive names for triggers, sheets, and ranges to make Apps Script automation clear.
- When touching HTML, keep client-side scripts minimal and avoid inline styles if possible.
- Do not wrap import statements in try/catch blocks.

## Documentation & comments
- Add or update short comments for non-obvious logic, especially around Google Sheets ranges or advanced services.
- Keep README or inline usage notes in sync if behavior changes.
- For every `Index.html` and `Code.gs`, prepend a short comment containing a version name/number and the current git commit hash (e.g., `<!-- Version 1.0.0 | abc1234 -->` for HTML, `// Version 1.0.0 | abc1234` for `.gs`). Update this header whenever the file changes.
- When updating these headers, set the hash to the short ID of the current `HEAD` **before** you make changes; this avoids a self-referential hash update cycle and keeps traceability to the prior commit.

## Testing & validation
- Prefer lightweight verification (running scripts locally via clasp or dry-run functions) when feasible; note any manual steps in the PR.
- If no automated checks exist, describe the manual validation performed in the PR/testing notes.

## Git & PR expectations
- Keep commits scoped and descriptive.
- Summarize changes clearly in the PR body and list any validation steps taken.
- If modifying triggers or deployment steps, call them out explicitly.

## Files & organization
- Place shared utilities in the existing top-level `.gs`/`.js` files unless a new file materially improves clarity.
- Keep HTML assets in `Index.html` unless a separate template is needed for readability.

Following these conventions will help keep the Apps Script project maintainable and easy to review.
