# Confidentiality check

Demo packs get forwarded. For a legal team, one visible client name, matter number or privileged sentence is a real
problem, so this check is a gate, not a suggestion. Run it before anything is copied out of `demo/takes/`.

## What to look for

- Everything on the team's never-on-screen list (question 3 of the intake).
- Names of real people other than the team, unless they agreed to appear.
- Client, counterparty, matter or deal names; matter, contract or ticket numbers.
- Internal URLs, server names, SharePoint paths, email addresses, phone numbers.
- Privileged or personal content in documents, chat history, sidebars and recent-file lists.
- Window titles, tabs and taskbars: a browser tab or document title often leaks what the content area hides.
- Notifications and toasts that appeared mid-take.

## How

1. `gifcap sheet` every take that will be used (`--every 0.5` for short ones) and look at every frame. Also check
   every still.
2. For each problem, find the area in recording pixels (frame size = the take's `rect` width and height; the sheet's
   frames are scaled copies of it) and cover it:
   `gifcap gif take.mkv out.gif --redact 1830,2,190,28` (a solid box; repeatable).
   Prefer `--redact` to `--blur`: a solid box can't be reversed.
3. If something sensitive moves around or fills the frame, don't redact: retake with sample data.
4. Re-sheet the output and confirm the box covers it in every frame.
5. List each redaction (file, area, what it covers, why) in `demo/REVIEW.md`.

If you're unsure whether something is sensitive, cover it and flag it in REVIEW.md for the team to decide.
